import io
import json
import os
import re
import zipfile
from collections import Counter
from datetime import datetime, timezone
from xml.sax.saxutils import escape

import requests


def _token():
    tenant = os.environ["TENANT_ID"]
    response = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": os.environ["CLIENT_ID"],
            "client_secret": os.environ["CLIENT_SECRET"],
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def _get_bytes(url, token):
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    response.raise_for_status()
    return response.content


def _load_mapping(token):
    drive = os.environ["RECONCILE_PAYLOAD_DRIVE_ID"]
    item = os.environ["RECONCILE_PAYLOAD_ITEM_ID"]
    raw = _get_bytes(f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}/content", token)
    payload = json.loads(raw.decode("utf-8"))
    actions = payload.get("actions", [])
    summary = payload.get("reconciliation", {})
    if len(actions) != 163 or summary.get("after_count") != 163:
        raise RuntimeError("reconciliation payload is not finalized at 163 actions")
    mapping = {}
    for action in actions:
        row = int(action["row"])
        item_id = str(action.get("SharePointID", "")).strip()
        if not item_id:
            raise RuntimeError(f"missing SharePointID for row {row}")
        mapping[row] = item_id
    if set(mapping) != set(range(2, 165)):
        raise RuntimeError("Excel row mapping is not exactly rows 2..164")
    return mapping


def _find_plan_sheet_path(zf):
    wb = zf.read("xl/workbook.xml").decode("utf-8")
    match = re.search(
        r'<sheet\b[^>]*\bname="PLAN_D_ACTION_V2"[^>]*\br:id="([^"]+)"[^>]*/?>', wb
    )
    if not match:
        match = re.search(
            r'<sheet\b[^>]*\br:id="([^"]+)"[^>]*\bname="PLAN_D_ACTION_V2"[^>]*/?>', wb
        )
    if not match:
        raise RuntimeError("PLAN_D_ACTION_V2 not found in workbook.xml")
    rid = match.group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel_match = re.search(
        rf'<Relationship\b[^>]*\bId="{re.escape(rid)}"[^>]*\bTarget="([^"]+)"[^>]*/?>', rels
    )
    if not rel_match:
        rel_match = re.search(
            rf'<Relationship\b[^>]*\bTarget="([^"]+)"[^>]*\bId="{re.escape(rid)}"[^>]*/?>', rels
        )
    if not rel_match:
        raise RuntimeError("worksheet relationship not found")
    target = rel_match.group(1).lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target


def _cell_pattern(ref):
    return re.compile(
        rf'<c\b[^>]*\br="{re.escape(ref)}"[^>]*(?:/>|>.*?</c>)',
        re.DOTALL,
    )


def _replace_or_append_cell(row_xml, ref, cell_xml):
    pattern = _cell_pattern(ref)
    if pattern.search(row_xml):
        return pattern.sub(cell_xml, row_xml, count=1)
    pos = row_xml.rfind("</row>")
    if pos < 0:
        raise RuntimeError(f"invalid row XML for {ref}")
    return row_xml[:pos] + cell_xml + row_xml[pos:]


def _patch_sheet(sheet_xml, mapping):
    # Reuse the workbook's established style for SharePoint IDs when available.
    styles = re.findall(r'<c\b[^>]*\br="R\d+"[^>]*\bs="(\d+)"', sheet_xml)
    r_style = Counter(styles).most_common(1)[0][0] if styles else None
    sync_text = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    for row_num in range(2, 165):
        row_pattern = re.compile(
            rf'<row\b(?=[^>]*\br="{row_num}")[^>]*>.*?</row>', re.DOTALL
        )
        row_match = row_pattern.search(sheet_xml)
        if not row_match:
            raise RuntimeError(f"worksheet row {row_num} not found")
        row_xml = row_match.group(0)
        item_id = escape(str(mapping[row_num]))
        style_attr = f' s="{r_style}"' if r_style else ""
        r_cell = f'<c r="R{row_num}"{style_attr}><v>{item_id}</v></c>'
        s_cell = (
            f'<c r="S{row_num}" t="inlineStr"><is><t>{escape(sync_text)}</t></is></c>'
        )
        row_xml = _replace_or_append_cell(row_xml, f"R{row_num}", r_cell)
        row_xml = _replace_or_append_cell(row_xml, f"S{row_num}", s_cell)
        sheet_xml = sheet_xml[:row_match.start()] + row_xml + sheet_xml[row_match.end():]

    # Guard against the old orphan-ID class of bug outside the authoritative range.
    for row_num in range(165, 501):
        for col in ("R", "S"):
            sheet_xml = _cell_pattern(f"{col}{row_num}").sub("", sheet_xml)
    return sheet_xml


def _patch_workbook_bytes(raw, mapping):
    src = io.BytesIO(raw)
    dst = io.BytesIO()
    with zipfile.ZipFile(src, "r") as zin:
        sheet_path = _find_plan_sheet_path(zin)
        names = zin.namelist()
        if "xl/vbaProject.bin" not in names:
            raise RuntimeError("vbaProject.bin missing before patch")
        original_vba = zin.read("xl/vbaProject.bin")
        sheet_xml = zin.read(sheet_path).decode("utf-8")
        patched_sheet = _patch_sheet(sheet_xml, mapping).encode("utf-8")
        with zipfile.ZipFile(dst, "w") as zout:
            for info in zin.infolist():
                data = patched_sheet if info.filename == sheet_path else zin.read(info.filename)
                zout.writestr(info, data)
    patched = dst.getvalue()
    with zipfile.ZipFile(io.BytesIO(patched), "r") as verify:
        if verify.read("xl/vbaProject.bin") != original_vba:
            raise RuntimeError("VBA project changed during workbook patch")
        final_sheet = verify.read(sheet_path).decode("utf-8")
        for row_num, item_id in mapping.items():
            if not re.search(rf'<c\b[^>]*\br="R{row_num}"[^>]*>\s*<v>{re.escape(item_id)}</v>\s*</c>', final_sheet):
                raise RuntimeError(f"ID verification failed for row {row_num}")
    return patched


def write_ids_once():
    if os.getenv("WRITE_IDS_RUN", "") != "1":
        return {"status": "skipped"}
    token = _token()
    mapping = _load_mapping(token)
    drive = os.environ["WORKBOOK_DRIVE_ID"]
    item = os.environ["WORKBOOK_ITEM_ID"]
    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}/content"
    raw = _get_bytes(content_url, token)
    patched = _patch_workbook_bytes(raw, mapping)
    response = requests.put(
        content_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.ms-excel.sheet.macroEnabled.12",
        },
        data=patched,
        timeout=90,
    )
    response.raise_for_status()
    result = {
        "status": "ok",
        "rows": len(mapping),
        "bytes_before": len(raw),
        "bytes_after": len(patched),
        "item_id": item,
    }
    print("WORKBOOK_IDS_SUMMARY " + json.dumps(result, ensure_ascii=False), flush=True)
    return result


if __name__ == "__main__":
    print(json.dumps(write_ids_once(), ensure_ascii=False))
