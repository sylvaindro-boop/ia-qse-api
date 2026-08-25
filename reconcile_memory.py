import json
import os
import re
import unicodedata
from collections import defaultdict

import requests


def _norm(value):
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = re.sub(r"\s+", " ", text)
    return text


def _date_key(value):
    text = str(value or "").strip()
    return text[:10] if len(text) >= 10 else text


def _slug(value):
    text = _norm(value)
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return text[:120]


def _item_constat(item):
    fields = item.get("fields", {})
    return fields.get("ConstatComplet") or fields.get("Title", "")


def _excel_fields(action, source_name):
    constat = str(action.get("Constat", "") or "")
    tags = [
        action.get("Source", ""),
        action.get("Activite", ""),
        action.get("Typologie", ""),
        "excel-master",
        f"action-necessaire-{action.get('ActionCorrectiveNecessaire', '')}",
        f"resp-{action.get('Responsable', '')}",
        f"delai-{action.get('Delai', '')}",
        f"statut-{action.get('Statut', '')}",
        f"redacteur-{action.get('Redacteur', '')}",
        f"date-efficacite-{action.get('DateEfficacite', '')}",
    ]
    if action.get("Commentaire"):
        tags.append("avec-commentaire")
    tags = ";".join(x for x in (_slug(v) for v in tags) if x)
    return {
        "Title": constat[:255],
        "ConstatComplet": constat,
        "Source": action.get("Source", ""),
        "Activite": action.get("Activite", ""),
        "DateCas": _date_key(action.get("DateCas", "")),
        "ActionImmediate_IA": action.get("ActionImmediate", ""),
        "ActionImmediate_Finale": action.get("ActionImmediate", ""),
        "Analyse_IA": action.get("Analyse", ""),
        "Analyse_Finale": action.get("Analyse", ""),
        "Typologie_IA": action.get("Typologie", ""),
        "Typologie_Finale": action.get("Typologie", ""),
        "ActionCorrective_IA": action.get("ActionCorrective", ""),
        "ActionCorrective_Finale": action.get("ActionCorrective", ""),
        "MesureEfficacite_IA": action.get("MesureEfficacite", ""),
        "MesureEfficacite_Finale": action.get("MesureEfficacite", ""),
        "ModifieParHumain": True,
        "QualiteCas": "Bon",
        "Tags": tags,
        "NomFichierSource": source_name,
    }


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


def _request(method, url, token, **kwargs):
    response = requests.request(
        method,
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=45,
        **kwargs,
    )
    if response.status_code < 200 or response.status_code >= 300:
        raise RuntimeError(f"Graph {method} {response.status_code}: {response.text[:1500]}")
    if not response.content:
        return {}
    try:
        return response.json()
    except Exception:
        return {}


def _load_payload(token):
    drive = os.environ["RECONCILE_PAYLOAD_DRIVE_ID"]
    item = os.environ["RECONCILE_PAYLOAD_ITEM_ID"]
    url = f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}/content"
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=45)
    response.raise_for_status()
    payload = response.json()
    if payload.get("count") != len(payload.get("actions", [])):
        raise RuntimeError("payload count mismatch")
    return payload


def _save_payload(token, payload):
    drive = os.environ["RECONCILE_PAYLOAD_DRIVE_ID"]
    item = os.environ["RECONCILE_PAYLOAD_ITEM_ID"]
    url = f"https://graph.microsoft.com/v1.0/drives/{drive}/items/{item}/content"
    body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    response = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"},
        data=body,
        timeout=45,
    )
    response.raise_for_status()


def _load_items(token):
    site = os.environ["SITE_ID"]
    list_id = os.environ["LIST_ID"]
    url = f"https://graph.microsoft.com/v1.0/sites/{site}/lists/{list_id}/items?$expand=fields&$top=200"
    result = []
    while url:
        data = _request("GET", url, token)
        result.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return result


def _ensure_full_constat_column(token):
    site = os.environ["SITE_ID"]
    list_id = os.environ["LIST_ID"]
    columns_url = f"https://graph.microsoft.com/v1.0/sites/{site}/lists/{list_id}/columns"
    columns = _request("GET", columns_url, token).get("value", [])
    for col in columns:
        if col.get("name") == "ConstatComplet" or col.get("displayName") == "ConstatComplet":
            return col.get("name") or "ConstatComplet"
    definition = {
        "name": "ConstatComplet",
        "displayName": "ConstatComplet",
        "description": "Constat intégral issu du plan d'action Excel maître.",
        "text": {
            "allowMultipleLines": True,
            "appendChangesToExistingText": False,
            "linesForEditing": 10,
            "textType": "plain",
        },
    }
    created = _request("POST", columns_url, token, json=definition)
    internal_name = created.get("name") or "ConstatComplet"
    print("RECON_SCHEMA " + json.dumps({"created": True, "column": internal_name}, ensure_ascii=False), flush=True)
    return internal_name


def _composite_excel(action):
    return (
        _norm(action.get("Constat")),
        _norm(action.get("Source")),
        _norm(action.get("Activite")),
        _date_key(action.get("DateCas")),
    )


def _composite_item(item):
    f = item.get("fields", {})
    return (
        _norm(_item_constat(item)),
        _norm(f.get("Source")),
        _norm(f.get("Activite")),
        _date_key(f.get("DateCas")),
    )


def _legacy_prefix_match(action, item):
    f = item.get("fields", {})
    excel_constat = _norm(action.get("Constat"))
    item_title = _norm(f.get("Title"))
    if len(item_title) < 180 or not excel_constat.startswith(item_title):
        return False
    if _norm(f.get("Source")) != _norm(action.get("Source")):
        return False
    if _norm(f.get("Activite")) != _norm(action.get("Activite")):
        return False
    return _date_key(f.get("DateCas")) == _date_key(action.get("DateCas"))


def reconcile_once():
    if os.getenv("RECONCILE_RUN", "") != "1":
        return {"status": "skipped"}

    token = _token()
    full_constat_column = _ensure_full_constat_column(token)
    if full_constat_column != "ConstatComplet":
        raise RuntimeError(f"unexpected full constat internal column name: {full_constat_column}")

    payload = _load_payload(token)
    actions = payload["actions"]
    if len(actions) != 163:
        raise RuntimeError(f"unexpected Excel action count: {len(actions)}")
    items = _load_items(token)

    by_id = {str(item.get("id", "")): item for item in items}
    by_composite = defaultdict(list)
    by_title = defaultdict(list)
    for item in items:
        item_id = str(item.get("id", ""))
        by_composite[_composite_item(item)].append(item_id)
        by_title[_norm(_item_constat(item))].append(item_id)

    assigned = {}
    used_ids = set()

    # 1. Existing Excel IDs, only if the list item still represents the same action.
    for action in actions:
        row = int(action["row"])
        old_id = str(action.get("SharePointID", "")).strip()
        item = by_id.get(old_id)
        if item:
            item_constat = _norm(_item_constat(item))
            excel_constat = _norm(action.get("Constat"))
            if item_constat == excel_constat or _legacy_prefix_match(action, item):
                assigned[row] = old_id
                used_ids.add(old_id)

    # 2. Exact composite match: constat + source + activity + date.
    for action in actions:
        row = int(action["row"])
        if row in assigned:
            continue
        candidates = [x for x in by_composite.get(_composite_excel(action), []) if x not in used_ids]
        if candidates:
            chosen = sorted(candidates, key=lambda x: int(x))[0]
            assigned[row] = chosen
            used_ids.add(chosen)

    # 3. Exact normalized constat match.
    for action in actions:
        row = int(action["row"])
        if row in assigned:
            continue
        candidates = [x for x in by_title.get(_norm(action.get("Constat")), []) if x not in used_ids]
        if candidates:
            scored = []
            for item_id in candidates:
                f = by_id[item_id].get("fields", {})
                score = 0
                if _norm(f.get("Source")) == _norm(action.get("Source")):
                    score += 2
                if _norm(f.get("Activite")) == _norm(action.get("Activite")):
                    score += 2
                if _date_key(f.get("DateCas")) == _date_key(action.get("DateCas")):
                    score += 1
                scored.append((-score, int(item_id), item_id))
            scored.sort()
            chosen = scored[0][2]
            assigned[row] = chosen
            used_ids.add(chosen)

    # 4. Legacy SharePoint Title may have been truncated at 255 characters.
    for action in actions:
        row = int(action["row"])
        if row in assigned:
            continue
        candidates = []
        for item_id, item in by_id.items():
            if item_id not in used_ids and _legacy_prefix_match(action, item):
                candidates.append(item_id)
        if candidates:
            chosen = sorted(candidates, key=lambda x: int(x))[0]
            assigned[row] = chosen
            used_ids.add(chosen)

    site = os.environ["SITE_ID"]
    list_id = os.environ["LIST_ID"]
    base = f"https://graph.microsoft.com/v1.0/sites/{site}/lists/{list_id}/items"
    source_name = payload.get("source") or "Portail_Battaglino-Déconstruction_V10.xlsm"

    created = 0
    updated = 0
    for action in actions:
        row = int(action["row"])
        fields = _excel_fields(action, source_name)
        item_id = assigned.get(row)
        if item_id:
            _request("PATCH", f"{base}/{item_id}/fields", token, json=fields)
            updated += 1
        else:
            detail = _request("POST", base, token, json={"fields": fields})
            item_id = str(detail.get("id", "")).strip()
            if not item_id:
                raise RuntimeError(f"creation without id for Excel row {row}")
            assigned[row] = item_id
            used_ids.add(item_id)
            created += 1
        print("RECON_MAP " + json.dumps({"row": row, "id": item_id}, ensure_ascii=False), flush=True)

    # Excel is the master: every pre-existing list item not assigned to an Excel row is removed.
    keep_ids = set(assigned.values())
    deleted = 0
    deleted_ids = []
    for item in items:
        item_id = str(item.get("id", ""))
        if item_id and item_id not in keep_ids:
            _request("DELETE", f"{base}/{item_id}", token)
            deleted += 1
            deleted_ids.append(item_id)

    final_items = _load_items(token)
    final_ids = {str(item.get("id", "")) for item in final_items}
    if len(final_items) != len(actions):
        raise RuntimeError(f"final list count {len(final_items)} != Excel count {len(actions)}")
    missing_ids = [item_id for item_id in assigned.values() if item_id not in final_ids]
    if missing_ids:
        raise RuntimeError(f"assigned IDs missing after reconciliation: {missing_ids[:10]}")

    for action in actions:
        action["SharePointID"] = assigned[int(action["row"])]
    result = {
        "status": "ok",
        "excel_count": len(actions),
        "before_count": len(items),
        "after_count": len(final_items),
        "updated": updated,
        "created": created,
        "deleted": deleted,
        "deleted_ids": deleted_ids,
    }
    payload["reconciliation"] = result
    _save_payload(token, payload)
    print("RECON_SUMMARY " + json.dumps(result, ensure_ascii=False), flush=True)
    return result


if __name__ == "__main__":
    print(json.dumps(reconcile_once(), ensure_ascii=False))
