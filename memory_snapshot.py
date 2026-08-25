import json
import os
import requests


def snapshot_memory():
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_id = os.getenv("SITE_ID")
    list_id = os.getenv("LIST_ID")
    if not all([tenant_id, client_id, client_secret, site_id, list_id]):
        return {"status": "error", "detail": "missing configuration"}
    token_response = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    token_response.raise_for_status()
    token = token_response.json()["access_token"]
    headers = {"Authorization": f"Bearer {token}"}

    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?$expand=fields&$top=200"
    count = 0
    while url:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        count += len(data.get("value", []))
        url = data.get("@odata.nextLink")

    result = {"status": "ok", "count": count}
    if os.getenv("RECONCILE_RUN", "") == "1":
        from reconcile_memory import reconcile_once
        result["reconciliation"] = reconcile_once()
    if os.getenv("WRITE_IDS_RUN", "") == "1":
        from write_ids_workbook import write_ids_once
        result["workbook_ids"] = write_ids_once()
    return result


if __name__ == "__main__":
    print(json.dumps(snapshot_memory(), ensure_ascii=False))
