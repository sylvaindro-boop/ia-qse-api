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
    rows = []
    while url:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()
        for item in data.get("value", []):
            f = item.get("fields", {})
            row = {
                "id": str(item.get("id", "")),
                "Title": f.get("Title", ""),
                "Source": f.get("Source", ""),
                "Activite": f.get("Activite", ""),
                "DateCas": f.get("DateCas", ""),
                "ActionImmediate_Finale": f.get("ActionImmediate_Finale", ""),
                "Analyse_Finale": f.get("Analyse_Finale", ""),
                "Typologie_Finale": f.get("Typologie_Finale", ""),
                "ActionCorrective_Finale": f.get("ActionCorrective_Finale", ""),
                "MesureEfficacite_Finale": f.get("MesureEfficacite_Finale", ""),
                "NomFichierSource": f.get("NomFichierSource", ""),
                "Tags": f.get("Tags", ""),
            }
            rows.append(row)
            print("MEMORY_ITEM " + json.dumps(row, ensure_ascii=False))
        url = data.get("@odata.nextLink")
    return {"status": "ok", "count": len(rows)}


if __name__ == "__main__":
    print(json.dumps(snapshot_memory(), ensure_ascii=False))
