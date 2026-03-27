from fastapi import FastAPI
import requests

app = FastAPI()

# ===== INFOS AZURE =====
TENANT_ID = "f4eee41d-b0b9-456d-a1f5-b4d044bf0b8f"
CLIENT_ID = "e81d7039-c7f2-4fd2-af4d-da252299be5b"
CLIENT_SECRET = "nEH8Q~eREPIlWz2wrN.YZFennIw2efl8qFFreau1"

SITE_ID = "397a5d9b-d3ee-41f8-b105-5049851c80a1"
LIST_ID = "862e787a-2ef4-4cb9-8fac-9b7c64afeeeb"

OPENAI_API_KEY = "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"

@app.get("/analyse")
def analyse(constat: str):

    # ===== AUTHENTIFICATION =====
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    token_response = requests.post(token_url, data=token_data)
    access_token = token_response.json().get("access_token")

    # ===== SHAREPOINT =====
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items?expand=fields"

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    data = response.json()

    # ===== MEMOIRE =====
    memoire = ""
    count = 0

    for item in data["value"]:
        fields = item["fields"]

        title = fields.get("Title", "").lower()
        cause = fields.get("Cause_Finale")

        if "déchets" in title and cause:
            memoire += f"""
Constat : {fields.get("Title")}
Cause : {fields.get("Cause_Finale")}
Action : {fields.get("ActionCorrective_Finale")}
---
"""
            count += 1

        if count >= 3:
            break

    # ===== PROMPT =====
    prompt = f"""
Tu es un expert QSE terrain.

Base-toi sur ces cas réels :
{memoire}

Nouveau constat :
{constat}

Donne :
1. Cause racine terrain
2. Action corrective concrète
"""

    # ===== OPENAI =====
    url_ai = "https://api.openai.com/v1/responses"

    headers_ai = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }

    data_ai = {
        "model": "gpt-4.1-mini",
        "input": prompt
    }

    response_ai = requests.post(url_ai, headers=headers_ai, json=data_ai)
    result = response_ai.json()

    texte = result["output"][0]["content"][0]["text"]

    return {"resultat": texte}
