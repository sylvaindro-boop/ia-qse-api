from fastapi import FastAPI
import requests

app = FastAPI()

# ===== INFOS AZURE =====
TENANT_ID = "f4eee41d-b0b9-456d-a1f5-b4d044bf0b8f"
CLIENT_ID = "e81d7039-c7f2-4fd2-af4d-da252299be5b"
CLIENT_SECRET = "nEH8Q~eREPIlWz2wrN.YZFennIw2efl8qFFreau1"

SITE_ID = "397a5d9b-d3ee-41f8-b105-5049851c80a1"
LIST_ID = "862e787a-2ef4-4cb9-8fac-9b7c64afeeeb"

# ===== OPENAI =====
OPENAI_API_KEY = "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"

@app.get("/analyse")
def analyse(constat: str):

    # ===== AUTHENTIFICATION AZURE =====
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

    for item in data.get("value", []):
        fields = item.get("fields", {})

        if "déchets" in fields.get("Title", "").lower():
            memoire += f"""
CAS REEL :
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

Tu dois répondre SIMPLE, CONCRET, PRATIQUE.

Règles :
- Cause racine = courte (pas de phrase)
- Actions terrain réelles
- Pas de blabla

EXEMPLES :
{memoire}

NOUVEAU CAS :
Constat : {constat}

Réponds EXACTEMENT comme ça :

CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=
"""

    # ===== OPENAI =====
    url_ai = "https://api.openai.com/v1/chat/completions"

    headers_ai = {
        "Authorization": f"Bearer {OPENAI_API_KEY}",
        "Content-Type": "application/json"
    }

    data_ai = {
        "model": "gpt-4.1-mini",
        "messages": [
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.3
    }

response_ai = requests.post(url_ai, headers=headers_ai, json=data_ai)
result = response_ai.json()

print(result)  # 👈 IMPORTANT pour voir l'erreur dans logs

if "choices" in result:
    texte = result["choices"][0]["message"]["content"]
else:
    texte = "ERREUR OPENAI : " + str(result)

    return {"resultat": texte}
