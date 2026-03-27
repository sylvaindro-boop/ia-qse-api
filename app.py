from fastapi import FastAPI
import requests
import os

app = FastAPI()

TENANT_ID = "f4eee41d-b0b9-456d-a1f5-b4d044bf0b8f"
CLIENT_ID = "e81d7039-c7f2-4fd2-af4d-da252299be5b"
CLIENT_SECRET = "nEH8Q~eREPIlWz2wrN.YZFennIw2efl8qFFreau1"

SITE_ID = "397a5d9b-d3ee-41f8-b105-5049851c80a1"
LIST_ID = "862e787a-2ef4-4cb9-8fac-9b7c64afeeeb"

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


@app.get("/")
def root():
    return {"status": "ok", "message": "API vivante"}


def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    token_response = requests.post(token_url, data=token_data, timeout=30)
    token_json = token_response.json()

    if "access_token" not in token_json:
        return None, token_json

    return token_json["access_token"], None


@app.get("/analyse")
def analyse(constat: str):
    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"resultat": "ERREUR AZURE TOKEN : " + str(token_error)}

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items?expand=fields"
        sp_headers = {"Authorization": f"Bearer {access_token}"}

        sp_response = requests.get(sp_url, headers=sp_headers, timeout=30)
        sp_json = sp_response.json()

        if "value" not in sp_json:
            return {"resultat": "ERREUR SHAREPOINT : " + str(sp_json)}

        memoire = ""
        count = 0

        for item in sp_json.get("value", []):
            fields = item.get("fields", {})
            title = fields.get("Title", "")

            if "déchets" in title.lower():
                memoire += f"""
CAS REEL :
Constat : {fields.get("Title", "")}
Cause : {fields.get("Cause_Finale", "")}
Action : {fields.get("ActionCorrective_Finale", "")}
---
"""
                count += 1

            if count >= 3:
                break

        prompt = f"""
Tu es un expert QSE terrain.

Tu dois répondre EXACTEMENT avec ces 6 lignes, sans texte avant ni après :

CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=

Base-toi sur les cas réels suivants :
{memoire}

Nouveau constat :
{constat}

Règles :
- Cause racine courte
- Actions concrètes terrain
- Pas de blabla
"""

        ai_url = "https://api.openai.com/v1/chat/completions"
        ai_headers = {
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json"
        }
        ai_data = {
            "model": "gpt-4.1-mini",
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.3
        }

        ai_response = requests.post(ai_url, headers=ai_headers, json=ai_data, timeout=60)
        ai_json = ai_response.json()

        if "choices" not in ai_json:
            return {"resultat": "ERREUR OPENAI : " + str(ai_json)}

        texte = ai_json["choices"][0]["message"]["content"]
        return {"resultat": texte}

    except Exception as e:
        return {"resultat": "ERREUR PYTHON : " + str(e)}


@app.get("/memoire_test")
def memoire_test():
    data_input = {
        "Constat": "Test déchets chantier",
        "Client": "Test client",
        "Source": "Audit chantier",
        "Activite": "Déchets",
        "ActionImmediate": "Nettoyage immédiat",
        "Cause": "Rigueur",
        "ActionCorrective": "Rappel des consignes",
        "MesureEfficacite": "Contrôle terrain"
    }

    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        fields = {
            "Title": data_input.get("Constat"),
            "Client": data_input.get("Client"),
            "Source": data_input.get("Source"),
            "Activite": data_input.get("Activite"),
            "ActionImmediate_Finale": data_input.get("ActionImmediate"),
            "Cause_Finale": data_input.get("Cause"),
            "ActionCorrective_Finale": data_input.get("ActionCorrective"),
            "MesureEfficacite_Finale": data_input.get("MesureEfficacite")
        }

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
        sp_headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        payload = {"fields": fields}

        sp_response = requests.post(sp_url, headers=sp_headers, json=payload, timeout=30)

        return {
            "status": "ok",
            "http_status": sp_response.status_code,
            "detail": sp_response.json()
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}
