from fastapi import FastAPI
from pydantic import BaseModel
import requests
import os

app = FastAPI()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")
LIST_ID = os.getenv("LIST_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


class MemoirePayload(BaseModel):
    Constat: str = ""
    Client: str = ""
    Source: str = ""
    Activite: str = ""
    DateCas: str = ""

    ActionImmediate_IA: str = ""
    ActionImmediate_Finale: str = ""

    Analyse_IA: str = ""
    Analyse_Finale: str = ""

    Cause_IA: str = ""
    Cause_Finale: str = ""
    Typologie_Finale: str = ""

    ActionCorrective_IA: str = ""
    ActionCorrective_Finale: str = ""

    MesureEfficacite_IA: str = ""
    MesureEfficacite_Finale: str = ""

    ModifieParHumain: str = ""
    QualiteCas: str = ""
    Tags: str = ""
    NomFichierSource: str = ""


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


@app.post("/memoire")
def memoire(data: MemoirePayload):
    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        fields = {
            "Title": data.Constat,
            "Client": data.Client,
            "Source": data.Source,
            "Activite": data.Activite,
            "DateCas": data.DateCas,

            "ActionImmediate_IA": data.ActionImmediate_IA,
            "ActionImmediate_Finale": data.ActionImmediate_Finale,

            "Analyse_IA": data.Analyse_IA,
            "Analyse_Finale": data.Analyse_Finale,

            "Cause_IA": data.Cause_IA,
            "Cause_Finale": data.Cause_Finale,
            "Typologie_Finale": data.Typologie_Finale,

            "ActionCorrective_IA": data.ActionCorrective_IA,
            "ActionCorrective_Finale": data.ActionCorrective_Finale,

            "MesureEfficacite_IA": data.MesureEfficacite_IA,
            "MesureEfficacite_Finale": data.MesureEfficacite_Finale,

            "ModifieParHumain": data.ModifieParHumain,
            "QualiteCas": data.QualiteCas,
            "Tags": data.Tags,
            "NomFichierSource": data.NomFichierSource
        }

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
        sp_headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        payload = {"fields": fields}

        sp_response = requests.post(sp_url, headers=sp_headers, json=payload, timeout=30)

        try:
            detail = sp_response.json()
        except Exception:
            detail = sp_response.text

        if sp_response.status_code not in [200, 201]:
            return {
                "status": "error",
                "step": "sharepoint_create",
                "http_status": sp_response.status_code,
                "detail": detail
            }

        return {
            "status": "ok",
            "http_status": sp_response.status_code,
            "detail": detail
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}


@app.get("/memoire_test")
def memoire_test():
    data_input = {
        "Constat": "Test déchets chantier",
        "Client": "Test client",
        "Source": "Audit chantier",
        "Activite": "Déchets",
        "DateCas": "2026-03-28",

        "ActionImmediate_IA": "Nettoyage immédiat",
        "ActionImmediate_Finale": "Nettoyage immédiat",

        "Analyse_IA": "Tri non respecté",
        "Analyse_Finale": "Tri non respecté",

        "Cause_IA": "Rigueur",
        "Cause_Finale": "Rigueur",
        "Typologie_Finale": "Rigueur",

        "ActionCorrective_IA": "Rappel des consignes",
        "ActionCorrective_Finale": "Rappel des consignes",

        "MesureEfficacite_IA": "Contrôle terrain",
        "MesureEfficacite_Finale": "Contrôle terrain",

        "ModifieParHumain": "Non",
        "QualiteCas": "Moyen",
        "Tags": "dechets;chantier",
        "NomFichierSource": "Portail_Battaglino-Déconstruction_V5.xlsm"
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
            "DateCas": data_input.get("DateCas"),

            "ActionImmediate_IA": data_input.get("ActionImmediate_IA"),
            "ActionImmediate_Finale": data_input.get("ActionImmediate_Finale"),

            "Analyse_IA": data_input.get("Analyse_IA"),
            "Analyse_Finale": data_input.get("Analyse_Finale"),

            "Cause_IA": data_input.get("Cause_IA"),
            "Cause_Finale": data_input.get("Cause_Finale"),
            "Typologie_Finale": data_input.get("Typologie_Finale"),

            "ActionCorrective_IA": data_input.get("ActionCorrective_IA"),
            "ActionCorrective_Finale": data_input.get("ActionCorrective_Finale"),

            "MesureEfficacite_IA": data_input.get("MesureEfficacite_IA"),
            "MesureEfficacite_Finale": data_input.get("MesureEfficacite_Finale"),

            "ModifieParHumain": data_input.get("ModifieParHumain"),
            "QualiteCas": data_input.get("QualiteCas"),
            "Tags": data_input.get("Tags"),
            "NomFichierSource": data_input.get("NomFichierSource")
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
