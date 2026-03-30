from fastapi import FastAPI
from pydantic import BaseModel
import requests
import os
import re
import unicodedata
from urllib.parse import unquote

app = FastAPI()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")
LIST_ID = os.getenv("LIST_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")


class MemoirePayload(BaseModel):
    Constat: str = ""
    Source: str = ""
    Activite: str = ""
    DateCas: str = ""

    ActionImmediate_IA: str = ""
    ActionImmediate_Finale: str = ""

    Analyse_IA: str = ""
    Analyse_Finale: str = ""

    Typologie_IA: str = ""
    Typologie_Finale: str = ""

    ActionCorrective_IA: str = ""
    ActionCorrective_Finale: str = ""

    MesureEfficacite_IA: str = ""
    MesureEfficacite_Finale: str = ""

    ModifieParHumain: str = ""
    QualiteCas: str = ""
    Tags: str = ""
    NomFichierSource: str = ""


class MemoireUpdatePayload(BaseModel):
    SharePointID: str = ""
    ActionImmediate_Finale: str = ""
    Analyse_Finale: str = ""
    Typologie_Finale: str = ""
    ActionCorrective_Finale: str = ""
    MesureEfficacite_Finale: str = ""


@app.get("/")
def root():
    return {"status": "ok", "message": "API vivante"}


# =========================
# (tout ton code existant inchangé)
# =========================

# 👉 JE NE MODIFIE QUE CETTE PARTIE
@app.get("/analyse")
def analyse(constat: str, source: str = "", activite: str = "", typologies: str = ""):
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

        items = sp_json.get("value", [])
        memoire = build_memory_context(constat, items, limit=5)

        typologies_list = ""
        typologies_brutes = []

        if typologies:
            typologies = unquote(typologies)
            lignes = typologies.split("||")

            for l in lignes:
                l = l.strip()
                if l != "":
                    l = l.lstrip("-").strip()
                    if l != "":
                        typologies_brutes.append(l)

        if typologies_brutes:
            typologies_list = "\n".join([f"- {x}" for x in typologies_brutes])

        # 🔥 PROMPT OPTIMISÉ
        prompt = f"""
Tu es un responsable QSE chantier expérimenté.

Tu dois produire un plan d'action concret, utile et directement applicable.

Réponds uniquement avec ces 6 lignes :
CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=

--- NOUVEAU CAS ---
Source : {source}
Activité : {activite}
Constat : {constat}

--- TYPOLOGIES (LISTE FERMEE) ---
{typologies_list}

Règle :
CAUSE_RACINE = une seule valeur de la liste

--- CAS REELS PRIORITAIRES ---
{memoire}

--- CONSIGNES ---
- comprendre le problème réel
- s’inspirer des meilleurs cas (score élevé + modifié humain)
- reproduire la logique cause → action
- adapter au contexte (ne pas copier)

--- REDACTION ---
- concret terrain
- pas de blabla
- pas de phrases vagues
- actionnable immédiatement

Réponds uniquement avec les 6 lignes.
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
            "temperature": 0.1
        }

        ai_response = requests.post(ai_url, headers=ai_headers, json=ai_data, timeout=60)
        ai_json = ai_response.json()

        if "choices" not in ai_json:
            return {"resultat": "ERREUR OPENAI : " + str(ai_json)}

        texte = ai_json["choices"][0]["message"]["content"]
        return {"resultat": texte}

    except Exception as e:
        return {"resultat": "ERREUR PYTHON : " + str(e)}
