from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import requests
import os
from datetime import datetime

app = FastAPI()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")
LIST_ID = os.getenv("LIST_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class MemoirePayload(BaseModel):
    MemoireID: str | None = None
    Constat: str = ""
    Client: str = "Battaglino"
    Source: str = ""
    Activite: str = ""
    DateCas: str | None = None

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

    ModifieParHumain: str = "Non"
    QualiteCas: str = "Moyen"
    Tags: str = ""
    NomFichierSource: str = "Portail_Battaglino-Déconstruction_V5.xlsm"


def get_access_token():
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    r = requests.post(token_url, data=token_data, timeout=30)
    r.raise_for_status()
    token_json = r.json()

    access_token = token_json.get("access_token")
    if not access_token:
        raise HTTPException(status_code=500, detail=f"Token Azure invalide: {token_json}")

    return access_token


def graph_headers(token: str):
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


def graph_list_items_url(expand_fields=True):
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_ID}/items"
    if expand_fields:
        url += "?expand=fields"
    return url


def normalize(txt: str) -> str:
    if not txt:
        return ""
    txt = txt.lower().strip()
    repl = {
        "é": "e", "è": "e", "ê": "e", "ë": "e",
        "à": "a", "â": "a",
        "ù": "u", "û": "u",
        "ô": "o", "ö": "o",
        "î": "i", "ï": "i",
        "ç": "c"
    }
    for k, v in repl.items():
        txt = txt.replace(k, v)
    return txt


def score_case(fields: dict, constat: str, source: str, activite: str) -> int:
    score = 0
    c = normalize(constat)
    f_constat = normalize(fields.get("Title", ""))
    f_source = normalize(fields.get("Source", ""))
    f_activite = normalize(fields.get("Activite", ""))
    f_tags = normalize(fields.get("Tags", ""))

    for mot in c.split():
        if len(mot) >= 4 and mot in f_constat:
            score += 3
        if len(mot) >= 4 and mot in f_tags:
            score += 2

    if normalize(source) and normalize(source) == f_source:
        score += 4

    if normalize(activite) and normalize(activite) == f_activite:
        score += 4

    if normalize(fields.get("ModifieParHumain", "")) == "oui":
        score += 3

    qualite = normalize(fields.get("QualiteCas", ""))
    if qualite in ("bon", "bonne", "excellent", "haute"):
        score += 2

    return score


def build_memory_examples(items: list[dict], constat: str, source: str, activite: str, max_items: int = 5) -> str:
    ranked = []
    for item in items:
        fields = item.get("fields", {})
        ranked.append((score_case(fields, constat, source, activite), fields))

    ranked.sort(key=lambda x: x[0], reverse=True)
    selected = [x[1] for x in ranked if x[0] > 0][:max_items]

    if not selected:
        selected = [item.get("fields", {}) for item in items[:3]]

    blocs = []
    for i, f in enumerate(selected, start=1):
        blocs.append(
            f"""CAS REEL {i} :
Constat : {f.get("Title", "")}
Source : {f.get("Source", "")}
Activite : {f.get("Activite", "")}
Action immédiate finale : {f.get("ActionImmediate_Finale", "")}
Analyse finale : {f.get("Analyse_Finale", "")}
Cause finale : {f.get("Cause_Finale", "")}
Action corrective finale : {f.get("ActionCorrective_Finale", "")}
Mesure efficacité finale : {f.get("MesureEfficacite_Finale", "")}
Modifié par humain : {f.get("ModifieParHumain", "")}
Qualité : {f.get("QualiteCas", "")}
---"""
        )

    return "\n".join(blocs)


def get_all_memory_items(token: str):
    url = graph_list_items_url(expand_fields=True)
    r = requests.get(url, headers=graph_headers(token), timeout=30)
    r.raise_for_status()
    data = r.json()
    return data.get("value", [])


def create_sharepoint_fields(data: MemoirePayload):
    date_cas = data.DateCas or datetime.now().strftime("%Y-%m-%d")

    return {
        "Title": data.Constat,
        "Client": data.Client,
        "Source": data.Source,
        "Activite": data.Activite,
        "DateCas": date_cas,

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


@app.get("/")
def root():
    return {"status": "ok", "message": "API vivante"}


@app.get("/analyse")
def analyse(constat: str, source: str = "", activite: str = ""):
    try:
        token = get_access_token()
        items = get_all_memory_items(token)

        memoire = build_memory_examples(items, constat, source, activite, max_items=5)

        prompt = f"""
Tu es un expert QSE terrain.

Tu dois répondre EXACTEMENT avec ces 6 lignes, sans texte avant ni après :

CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=

Tu dois t'inspirer d'abord des cas réels ci-dessous, en privilégiant les cas modifiés par humain et de bonne qualité.

{memoire}

Nouveau constat :
{constat}

Source :
{source}

Activité :
{activite}

Règles :
- Cause racine courte
- Actions concrètes terrain
- Pas de blabla
- Sois utile et opérationnel
"""

        ai_url = "https://api.openai.com/v1/chat/completions"
        ai_headers = {
            "Authorization": f"Bearer {OPENAI_API_KEY}",
            "Content-Type": "application/json"
        }
        ai_data = {
            "model": "gpt-4.1-mini",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.2
        }

        ai_response = requests.post(ai_url, headers=ai_headers, json=ai_data, timeout=60)
        ai_response.raise_for_status()
        ai_json = ai_response.json()

        texte = ai_json["choices"][0]["message"]["content"]
        return {"resultat": texte}

    except Exception as e:
        return {"resultat": "ERREUR PYTHON : " + str(e)}


@app.post("/memoire")
def create_memoire(data: MemoirePayload):
    try:
        token = get_access_token()
        fields = create_sharepoint_fields(data)

        url = graph_list_items_url(expand_fields=False)
        payload = {"fields": fields}

        r = requests.post(url, headers=graph_headers(token), json=payload, timeout=30)
        r.raise_for_status()
        j = r.json()

        return {
            "status": "ok",
            "action": "created",
            "sharepoint_item_id": j.get("id"),
            "detail": j
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.put("/memoire/{item_id}")
def update_memoire(item_id: str, data: MemoirePayload):
    try:
        token = get_access_token()
        fields = create_sharepoint_fields(data)

        url = f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields"
        r = requests.patch(url, headers=graph_headers(token), json=fields, timeout=30)
        r.raise_for_status()

        return {
            "status": "ok",
            "action": "updated",
            "sharepoint_item_id": item_id
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/memoire_test")
def memoire_test():
    data_input = MemoirePayload(
        Constat="Test déchets chantier",
        Client="Test client",
        Source="Audit chantier",
        Activite="Déchets",
        ActionImmediate_IA="Nettoyage immédiat",
        ActionImmediate_Finale="Nettoyage immédiat",
        Analyse_IA="Tri non respecté",
        Analyse_Finale="Tri non respecté",
        Cause_IA="Rigueur",
        Cause_Finale="Rigueur",
        Typologie_Finale="Rigueur",
        ActionCorrective_IA="Rappel des consignes",
        ActionCorrective_Finale="Rappel des consignes",
        MesureEfficacite_IA="Contrôle terrain",
        MesureEfficacite_Finale="Contrôle terrain",
        ModifieParHumain="Non",
        QualiteCas="Bon",
        Tags="dechets;chantier"
    )

    return create_memoire(data_input)
