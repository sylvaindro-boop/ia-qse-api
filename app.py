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


def get_headers(access_token):
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }


def normalize_bool(value):
    if isinstance(value, bool):
        return value

    if value is None:
        return False

    txt = str(value).strip().lower()

    if txt in ["oui", "true", "1", "vrai", "yes"]:
        return True

    if txt in ["non", "false", "0", "faux", "no", ""]:
        return False

    return False


def normalize_qualite(value):
    if value is None:
        return "Bon"

    txt = str(value).strip().lower()

    if txt == "faible":
        return "Faible"
    if txt == "bon":
        return "Bon"
    if txt in ["reference", "référence", "ref"]:
        return "Reference"

    return "Bon"


def build_sharepoint_fields_from_payload(data: MemoirePayload):
    return {
        "Title": data.Constat,
        "Source": data.Source,
        "Activite": data.Activite,
        "DateCas": data.DateCas,

        "ActionImmediate_IA": data.ActionImmediate_IA,
        "ActionImmediate_Finale": data.ActionImmediate_Finale,

        "Analyse_IA": data.Analyse_IA,
        "Analyse_Finale": data.Analyse_Finale,

        "Typologie_IA": data.Typologie_IA,
        "Typologie_Finale": data.Typologie_Finale,

        "ActionCorrective_IA": data.ActionCorrective_IA,
        "ActionCorrective_Finale": data.ActionCorrective_Finale,

        "MesureEfficacite_IA": data.MesureEfficacite_IA,
        "MesureEfficacite_Finale": data.MesureEfficacite_Finale,

        "ModifieParHumain": normalize_bool(data.ModifieParHumain),
        "QualiteCas": normalize_qualite(data.QualiteCas),
        "Tags": data.Tags,
        "NomFichierSource": data.NomFichierSource
    }


def build_sharepoint_update_fields(data: MemoireUpdatePayload):
    return {
        "ActionImmediate_Finale": data.ActionImmediate_Finale,
        "Analyse_Finale": data.Analyse_Finale,
        "Typologie_Finale": data.Typologie_Finale,
        "ActionCorrective_Finale": data.ActionCorrective_Finale,
        "MesureEfficacite_Finale": data.MesureEfficacite_Finale,
        "ModifieParHumain": True
    }


def normalize_text(txt):
    if txt is None:
        return ""

    txt = str(txt).lower().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt


def extract_keywords(txt):
    stopwords = {
        "de", "des", "du", "le", "la", "les", "un", "une", "et", "ou", "a", "au",
        "aux", "sur", "dans", "par", "pour", "avec", "sans", "en", "d", "l",
        "est", "sont", "etre", "avoir", "plus", "moins", "tres", "non", "pas",
        "qui", "que", "quoi", "dont", "se", "ce", "cet", "cette", "ces"
    }

    txt = normalize_text(txt)
    mots = txt.split()

    resultat = []
    for mot in mots:
        if len(mot) >= 4 and mot not in stopwords:
            resultat.append(mot)

    return list(dict.fromkeys(resultat))


def score_case(constat, fields):
    score = 0

    constat_txt = normalize_text(constat)
    constat_keywords = extract_keywords(constat)

    title = normalize_text(fields.get("Title", ""))
    activite = normalize_text(fields.get("Activite", ""))
    source = normalize_text(fields.get("Source", ""))
    tags = normalize_text(fields.get("Tags", ""))

    action_immediate_finale = normalize_text(fields.get("ActionImmediate_Finale", ""))
    analyse_finale = normalize_text(fields.get("Analyse_Finale", ""))
    typologie_finale = normalize_text(fields.get("Typologie_Finale", ""))
    action_corrective_finale = normalize_text(fields.get("ActionCorrective_Finale", ""))
    mesure_finale = normalize_text(fields.get("MesureEfficacite_Finale", ""))

    bloc = " ".join([
        title,
        activite,
        source,
        tags,
        action_immediate_finale,
        analyse_finale,
        typologie_finale,
        action_corrective_finale,
        mesure_finale
    ])

    # 1) priorité forte à la source
    if source:
        for mot in extract_keywords(source):
            if mot and mot in constat_txt:
                score += 20

    # 2) activité = bonus léger seulement
    if activite:
        for mot in extract_keywords(activite):
            if mot and mot in constat_txt:
                score += 4

    # 3) similarité du constat avec les champs finaux
    for mot in constat_keywords:
        if mot in title:
            score += 10
        if mot in analyse_finale:
            score += 6
        if mot in typologie_finale:
            score += 4
        if mot in action_corrective_finale:
            score += 4
        if mot in action_immediate_finale:
            score += 3
        if mot in mesure_finale:
            score += 2
        if mot in tags:
            score += 2

    # 4) bonus si le constat complet est retrouvé
    if constat_txt and constat_txt in bloc:
        score += 12

    # 5) bonus si modifié par humain
    modifie = fields.get("ModifieParHumain", False)
    if modifie is True:
        score += 8

    return score

def build_memory_context(constat, items, limit=5):
    scored = []

    for item in items:
        fields = item.get("fields", {})
        if not fields:
            continue

        titre = str(fields.get("Title", "")).strip()
        if titre == "":
            continue

        s = score_case(constat, fields)
        scored.append((s, fields))

    scored.sort(key=lambda x: x[0], reverse=True)

    meilleurs = []
    for s, f in scored:
        if len(meilleurs) >= limit:
            break
        meilleurs.append((s, f))

    if len(meilleurs) == 0:
        return "Aucun cas antérieur pertinent trouvé dans SharePoint."

    blocs = []
    rang = 1

    for s, f in meilleurs:
        bloc = f"""
CAS REEL {rang} :
Score pertinence : {s}
Constat : {f.get("Title", "")}
Source : {f.get("Source", "")}
Activite : {f.get("Activite", "")}
Action immédiate IA : {f.get("ActionImmediate_IA", "")}
Action immédiate finale : {f.get("ActionImmediate_Finale", "")}
Analyse IA : {f.get("Analyse_IA", "")}
Analyse finale : {f.get("Analyse_Finale", "")}
Typologie IA : {f.get("Typologie_IA", "")}
Typologie finale : {f.get("Typologie_Finale", "")}
Action corrective IA : {f.get("ActionCorrective_IA", "")}
Action corrective finale : {f.get("ActionCorrective_Finale", "")}
Mesure efficacité IA : {f.get("MesureEfficacite_IA", "")}
Mesure efficacité finale : {f.get("MesureEfficacite_Finale", "")}
Modifié par humain : {f.get("ModifieParHumain", "")}
Qualité du cas : {f.get("QualiteCas", "")}
Tags : {f.get("Tags", "")}
---
"""
        blocs.append(bloc)
        rang += 1

    return "\n".join(blocs)


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

        # Typologies
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

Tu dois produire un plan d'action concret, utile et directement applicable terrain.

Réponds STRICTEMENT avec ces 6 lignes :
CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=

-----------------------
NOUVEAU CAS
-----------------------
Source : {source}
Activité : {activite}
Constat : {constat}

-----------------------
TYPOLOGIES (LISTE FERMEE)
-----------------------
{typologies_list}

Règle :
CAUSE_RACINE = UNE SEULE valeur de la liste

-----------------------
CAS MEMOIRE (PRIORITAIRES)
-----------------------
{memoire}

-----------------------
METHODE
-----------------------
1. Comprendre le problème réel
2. Identifier les cas similaires
3. S’inspirer des meilleurs cas (score élevé + modifié humain)
4. Reproduire la logique métier (cause → action → contrôle)
5. Adapter sans copier

-----------------------
REGLES
-----------------------
- concret terrain
- pas de blabla
- pas de phrases vagues
- actionnable immédiatement
- privilégier les champs FINAUX
- privilégier les cas modifiés par humain

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


@app.get("/sharepoint_columns")
def sharepoint_columns():
    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/columns"
        response = requests.get(url, headers=get_headers(access_token), timeout=30)

        try:
            detail = response.json()
        except Exception:
            detail = response.text

        return {
            "status": "ok" if response.status_code == 200 else "error",
            "http_status": response.status_code,
            "detail": detail
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}


@app.post("/memoire")
def memoire(data: MemoirePayload):
    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        fields = build_sharepoint_fields_from_payload(data)

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
        payload = {"fields": fields}

        sp_response = requests.post(
            sp_url,
            headers=get_headers(access_token),
            json=payload,
            timeout=30
        )

        try:
            detail = sp_response.json()
        except Exception:
            detail = sp_response.text

        if sp_response.status_code not in [200, 201]:
            return {
                "status": "error",
                "step": "sharepoint_create",
                "http_status": sp_response.status_code,
                "detail": detail,
                "payload_sent": payload
            }

        return {
            "status": "ok",
            "http_status": sp_response.status_code,
            "detail": detail,
            "payload_sent": payload
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}


@app.post("/memoire_update")
def memoire_update(data: MemoireUpdatePayload):
    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        item_id = data.SharePointID.strip()
        if item_id == "":
            return {"status": "error", "step": "missing_sharepoint_id"}

        fields = build_sharepoint_update_fields(data)

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields"

        sp_response = requests.patch(
            sp_url,
            headers=get_headers(access_token),
            json=fields,
            timeout=30
        )

        try:
            detail = sp_response.json()
        except Exception:
            detail = sp_response.text

        if sp_response.status_code not in [200, 201]:
            return {
                "status": "error",
                "step": "sharepoint_update",
                "http_status": sp_response.status_code,
                "detail": detail,
                "payload_sent": fields
            }

        return {
            "status": "ok",
            "http_status": sp_response.status_code,
            "detail": detail,
            "payload_sent": fields
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}


@app.get("/memoire_test")
def memoire_test():
    data_input = MemoirePayload(
        Constat="Test déchets chantier",
        Source="Audit chantier",
        Activite="Déchets",
        DateCas="2026-03-28",

        ActionImmediate_IA="Nettoyage immédiat",
        ActionImmediate_Finale="Nettoyage immédiat",

        Analyse_IA="Tri non respecté",
        Analyse_Finale="Tri non respecté",

        Typologie_IA="Rigueur",
        Typologie_Finale="Rigueur",

        ActionCorrective_IA="Rappel des consignes",
        ActionCorrective_Finale="Rappel des consignes",

        MesureEfficacite_IA="Contrôle terrain",
        MesureEfficacite_Finale="Contrôle terrain",

        ModifieParHumain="Non",
        QualiteCas="Bon",
        Tags="dechets;chantier",
        NomFichierSource="Portail_Battaglino-Déconstruction_V6.xlsm"
    )

    try:
        access_token, token_error = get_access_token()
        if not access_token:
            return {"status": "error", "step": "token", "detail": token_error}

        fields = build_sharepoint_fields_from_payload(data_input)

        sp_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
        payload = {"fields": fields}

        sp_response = requests.post(
            sp_url,
            headers=get_headers(access_token),
            json=payload,
            timeout=30
        )

        try:
            detail = sp_response.json()
        except Exception:
            detail = sp_response.text

        return {
            "status": "ok" if sp_response.status_code in [200, 201] else "error",
            "http_status": sp_response.status_code,
            "detail": detail,
            "payload_sent": payload
        }

    except Exception as e:
        return {"status": "error", "step": "python", "detail": str(e)}
