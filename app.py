from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel
import hmac
import os
import re
import unicodedata
import requests

app = FastAPI()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_ID = os.getenv("SITE_ID")
LIST_ID = os.getenv("LIST_ID")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
API_KEY = os.getenv("API_KEY", "").strip()


class AnalysePayload(BaseModel):
    constat: str = ""
    source: str = ""
    activite: str = ""
    typologies: str = ""
    action_immediate: str = ""
    analyse: str = ""
    cause_racine: str = ""
    action_corrective_necessaire: str = ""
    action_corrective: str = ""
    mesure_efficacite: str = ""


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
    Constat: str = ""
    Source: str = ""
    Activite: str = ""
    DateCas: str = ""
    ActionImmediate_Finale: str = ""
    Analyse_Finale: str = ""
    Typologie_Finale: str = ""
    ActionCorrective_Finale: str = ""
    MesureEfficacite_Finale: str = ""
    Tags: str = ""
    NomFichierSource: str = ""


@app.get("/")
def root():
    return {"status": "ok", "message": "API vivante"}


@app.get("/health")
def health():
    return {"status": "ok"}


def require_api_key(request: Request):
    if not API_KEY:
        return
    supplied = request.headers.get("X-API-Key", "")
    if not supplied or not hmac.compare_digest(supplied, API_KEY):
        raise HTTPException(status_code=401, detail="unauthorized")


def get_access_token():
    missing = [name for name, value in {
        "TENANT_ID": TENANT_ID,
        "CLIENT_ID": CLIENT_ID,
        "CLIENT_SECRET": CLIENT_SECRET,
    }.items() if not value]
    if missing:
        raise HTTPException(status_code=503, detail=f"configuration Azure manquante: {', '.join(missing)}")

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    try:
        response = requests.post(token_url, data=token_data, timeout=30)
        response.raise_for_status()
        payload = response.json()
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"Azure token indisponible: {exc}") from exc

    access_token = payload.get("access_token")
    if not access_token:
        raise HTTPException(status_code=502, detail="Azure n'a pas renvoyé de jeton d'accès")
    return access_token


def graph_headers(access_token):
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def graph_json(method, url, access_token, **kwargs):
    try:
        response = requests.request(
            method,
            url,
            headers=graph_headers(access_token),
            timeout=30,
            **kwargs,
        )
    except requests.RequestException as exc:
        raise HTTPException(status_code=502, detail=f"Microsoft Graph inaccessible: {exc}") from exc

    try:
        detail = response.json() if response.content else {}
    except ValueError:
        detail = {"raw": response.text}

    if response.status_code < 200 or response.status_code >= 300:
        raise HTTPException(
            status_code=502,
            detail={"step": "sharepoint", "graph_status": response.status_code, "graph_detail": detail},
        )
    return detail


def get_all_memory_items(access_token):
    if not SITE_ID or not LIST_ID:
        raise HTTPException(status_code=503, detail="SITE_ID/LIST_ID manquant")

    url = (
        f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
        "?$expand=fields&$top=200"
    )
    items = []
    pages = 0
    while url:
        pages += 1
        if pages > 100:
            raise HTTPException(status_code=502, detail="pagination SharePoint anormalement longue")
        payload = graph_json("GET", url, access_token)
        items.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return items


def normalize_bool(value):
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"oui", "true", "1", "vrai", "yes"}


def normalize_qualite(value):
    txt = str(value or "").strip().lower()
    if txt == "faible":
        return "Faible"
    if txt in {"reference", "référence", "ref"}:
        return "Reference"
    return "Bon"


def short_text(value, max_length=255):
    return str(value or "")[:max_length]


def full_constat(fields):
    return str(fields.get("ActionImmediate_IA") or fields.get("Title") or "")


def build_sharepoint_fields_from_payload(data: MemoirePayload):
    constat = str(data.Constat or "")
    return {
        "Title": constat[:255],
        "Source": short_text(data.Source),
        "Activite": short_text(data.Activite),
        "DateCas": data.DateCas,
        "ActionImmediate_IA": constat,
        "ActionImmediate_Finale": data.ActionImmediate_Finale,
        "Analyse_IA": data.Analyse_IA,
        "Analyse_Finale": data.Analyse_Finale,
        "Typologie_IA": short_text(data.Typologie_IA),
        "Typologie_Finale": short_text(data.Typologie_Finale),
        "ActionCorrective_IA": data.ActionCorrective_IA,
        "ActionCorrective_Finale": data.ActionCorrective_Finale,
        "MesureEfficacite_IA": data.MesureEfficacite_IA,
        "MesureEfficacite_Finale": data.MesureEfficacite_Finale,
        "ModifieParHumain": normalize_bool(data.ModifieParHumain),
        "QualiteCas": normalize_qualite(data.QualiteCas),
        "Tags": short_text(data.Tags),
        "NomFichierSource": short_text(data.NomFichierSource),
    }


def build_sharepoint_update_fields(data: MemoireUpdatePayload):
    constat = str(data.Constat or "")
    fields = {
        "Title": constat[:255],
        "ActionImmediate_IA": constat,
        "Source": short_text(data.Source),
        "Activite": short_text(data.Activite),
        "DateCas": data.DateCas,
        "ActionImmediate_Finale": data.ActionImmediate_Finale,
        "Analyse_Finale": data.Analyse_Finale,
        "Typologie_Finale": short_text(data.Typologie_Finale),
        "ActionCorrective_Finale": data.ActionCorrective_Finale,
        "MesureEfficacite_Finale": data.MesureEfficacite_Finale,
        "ModifieParHumain": True,
        "Tags": short_text(data.Tags),
    }
    if data.NomFichierSource:
        fields["NomFichierSource"] = short_text(data.NomFichierSource)
    return fields


def normalize_text(txt):
    txt = str(txt or "").lower().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^a-z0-9\s]", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


def extract_keywords(txt):
    stopwords = {
        "de", "des", "du", "le", "la", "les", "un", "une", "et", "ou", "a", "au",
        "aux", "sur", "dans", "par", "pour", "avec", "sans", "en", "d", "l", "est",
        "sont", "etre", "avoir", "plus", "moins", "tres", "non", "pas", "qui", "que",
        "quoi", "dont", "se", "ce", "cet", "cette", "ces",
    }
    return list(dict.fromkeys(
        mot for mot in normalize_text(txt).split() if len(mot) >= 4 and mot not in stopwords
    ))


def shared_keyword_score(a, b, weight, cap):
    common = set(extract_keywords(a)) & set(extract_keywords(b))
    return min(cap, len(common) * weight)


def score_case(constat, source, activite, fields):
    score = 0
    title = normalize_text(full_constat(fields))
    old_activite = normalize_text(fields.get("Activite", ""))
    old_source = normalize_text(fields.get("Source", ""))
    tags = normalize_text(fields.get("Tags", ""))
    action_immediate_finale = normalize_text(fields.get("ActionImmediate_Finale", ""))
    analyse_finale = normalize_text(fields.get("Analyse_Finale", ""))
    typologie_finale = normalize_text(fields.get("Typologie_Finale", ""))
    action_corrective_finale = normalize_text(fields.get("ActionCorrective_Finale", ""))
    mesure_finale = normalize_text(fields.get("MesureEfficacite_Finale", ""))

    if normalize_text(source) and normalize_text(source) == old_source:
        score += 45
    else:
        score += shared_keyword_score(source, old_source, 10, 30)

    if normalize_text(activite) and normalize_text(activite) == old_activite:
        score += 25
    else:
        score += shared_keyword_score(activite, old_activite, 5, 15)

    constat_keywords = extract_keywords(constat)
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

    constat_norm = normalize_text(constat)
    bloc = " ".join([
        title,
        old_activite,
        old_source,
        tags,
        action_immediate_finale,
        analyse_finale,
        typologie_finale,
        action_corrective_finale,
        mesure_finale,
    ])
    if constat_norm and constat_norm in bloc:
        score += 12
    if normalize_bool(fields.get("ModifieParHumain", False)):
        score += 8
    return score


def build_memory_context(constat, source, activite, items, limit=5):
    scored = []
    for item in items:
        fields = item.get("fields", {})
        if not full_constat(fields).strip():
            continue
        score = score_case(constat, source, activite, fields)
        if score > 0:
            scored.append((score, fields))
    scored.sort(key=lambda x: x[0], reverse=True)

    if not scored:
        return "Aucun cas antérieur pertinent trouvé dans SharePoint."

    blocks = []
    for rank, (score, f) in enumerate(scored[:limit], 1):
        blocks.append(
            f"CAS REEL {rank}:\n"
            f"Score pertinence: {score}\n"
            f"Constat: {full_constat(f)}\n"
            f"Source: {f.get('Source', '')}\n"
            f"Activite: {f.get('Activite', '')}\n"
            f"Action immédiate finale: {f.get('ActionImmediate_Finale', '')}\n"
            f"Analyse finale: {f.get('Analyse_Finale', '')}\n"
            f"Typologie finale: {f.get('Typologie_Finale', '')}\n"
            f"Action corrective finale: {f.get('ActionCorrective_Finale', '')}\n"
            f"Mesure efficacité finale: {f.get('MesureEfficacite_Finale', '')}\n"
            f"Modifié par humain: {f.get('ModifieParHumain', '')}\n"
            f"Qualité du cas: {f.get('QualiteCas', '')}\n"
            f"Tags: {f.get('Tags', '')}\n---"
        )
    return "\n".join(blocks)


def analyse_impl(data: AnalysePayload):
    access_token = get_access_token()
    items = get_all_memory_items(access_token)
    memory = build_memory_context(data.constat, data.source, data.activite, items, limit=5)

    typologies = []
    raw_typologies = (data.typologies or "").replace("||", "\n")
    for line in raw_typologies.splitlines():
        line = line.strip().lstrip("-").strip()
        if line:
            typologies.append(line)
    typologies_list = "\n".join(f"- {x}" for x in typologies)

    corrective_rule = ""
    if normalize_text(data.action_corrective_necessaire) == "non":
        corrective_rule = (
            "L'utilisateur a décidé qu'aucune action corrective n'est nécessaire. "
            "ACTION_CORRECTIVE et MESURE_EFFICACITE doivent donc rester vides."
        )

    prompt = f"""
Tu es un responsable QSE chantier expérimenté.
Tu dois améliorer un plan d'action existant sans détruire le travail déjà fait.

Réponds STRICTEMENT avec ces 6 lignes :
CONSTAT=
ACTION_IMMEDIATE=
ANALYSE=
CAUSE_RACINE=
ACTION_CORRECTIVE=
MESURE_EFFICACITE=

NOUVEAU CAS
Source : {data.source}
Activité : {data.activite}
Constat : {data.constat}

DONNEES UTILISATEUR DEJA SAISIES
Action immédiate : {data.action_immediate}
Analyse : {data.analyse}
Cause racine : {data.cause_racine}
Action corrective nécessaire : {data.action_corrective_necessaire}
Action corrective : {data.action_corrective}
Mesure efficacité : {data.mesure_efficacite}

TYPOLOGIES (LISTE FERMEE)
{typologies_list}
Règle : CAUSE_RACINE = UNE SEULE valeur exacte de la liste.

CAS MEMOIRE PRIORITAIRES
{memory}

REGLES
- concret terrain, actionnable immédiatement, pas de phrases vagues
- si un champ est déjà rempli, l'améliorer sans supprimer une information pertinente
- si un champ est vide, le compléter si nécessaire
- respecter la décision de l'utilisateur sur la nécessité d'une action corrective
- {corrective_rule or "si une action corrective est nécessaire, proposer une action proportionnée et une mesure d'efficacité vérifiable"}

Réponds uniquement avec les 6 lignes.
"""

    if not OPENAI_API_KEY:
        raise HTTPException(status_code=503, detail="OPENAI_API_KEY manquante")

    try:
        ai_response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "gpt-4.1-mini",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.1,
            },
            timeout=60,
        )
        ai_response.raise_for_status()
        ai_json = ai_response.json()
        text = ai_json["choices"][0]["message"]["content"]
    except Exception as exc:
        raise HTTPException(status_code=502, detail=f"OpenAI indisponible: {exc}") from exc

    return {"status": "ok", "resultat": text}


@app.post("/analyse")
def analyse_post(data: AnalysePayload, request: Request):
    require_api_key(request)
    return analyse_impl(data)


@app.get("/analyse")
def analyse_get(
    request: Request,
    constat: str,
    source: str = "",
    activite: str = "",
    typologies: str = "",
    action_immediate: str = "",
    analyse: str = "",
    cause_racine: str = "",
    action_corrective_necessaire: str = "",
    action_corrective: str = "",
    mesure_efficacite: str = "",
):
    require_api_key(request)
    return analyse_impl(AnalysePayload(
        constat=constat,
        source=source,
        activite=activite,
        typologies=typologies,
        action_immediate=action_immediate,
        analyse=analyse,
        cause_racine=cause_racine,
        action_corrective_necessaire=action_corrective_necessaire,
        action_corrective=action_corrective,
        mesure_efficacite=mesure_efficacite,
    ))


@app.get("/sharepoint_columns")
def sharepoint_columns(request: Request):
    require_api_key(request)
    access_token = get_access_token()
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/columns"
    detail = graph_json("GET", url, access_token)
    return {"status": "ok", "detail": detail}


@app.post("/memoire")
def memoire(data: MemoirePayload, request: Request):
    require_api_key(request)
    access_token = get_access_token()
    fields = build_sharepoint_fields_from_payload(data)
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items"
    detail = graph_json("POST", url, access_token, json={"fields": fields})
    item_id = str(detail.get("id", "")).strip()
    if not item_id:
        raise HTTPException(status_code=502, detail="SharePoint a créé un élément sans renvoyer son ID")
    return {"status": "ok", "id": item_id}


@app.post("/memoire_update")
def memoire_update(data: MemoireUpdatePayload, request: Request):
    require_api_key(request)
    item_id = data.SharePointID.strip()
    if not item_id:
        raise HTTPException(status_code=400, detail="SharePointID manquant")
    access_token = get_access_token()
    fields = build_sharepoint_update_fields(data)
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/lists/{LIST_ID}/items/{item_id}/fields"
    graph_json("PATCH", url, access_token, json=fields)
    return {"status": "ok", "id": item_id}
