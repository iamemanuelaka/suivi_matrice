"""Phase 3 : questionnaire Kobo de suivi périodique.

  - la LISTE des activités (activites.csv) est générée depuis la base : c'est elle qui alimente la sélection du technicien
    et l'affichage en lecture seule des informations de référence (pulldata) ;
  - la LISTE des structures (kobo/structures.csv, tenue à la main : une ligne par structure, colonne « regions » = codes des
    régions couvertes séparés par des espaces, ou « * » pour toutes) ;
  - la LISTE des types d'activités (types_activite.csv, générée depuis les catégories de config.yaml) ;
  - choix en cascade : structure -> types d'activités de ses régions -> activités de ce type dans ses régions
    (colonnes « structures » et « type » des CSV + choice_filter) ;
  - le formulaire XLSForm est généré par script (il ne change que si les champs de suivi changent) ;
  - la synchronisation envoie le CSV à Kobo (fichier média du formulaire) puis redéploie.
"""
import csv
import io
import json
import re
from datetime import date, datetime
from pathlib import Path

import requests
from openpyxl import Workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from .referentiel import Referentiel

CSV_NAME = "activites.csv"
CSV_COLUMNS = ["name", "label", "ref_titre", "ref_marche", "ref_suivi", "structures", "type"]
STRUCTURES_NAME = "structures.csv"
STRUCTURES_COLUMNS = ["name", "label", "regions"]
TYPES_NAME = "types_activite.csv"
TYPES_COLUMNS = ["name", "label", "structures"]
# Cascade : structure -> types d'activités présents dans ses régions -> activités de ce type dans ses régions
FILTRE_TYPES = "selected(structures, ${structure})"
FILTRE_ACTIVITES = "selected(structures, ${structure}) and type = ${type_activite}"
FORM_ID = "suivi_activites_pcr"


# --------------------------------------------------------------------------
# Liste des activités
# --------------------------------------------------------------------------
def _clean(s) -> str:
    """Une seule ligne, sans retour à la ligne (le CSV lu par pulldata ne les supporte pas)."""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _montant(n) -> str:
    return f"{n:,}".replace(",", " ")


def derniers_suivis(ref: Referentiel) -> dict:
    """Dernier suivi VALIDE de chaque activité (date d'actualisation, puis date d'import, puis identifiant Kobo)."""
    best = {}
    for r in ref.base.rows("Suivi_Activites"):
        if r["Statut_Traitement"] != "VALIDE" or r["Taux_Execution_Physique"] is None:
            continue
        key = (r["Date_Actualisation"] or date.min, r["Date_Importation"] or datetime.min, str(r["ID_SOUMISSION_KOBO"]))
        if r["CODE_ACTIVITE"] not in best or key > best[r["CODE_ACTIVITE"]][0]:
            best[r["CODE_ACTIVITE"]] = (key, r)
    return {c: v[1] for c, v in best.items()}


# --------------------------------------------------------------------------
# Structures (UTR des régions, puis toute structure ajoutée au fichier)
# --------------------------------------------------------------------------
class StructuresError(ValueError):
    pass


def load_structures(cfg) -> list[dict]:
    """Lit kobo/structures.csv et le contrôle : noms uniques sans espace, régions connues (ou « * »)."""
    chemin = cfg.path("kobo_structures")
    if not chemin.is_file():
        raise StructuresError(f"Liste des structures introuvable : {chemin}")
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        lignes = [{k.strip(): (v or "").strip() for k, v in r.items() if k} for r in csv.DictReader(f)]
    out, vus = [], set()
    for n, r in enumerate(lignes, 2):
        if not any(r.values()):
            continue
        nom, lib, regs = r.get("name", ""), r.get("label", ""), r.get("regions", "").split()
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_\-.]*", nom):
            raise StructuresError(f"structures.csv ligne {n} : nom « {nom} » invalide (lettres, chiffres, _ sans espace).")
        if nom in vus:
            raise StructuresError(f"structures.csv ligne {n} : nom « {nom} » en double.")
        if not lib:
            raise StructuresError(f"structures.csv ligne {n} : libellé manquant pour {nom}.")
        inconnues = [x for x in regs if x != "*" and x not in cfg.regions]
        if not regs or inconnues:
            raise StructuresError(f"structures.csv ligne {n} : régions {' '.join(inconnues) or '(aucune)'} inconnues "
                                  f"(codes possibles : {' '.join(cfg.regions)} ou *).")
        vus.add(nom)
        out.append({"name": nom, "label": _clean(lib), "regions": " ".join(regs)})
    if not out:
        raise StructuresError("structures.csv ne contient aucune structure.")
    return out


def structures_de(region: str, structures: list[dict]) -> list[str]:
    return [s["name"] for s in structures if "*" in s["regions"].split() or region in s["regions"].split()]


def activities_csv_rows(ref: Referentiel, structures: list[dict] | None = None) -> list[dict]:
    """Une ligne par activité ACTIVE (les activités archivées ne sont plus proposées aux techniciens).
    Colonne « structures » : structures qui peuvent la choisir (filtre en cascade du formulaire)."""
    structures = load_structures(ref.cfg) if structures is None else structures
    last, rows = derniers_suivis(ref), []
    for a in sorted(ref.activities(), key=lambda x: x["CODE_ACTIVITE"]):
        if a["Statut"] != "Actif":
            continue
        code, region = a["CODE_ACTIVITE"], ref.cfg.regions.get(a["Region"], a["Region"])
        nature = ref.cfg.categories.get(a["Categorie"], a["Categorie"])
        label = f"{nature} · {region} · {_clean(a['Departement']) or '—'} — {_clean(a['Entreprise']) or 'titulaire non renseigné'}"
        pieces = []
        if a["Entreprise"]:
            pieces.append(f"Titulaire : {_clean(a['Entreprise'])}")
        if a["Montant_Marche"]:
            pieces.append(f"Marché : {_montant(a['Montant_Marche'])} FCFA")
        if a["Delai_Mois"]:
            pieces.append(f"Délai : {a['Delai_Mois']} mois")
        if a["Date_Debut"]:
            pieces.append(f"Début : {a['Date_Debut']:%d/%m/%Y}")
        if a["Date_Fin_Previsionnelle"]:
            pieces.append(f"Fin prévue : {a['Date_Fin_Previsionnelle']:%d/%m/%Y}")
        s = last.get(code)
        rows.append({
            "name": code, "label": label,
            "ref_titre": _clean(a['Designation']) or f"{nature} — désignation non renseignée",
            "ref_marche": " · ".join(pieces) or "Informations contractuelles non encore renseignées.",
            "ref_suivi": (f"Dernier suivi enregistré : {s['Taux_Execution_Physique']} % au {s['Date_Actualisation']:%d/%m/%Y}"
                          if s and s["Date_Actualisation"] else "Aucun suivi enregistré pour cette activité."),
            "structures": " ".join(structures_de(a["Region"], structures)),
            "type": a["Categorie"],
        })
    _libelles_distincts(rows, {a["CODE_ACTIVITE"]: a for a in ref.activities()})
    sans = [r["name"] for r in rows if not r["structures"]]
    if sans:
        raise StructuresError(f"Aucune structure ne couvre ces activités (elles seraient introuvables dans le formulaire) : "
                              f"{', '.join(sans)}. Ajoutez la structure de leur région dans kobo/structures.csv.")
    return rows


def types_csv_rows(ref: Referentiel, rows: list[dict], structures: list[dict]) -> list[dict]:
    """Types d'activités (catégories de config.yaml) ayant au moins une activité active ; « structures » = celles qui les voient."""
    ordre = [x["name"] for x in structures]
    out = []
    for code, libelle in ref.cfg.categories.items():
        sts = {st for r in rows if r["type"] == code for st in r["structures"].split()}
        if sts:
            out.append({"name": code, "label": _clean(libelle), "structures": " ".join(x for x in ordre if x in sts)})
    return out


def _libelles_distincts(rows: list[dict], acts: dict) -> None:
    """Les agents ne voient pas les codes : deux activités au même libellé (même nature, région, département, entreprise)
    sont distinguées par leurs premières localités, sinon par le montant, sinon par un numéro de lot."""
    def doublons():
        vus = {}
        for r in rows:
            vus.setdefault(r["label"], []).append(r)
        return [g for g in vus.values() if len(g) > 1]
    for groupe in doublons():
        for r in groupe:
            loc = [x.strip().title() for x in re.split(r"[,;/]", _clean(acts[r["name"]]["Localite"])) if x.strip()]
            if loc:
                r["label"] += f" · {' / '.join(loc[:2])}{'…' if len(loc) > 2 else ''}"      # pas de virgule dans le CSV
    for groupe in doublons():
        for r in groupe:
            m = acts[r["name"]]["Montant_Marche"]
            if m:
                r["label"] += f" · {_montant(m)} FCFA"
    for groupe in doublons():
        for n, r in enumerate(sorted(groupe, key=lambda x: x["name"]), 1):
            r["label"] += f" · lot {n}"


def csv_bytes(rows: list[dict], columns: list[str] = CSV_COLUMNS) -> bytes:
    buf = io.StringIO(newline="")
    w = csv.DictWriter(buf, fieldnames=columns, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue().encode("utf-8")          # UTF-8 SANS BOM : un BOM casse la colonne « name » côté Kobo Collect


def write_activities_csv(ref: Referentiel, path) -> int:
    """Écrit activites.csv et, à côté, structures.csv (les deux fichiers médias du formulaire)."""
    structures = load_structures(ref.cfg)
    rows = activities_csv_rows(ref, structures)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(csv_bytes(rows))
    (path.parent / STRUCTURES_NAME).write_bytes(csv_bytes(structures, STRUCTURES_COLUMNS))
    (path.parent / TYPES_NAME).write_bytes(csv_bytes(types_csv_rows(ref, rows, structures), TYPES_COLUMNS))
    return len(rows)


# --------------------------------------------------------------------------
# Formulaire XLSForm
# --------------------------------------------------------------------------
SURVEY_HEADER = ["type", "name", "label", "hint", "required", "required_message", "constraint", "constraint_message", "appearance",
                 "choice_filter", "calculation", "relevant", "default", "parameters"]
H = lambda **k: [k.get(c, "") for c in SURVEY_HEADER]   # noqa: E731
SURVEY = [
    H(type="start", name="start"),
    H(type="end", name="end"),
    H(type="begin_group", name="identification", label="Identification"),
    H(type="text", name="nom_agent", label="Nom et prénoms", required="yes", constraint="string-length(normalize-space(.)) >= 3",
      constraint_message="Saisissez votre nom et vos prénoms."),
    H(type="select_one_from_file " + STRUCTURES_NAME, name="structure", label="Structure",
      hint="Choisissez votre structure : seules ses activités seront proposées", required="yes", appearance="minimal"),
    H(type="select_one_from_file " + TYPES_NAME, name="type_activite", label="Type d'activité",
      hint="Types d'activités suivis par votre structure", required="yes", appearance="minimal",
      choice_filter=FILTRE_TYPES, relevant="${structure} != ''"),
    H(type="date", name="date_actualisation", label="Date de l'actualisation", hint="Date de la visite ou de la situation décrite", required="yes",
      constraint=". <= today()", constraint_message="La date ne peut pas être dans le futur.", default="today()"),
    H(type="select_one_from_file " + CSV_NAME, name="code_activite", label="Activité suivie",
      hint="Activités de ce type pour votre structure ; tapez le département ou l'entreprise pour filtrer", required="yes",
      appearance="autocomplete", choice_filter=FILTRE_ACTIVITES, relevant="${type_activite} != ''",
      constraint=(f"selected(pulldata('{CSV_NAME[:-4]}', 'structures', 'name', .), ${{structure}}) "
                  f"and pulldata('{CSV_NAME[:-4]}', 'type', 'name', .) = ${{type_activite}}"),
      constraint_message="Cette activité ne correspond pas à la structure ou au type choisi : choisissez à nouveau l'activité."),
    H(type="end_group"),
    H(type="calculate", name="ref_titre", calculation=f"pulldata('{CSV_NAME[:-4]}', 'ref_titre', 'name', ${{code_activite}})"),
    H(type="calculate", name="ref_marche", calculation=f"pulldata('{CSV_NAME[:-4]}', 'ref_marche', 'name', ${{code_activite}})"),
    H(type="calculate", name="ref_suivi", calculation=f"pulldata('{CSV_NAME[:-4]}', 'ref_suivi', 'name', ${{code_activite}})"),
    H(type="note", name="note_reference", label="**${ref_titre}**\n${ref_marche}\n${ref_suivi}", relevant="${code_activite} != ''"),
    H(type="begin_group", name="avancement", label="Etat d'avancement"),
    H(type="integer", name="taux_execution", label="Taux d'exécution physique actuel (%)", hint="Nombre entier de 0 à 100", required="yes",
      constraint=". >= 0 and . <= 100", constraint_message="Le taux doit être compris entre 0 et 100."),
    H(type="text", name="observations_techniques", label="Observations techniques",
      hint="Travaux réalisés, activités exécutées, état d'avancement, constats techniques. Un constat par ligne. "
           "Ne pas y mettre les difficultés ni les recommandations.", required="yes", appearance="multiline"),
    H(type="text", name="commentaires", label="Commentaires généraux", appearance="multiline"),
    H(type="text", name="difficultes", label="Difficultés rencontrées", hint="Laisser vide s'il n'y en a pas", appearance="multiline"),
    H(type="text", name="recommandations", label="Recommandations", appearance="multiline"),
    *[H(type="image", name=nom, label=f"Photo du site {i}/4", required="yes", parameters="max-pixels=1600")
      for i, nom in enumerate(("Photo_du_site", "Photo_du_site_001", "Photo_du_site_002", "Photo_du_site_003"), 1)],
    H(type="end_group"),
]


def build_xlsform(path, version: str | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "survey"
    ws.append(SURVEY_HEADER)
    for row in SURVEY:
        ws.append(row)
    st = wb.create_sheet("settings")
    st.append(["form_title", "form_id", "version", "default_language", "instance_name"])
    st.append(["Suivi périodique des activités PCR-CI", FORM_ID, version or datetime.now().strftime("%Y%m%d%H%M"), "French (fr)",
               "concat(${code_activite}, ' - ', format-date(${date_actualisation}, '%d/%m/%Y'))"])
    for sheet, widths in ((ws, (30, 24, 50, 50, 10, 20, 38, 38, 14, 36, 60, 24, 10, 18)), (st, (40, 24, 16, 18, 70))):
        for i, w in enumerate(widths):
            sheet.column_dimensions[chr(65 + i)].width = w
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------
# Synchronisation avec Kobo (API v2)
# --------------------------------------------------------------------------
class KoboError(RuntimeError):
    pass


def _hint(code):
    return {401: "jeton KOBO_TOKEN invalide ou expiré", 403: "ce jeton n'a pas le droit de modifier ce projet",
            404: "KOBO_ASSET_UID ou KOBO_URL incorrect (serveur kf.kobotoolbox.org ?)"}.get(code, "réponse inattendue du serveur Kobo")


class KoboClient:
    def __init__(self, base_url: str, asset_uid: str, token: str, session=None):
        self.base, self.uid = base_url.rstrip("/"), asset_uid
        self.s = session or self._session()
        self.s.headers.update({"Authorization": f"Token {token}"})

    @staticmethod
    def _session():
        s = requests.Session()
        s.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                                                           allowed_methods=("GET",), respect_retry_after_header=True)))
        return s

    def _call(self, method, path, ok=(200, 201, 204), **kw):
        url = f"{self.base}/api/v2/assets/{self.uid}/{path}"
        try:
            r = self.s.request(method, url, timeout=60, **kw)
        except requests.RequestException as e:
            raise KoboError(f"Kobo injoignable : {e}") from e
        if r.status_code not in ok:
            raise KoboError(f"Kobo a répondu HTTP {r.status_code} ({method} {path}) : {_hint(r.status_code)}. {r.text[:200]}")
        return r

    def media(self) -> list[dict]:
        return self._call("GET", "files/?format=json").json().get("results", [])

    def media_named(self, filename: str) -> list[dict]:
        return [m for m in self.media() if (m.get("metadata") or {}).get("filename") == filename and m.get("file_type") == "form_media"]

    def _upload(self, filename: str, content: bytes, ok=(201,)):
        return self._call("POST", "files/", ok=ok, data={"description": "default", "file_type": "form_media", "metadata": json.dumps({"filename": filename})},
                          files={"content": (filename, content, "text/csv")})

    def replace_media(self, filename: str, content: bytes) -> str:
        """Envoie d'abord la nouvelle version ; si Kobo refuse le doublon, supprime l'ancienne puis renvoie.
        Retourne « ajouté » ou « remplacé »."""
        old = self.media_named(filename)
        r = self._upload(filename, content, ok=(201, 400, 409))
        if r.status_code == 201:
            for m in old:
                self._call("DELETE", f"files/{m['uid']}/")
            return "remplacé" if old else "ajouté"
        for m in old:
            self._call("DELETE", f"files/{m['uid']}/")
        self._upload(filename, content)
        return "remplacé"

    def submissions(self) -> list[dict]:
        """Toutes les soumissions du projet (pagination suivie jusqu'au bout)."""
        out, url = [], f"{self.base}/api/v2/assets/{self.uid}/data/?format=json&limit=1000"
        while url:
            try:
                r = self.s.get(url, timeout=60)
            except requests.RequestException as e:
                raise KoboError(f"Kobo injoignable : {e}") from e
            if r.status_code != 200:
                raise KoboError(f"Récupération des soumissions impossible (HTTP {r.status_code}) : {_hint(r.status_code)}")
            payload = r.json()
            out.extend(payload.get("results", []))
            url = payload.get("next")
        return out

    def redeploy(self) -> str:
        asset = self.s.get(f"{self.base}/api/v2/assets/{self.uid}/?format=json", timeout=60)
        if asset.status_code != 200:
            raise KoboError(f"Lecture du projet impossible (HTTP {asset.status_code}) : {_hint(asset.status_code)}")
        version = asset.json().get("version_id")
        if not version:
            raise KoboError("Le projet n'a pas de version déployée : déployez d'abord le formulaire dans Kobo.")
        self._call("PATCH", "deployment/", ok=(200,), json={"active": True, "version_id": version})
        return version


def sync_to_kobo(client: KoboClient, rows: list[dict], redeploy: bool = True, structures: list[dict] | None = None,
                 types: list[dict] | None = None) -> dict:
    """Envoie structures.csv et types_activite.csv (si fournis) puis activites.csv, et redéploie si demandé."""
    if not rows:
        raise KoboError("Aucune activité active à envoyer : envoi annulé pour ne pas vider la liste des techniciens.")
    action_s = client.replace_media(STRUCTURES_NAME, csv_bytes(structures, STRUCTURES_COLUMNS)) if structures else None
    if types:
        client.replace_media(TYPES_NAME, csv_bytes(types, TYPES_COLUMNS))
    action = client.replace_media(CSV_NAME, csv_bytes(rows))
    version = client.redeploy() if redeploy else None
    return {"activites": len(rows), "fichier": action, "redeploye": bool(redeploy), "version": version,
            "structures": len(structures or []), "fichier_structures": action_s, "types": len(types or [])}
