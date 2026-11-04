"""Chargement initial du référentiel à partir des matrices Word (source la plus à jour).

Principes :
  - les informations viennent du Word ; l'ancien classeur Excel ne sert qu'à récupérer les LOCALITÉS des blocs sanitaires
    (absentes du Word), et seulement quand Word et Excel concordent (montant ET département) ;
  - rien n'est inventé : une information absente reste vide, une date incohérente n'alimente aucun calcul ;
  - chaque anomalie est listée dans un rapport à valider avant de figer les codes.
"""
import re
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from .calculs import fin_previsionnelle
from .config import normalize
from .referentiel import Referentiel

LABELS = r"(Intitulé|Montant marché|Titulaire|O\.S\.|OS|Date de démarrage|Délai|Fin prévue|Bureau de contrôle)"
USER = "IMPORT_INITIAL"


def parse_date_loose(s: str):
    """(date | None, remarque). Tolère « 17/012/2026 », « 6/01/2027 » ; refuse les dates impossibles."""
    m = re.search(r"(\d{1,2})/(\d{1,3})/(\d{4})", s or "")
    if not m:
        return None, None
    note = f"« {m[0]} » corrigée" if len(m[2]) > 2 else None
    try:
        return date(int(m[3]), int(m[2]), int(m[1])), note
    except ValueError:
        return None, f"date impossible « {m[0]} »"


def split_cell(text: str) -> dict:
    text = re.sub(r"\s+", " ", re.sub(r"\*+", "", text)).strip()
    parts = re.split(r"(?<![A-Za-zÀ-ÿ])" + LABELS + r"\s*:", text)
    out = {"_region": parts[0].strip()}
    for label, value in zip(parts[1::2], parts[2::2]):
        out["OS" if label in ("O.S.", "OS") else label] = value.strip()
    return out


def parse_word(path, categorie: str) -> list[dict]:
    acts = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if not line.startswith("| ") or "Intitulé" not in line.split(" | ")[0]:
            continue
        f = split_cell(line.strip().strip("|").split(" | ")[0])
        intit = f.get("Intitulé", "")
        flags = []
        if categorie == "BS":
            m = re.search(r"(?:du|des) départements? (?:d['’]|de |du |des )?(.+)$", intit)
            dep, loc = (m[1].replace(" et ", " / ").strip() if m else None), None
        else:
            m = re.search(r"localités? (?:d['’]|de |du |des )?(.+?) dans les? départements? (?:d['’]|de |du |des )?(.+)$", intit)
            dep = m[2].replace(" et ", " / ").strip() if m else None
            loc = m[1].replace(" et ", ", ").strip() if m else None
        if not m:
            flags.append("département/localités non extraits de l'intitulé")
        montant = re.search(r"\d[\d ]{4,}\d", f.get("Montant marché", ""))
        os_date, n1 = parse_date_loose(f.get("OS", ""))
        dem_date, n2 = parse_date_loose(f.get("Date de démarrage", ""))
        fin_word, n3 = parse_date_loose(f.get("Fin prévue", ""))
        flags += [n for n in (n1, n2, n3) if n]
        delai = re.search(r"\d+", f.get("Délai", ""))
        acts.append({
            "categorie": categorie, "region_txt": f["_region"], "Departement": dep, "Localite": loc, "Designation": intit or None,
            "Montant_Marche": int(montant[0].replace(" ", "")) if montant else None, "Entreprise": f.get("Titulaire") or None,
            "Date_OS": os_date, "Delai_Mois": int(delai[0]) if delai else None, "Bureau_Controle": f.get("Bureau de contrôle") or None,
            # BS : le Word calcule le délai depuis l'O.S. ; MR : depuis la date de démarrage effectif
            "Date_Debut": os_date if categorie == "BS" else dem_date, "_fin_word": fin_word, "_flags": flags,
        })
    return acts


def _dep_tokens(s) -> frozenset:
    return frozenset(normalize(s).split()) - {"ET", "DE", "DU", "DES"}


def load_excel_localites(path) -> list[dict]:
    """Lecture seule de l'ancienne matrice : lots, départements, montants, localités des blocs sanitaires."""
    ws = load_workbook(path, data_only=True)["MATRICE DE SUIVI"]
    out = []
    for r in range(4, ws.max_row + 1):
        lot = ws.cell(r, 2).value
        if not isinstance(lot, (int, float)):
            continue
        m = re.search(r"\d[\d\s]{4,}\d", str(ws.cell(r, 6).value or ""))
        loc = [x.strip() for x in str(ws.cell(r, 10).value or "").splitlines() if x.strip()]
        out.append({"lot": int(lot), "dep": ws.cell(r, 4).value, "montant": int(re.sub(r"\s", "", m[0])) if m else None,
                    "entreprise": ws.cell(r, 5).value, "localites": ", ".join(loc)})
    return out


def seed(ref: Referentiel, sources: list[tuple[str, Path]], excel_path=None) -> list[dict]:
    """Crée toutes les activités. Retourne une liste de lignes de rapport (une par activité)."""
    if ref.activities():
        raise RuntimeError("Le référentiel contient déjà des activités : le chargement initial ne s'exécute qu'une fois.")
    excel = load_excel_localites(excel_path) if excel_path else []
    parsed = [(cat, a) for cat, path in sources for a in parse_word(path, cat)]
    montants = {}
    for _, a in parsed:
        if a["Montant_Marche"]:
            montants.setdefault(a["Montant_Marche"], []).append(a)
    tol, report = ref.cfg.tolerance_jours, []
    for cat, a in parsed:
        flags = list(a["_flags"])
        # montant identique sur plusieurs marchés
        if a["Montant_Marche"] and len(montants[a["Montant_Marche"]]) > 1:
            autres = [x["Departement"] for x in montants[a["Montant_Marche"]] if x is not a]
            flags.append(f"montant identique à celui de : {', '.join(map(str, autres))} (erreur de copie possible)")
        # fin prévue du Word vs début + délai en mois calendaires
        incoherent = False
        debut, calc = a["Date_Debut"], fin_previsionnelle(a["Date_Debut"], a["Delai_Mois"])
        ecart = (calc - a["_fin_word"]).days if calc and a["_fin_word"] else None
        if ecart is not None and abs(ecart) > tol:
            flags.append(f"INCOHÉRENT : fin prévue du Word {a['_fin_word']:%d/%m/%Y} ≠ début {debut:%d/%m/%Y} + {a['Delai_Mois']} mois "
                         f"({calc:%d/%m/%Y}) : écart de {ecart} j. Date de début non enregistrée (aucun indicateur de délai tant que non confirmée).")
            a["Date_Debut"] = None
            incoherent = True
        elif ecart:
            flags.append(f"écart de {ecart:+d} j entre la fin prévue du Word et le calcul en mois calendaires (fin calculée : {calc:%d/%m/%Y})")
        if a["Date_Debut"] is None and not incoherent:
            flags.append("début des travaux absent du Word : aucun indicateur de délai tant qu'il n'est pas renseigné")
        # localités des BS : depuis l'Excel seulement si montant ET département concordent
        if cat == "BS" and excel:
            same_amount = [e for e in excel if e["montant"] == a["Montant_Marche"]]
            ok = [e for e in same_amount if a["Departement"] and _dep_tokens(e["dep"]) == _dep_tokens(a["Departement"])]
            if len(ok) == 1 and ok[0]["localites"]:
                a["Localite"] = ok[0]["localites"]
            elif same_amount and not ok:
                flags.append(f"localités non reprises : l'Excel place ce montant à « {same_amount[0]['dep']} » (Word : « {a['Departement']} »)")
            else:
                flags.append("localités non reprises : aucune ligne de l'Excel ne correspond (montant différent)")
        missing = [k for k in ("Entreprise", "Montant_Marche", "Date_OS", "Delai_Mois") if not a[k]]
        if missing:
            flags.append("informations absentes du Word : " + ", ".join(missing))
        code = ref.create_activity(cat, a["region_txt"], USER, **{k: a[k] for k in (
            "Departement", "Localite", "Designation", "Montant_Marche", "Entreprise", "Date_OS", "Date_Debut", "Delai_Mois", "Bureau_Controle")})
        report.append({"code": code, "categorie": cat, "region": a["region_txt"], "departement": a["Departement"],
                       "entreprise": a["Entreprise"], "montant": a["Montant_Marche"], "flags": flags})
    return report


def write_report(rows: list[dict], path, problems: list[str]) -> None:
    n_flag = sum(1 for r in rows if r["flags"])
    lines = ["# Rapport du chargement initial", "",
             f"{len(rows)} activités créées ({sum(r['categorie'] == 'BS' for r in rows)} blocs sanitaires, "
             f"{sum(r['categorie'] == 'MR' for r in rows)} marchés ruraux). {n_flag} avec au moins un point à vérifier.",
             f"Contrôle d'intégrité du référentiel : {'conforme' if not problems else 'PROBLÈMES : ' + '; '.join(problems)}.", "",
             "Numéro d'ordre de service : absent des matrices Word, donc vide pour toutes les activités (à compléter dans le fichier d'administration).", "",
             "## Codes attribués", "", "| Code | Région | Département | Entreprise | Montant (FCFA) |", "|---|---|---|---|---|"]
    for r in rows:
        mt = f"{r['montant']:,}".replace(",", " ") if r["montant"] else "—"
        lines.append(f"| {r['code']} | {r['region']} | {r['departement'] or '—'} | {r['entreprise'] or '—'} | {mt} |")
    lines += ["", "## Points à vérifier (par activité)", ""]
    for r in rows:
        if r["flags"]:
            lines.append(f"**{r['code']}** ({r['departement'] or r['region']}, {r['entreprise'] or 'titulaire non renseigné'})")
            lines += [f"- {f}" for f in r["flags"]] + [""]
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text("\n".join(lines), encoding="utf-8")
