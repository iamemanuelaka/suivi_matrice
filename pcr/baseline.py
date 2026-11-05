"""Import de l'état d'avancement actuel des matrices Word comme « suivi de référence » (une ligne par activité).

Sans lui, les matrices regénérées depuis la base seraient vides tant que les techniciens n'ont pas soumis de suivi.
Chaque ligne est marquée Source = BASELINE_WORD et reste dans l'historique comme n'importe quel suivi.
"""
import re
from datetime import date, datetime
from pathlib import Path

from .referentiel import Referentiel
from .seed import parse_date_loose, split_cell

BASELINE_ID = "INIT-{code}"


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"\*+", "", s or "")).strip()


def num(s):
    m = re.search(r"\d+(?:[.,]\d+)?", s or "")
    if not m:
        return None
    v = float(m[0].replace(",", "."))
    return int(v) if v.is_integer() else v


def _cells(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def parse_cell3(text: str) -> dict:
    t = clean(text)
    comm = re.split(r"Difficultés\s*:|Recommandations\s*:", t)[0].strip()
    d = re.search(r"Difficultés\s*:\s*(.*?)(?=Recommandations\s*:|Difficultés\s*:|$)", t)
    r = re.search(r"Recommandations\s*:\s*(.*?)(?=Difficultés\s*:|Recommandations\s*:|$)", t)
    return {"commentaires": comm or None, "difficultes": clean(d[1]) if d and d[1].strip() else None,
            "recommandations": clean(r[1]) if r and r[1].strip() else None,
            "duplique": len(re.findall(r"Difficultés\s*:", t)) > 1}


def parse_suivis(path, categorie: str) -> list[dict]:
    """Une entrée par activité, dans l'ordre du Word (même ordre que le chargement initial)."""
    out = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if not line.startswith("|"):
            continue
        cells = _cells(line)
        if len(cells) < 3:
            continue
        if "Intitulé" in cells[0]:
            c2, c3 = clean(cells[1]), parse_cell3(cells[2])
            cut = re.search(r"Taux d.avancement\s*:|Total factur", c2)
            obs = c2[: cut.start()].strip() if cut else c2
            taux = num(re.search(r"Taux d.avancement\s*:\s*([\d.,]+)", c2)[1]) if re.search(r"Taux d.avancement\s*:\s*[\d.,]", c2) else None
            tf = re.search(r"Total factur\w*\s*:?\s*([\d.,]+)", c2)
            vis = re.search(r"Visite effectuée(?: le)?\s*(\d{1,2}\s*/\s*\d{1,3}\s*/\s*\d{4})", obs)
            out.append({"region": split_cell(cells[0])["_region"], "observations": obs or None, "taux": taux,
                        "total_facture": num(tf[1]) if tf else None, "date_visite": parse_date_loose(vis[1].replace(" ", ""))[0] if vis else None,
                        "delai_word": num(re.search(r"Délai consommé\s*:\s*([\d.,]+)", c2)[1]) if re.search(r"Délai consommé\s*:\s*[\d.,]", c2) else None,
                        **c3})
        elif out and not cells[0] and categorie == "MR":                  # lignes « Taux » et « Délai consommé » des marchés ruraux
            c2 = clean(cells[1])
            if re.search(r"Taux d.avancement", c2) and re.search(r"\d", c2):
                out[-1]["taux"] = num(c2.split(":", 1)[1])
            elif re.search(r"Délai consommé", c2) and re.search(r"\d", c2):
                out[-1]["delai_word"] = num(c2.split(":", 1)[1])
    return out


def import_baseline(ref: Referentiel, sources: list[tuple[str, Path]], date_defaut: date) -> list[str]:
    """Insère le suivi de référence de chaque activité. Idempotent (une ligne INIT-<code> au plus). Retourne les avertissements."""
    existing = {str(r["ID_SOUMISSION_KOBO"]) for r in ref.base.rows("Suivi_Activites")}
    acts = ref.activities()
    warnings = []
    for cat, path in sources:
        counters: dict[str, int] = {}
        for s in parse_suivis(path, cat):
            reg = ref.cfg.region_code(s["region"])
            counters[reg] = counters.get(reg, 0) + 1
            code = f"{cat}{reg}{counters[reg]:03d}"
            if not any(a["CODE_ACTIVITE"] == code for a in acts):
                warnings.append(f"{code} : activité introuvable dans le référentiel, suivi ignoré")
                continue
            if s["duplique"]:
                warnings.append(f"{code} : le Word répète deux fois « Difficultés / Recommandations » ; seule la première occurrence est reprise")
            sid = BASELINE_ID.format(code=code)
            if sid in existing or (s["observations"] is None and s["taux"] is None):
                continue
            ref.base.append("Suivi_Activites", {
                "ID_SOUMISSION_KOBO": sid, "CODE_ACTIVITE": code, "Date_Actualisation": s["date_visite"] or date_defaut,
                "Observations_Techniques": s["observations"], "Taux_Execution_Physique": s["taux"], "Commentaires": s["commentaires"],
                "Difficultes": s["difficultes"], "Recommandations": s["recommandations"], "Date_Importation": datetime.now().replace(microsecond=0),
                "Statut_Traitement": "VALIDE", "Nom_Agent": "Import matrice Word", "Total_Facture": s["total_facture"], "Source": "BASELINE_WORD"})
    return warnings
