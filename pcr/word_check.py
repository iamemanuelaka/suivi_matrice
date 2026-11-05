"""Contrôle de fidélité : compare chaque cellule des Word générés au texte des matrices d'origine.

La comparaison ignore le bruit de mise en forme (étoiles, espaces, virgules de fin, « 5 % » / « 5% », dates mal saisies comme 17/012/2026,
« FCFA HT ») mais PAS le fond. Le « délai consommé » est comparé à part : il est recalculé à la date de situation.
"""
import difflib
import re
from datetime import date
from pathlib import Path

import docx

from .baseline import clean
from .seed import parse_date_loose

DELAI = re.compile(r"Délai consommé\s*:?\s*([\d.,]*)\s*%?")


def _norm_dates(t: str) -> str:
    def fix(m):
        d, _ = parse_date_loose(m[0].replace(" ", ""))
        return f"{d:%d/%m/%Y}" if d else m[0]
    return re.sub(r"\d{1,2}\s*/\s*\d{1,3}\s*/\s*\d{4}", fix, t)


def norm(t: str) -> str:
    t = clean(t).replace("’", "'")
    t = _norm_dates(t)
    t = re.sub(r"Total factur\w*\s*:?", "Total facturé :", t)
    t = re.sub(r"\bO\.?S\.?\s*:", "O.S. :", t)
    t = re.sub(r"\s*FCFA\s*HT", "", t)
    t = re.sub(r"(\d)\s+%", r"\1%", t)
    t = DELAI.sub("Délai consommé :", t)
    t = re.sub(r"\s*,\s*", " ", t)
    t = re.sub(r"\s*:\s*", " : ", t)
    return re.sub(r"\s+", " ", t).strip()


def delai_value(t: str):
    m = DELAI.search(clean(t))
    return float(m[1].replace(",", ".")) if m and m[1] else None


def source_rows(path, cat) -> list[list[str]]:
    rows = []
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 3:
            continue
        if "Intitulé" in cells[0]:
            rows.append([cells[0], cells[1], cells[2]])
        elif rows and cat == "MR" and not cells[0]:
            rows[-1][1] += " " + cells[1]
    return rows


def generated_rows(path, cat) -> list[list[str]]:
    t = docx.Document(path).tables[0]
    body, step = t.rows[1:], 3 if cat == "MR" else 1
    out = []
    for i in range(0, len(body), step):
        grp = body[i:i + step]
        c = [" ".join(p.text for p in cell.paragraphs) for cell in grp[0].cells]
        for extra in grp[1:]:
            c[1] += " " + " ".join(p.text for p in extra.cells[1].paragraphs)
        out.append(c)
    return out


def word_diff(a: str, b: str) -> str:
    x, y = a.split(), b.split()
    parts = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, x, y, autojunk=False).get_opcodes():
        if op != "equal":
            parts.append(f"source « {' '.join(x[i1:i2]) or '∅'} » → généré « {' '.join(y[j1:j2]) or '∅'} »")
    return " ; ".join(parts)


def compare(ref, sources: dict, generated: dict) -> str:
    """sources/generated : {catégorie: chemin}. Retourne le rapport Markdown."""
    L = ["# Contrôle de fidélité des matrices Word générées", "",
         "Chaque cellule générée est comparée à la cellule correspondante des matrices fournies (même ordre de lignes). "
         "Sont ignorés : étoiles de mise en forme, espaces, virgules, « 5 % » / « 5% », dates mal saisies (17/012/2026), « FCFA HT ». "
         "Le fond du texte, les chiffres, les dates et l'ordre des lignes sont comparés.", ""]
    detail, delais = [], []
    for cat in sources:
        s, g = source_rows(sources[cat], cat), generated_rows(generated[cat], cat)
        name = ref.cfg.categories[cat]
        L.append(f"## {name}")
        if len(s) != len(g):
            L += [f"**Nombre de lignes différent : source {len(s)}, généré {len(g)}**", ""]
        ok = [0, 0, 0]
        for i, (rs, rg) in enumerate(zip(s, g), 1):
            label = f"{cat} ligne {i} ({re.sub(r'[^A-Z]', '', clean(rs[0]).split(' ')[0])})"
            for c, cname in enumerate(("Désignation", "Avancement", "Commentaires")):
                a, b = norm(rs[c]), norm(rg[c])
                if a == b:
                    ok[c] += 1
                else:
                    detail.append(f"- **{label} – {cname}** : {word_diff(a, b)}")
            da, db = delai_value(rs[1]), delai_value(rg[1])
            if da is not None or db is not None:
                delais.append((label, da, db))
        n = min(len(s), len(g))
        L += ["| Colonne | Cellules identiques |", "|---|---|"] + [f"| {c} | {ok[k]} / {n} |" for k, c in enumerate(("Désignation", "Avancement", "Commentaires"))] + [""]
    L += ["## Différences de contenu restantes", ""] + (detail or ["Aucune."]) + [""]
    if detail:
        L += ["**Causes possibles** : (1) date « source → ∅ » : début ou fin prévue non enregistré parce que la donnée du Word est incohérente ou absente "
              "(voir rapport_chargement_initial.md) : à confirmer dans le fichier d'administration, jamais inventé ; (2) bloc « Difficultés / Recommandations » "
              "répété deux fois dans le Word d'origine : une seule occurrence est reprise ; (3) date de fin différente de 2 jours : la fin prévue du Word ne suit "
              "pas la convention de calcul du projet ; (4) libellés vides en plus : marché sans suivi, affiché avec les trois lignes habituelles.", ""]
    diffs = [(l, a, b) for l, a, b in delais if a is None or b is None or abs(a - b) > 0.005]
    L += ["## Délai consommé (recalculé)", "",
          f"{len(delais) - len(diffs)} valeur(s) identique(s) sur {len(delais)}. Le délai consommé du Word n'a pas été calculé à une date unique "
          "(ex. Koro : 54,09 % = situation au 22/10 ; Touba : 35,24 % = 29/09). Il est maintenant calculé pour la date de situation du document.", "",
          "| Ligne | Word | Généré |", "|---|---|---|"] + [f"| {l} | {'' if a is None else a} | {'' if b is None else b} |" for l, a, b in diffs]
    return "\n".join(L)
