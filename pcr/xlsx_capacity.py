"""Extension de la capacité du tableau de bord (45 -> N activités) SANS le reconstruire.

Principe : le modèle lit les activités par formules (feuille technique CALCULS : une ligne de formules par activité, plages
`$5:$49` dans les formules et les noms définis). L'extension CLONE ces formules (traduites ligne par ligne, comme le ferait une
recopie vers le bas dans Excel) et agrandit les plages. Rien n'est réécrit à la main : la logique du modèle (statuts, PGES, alertes,
filtres, listes déroulantes) est exactement celle du modèle d'origine.

Autres ajustements :
* « DÉTAIL DES ALERTES » : la liste (table tblAlertes) est agrandie pour pouvoir contenir toutes les alertes ; la note de règle est
  déplacée sous la table.
* « MATRICE DE SUIVI » : fusions de la colonne région supprimées (la région est écrite sur chaque ligne ; l'ajout d'une activité ne
  dépend plus de la mise en page) ; la feuille est MASQUÉE : c'est désormais une feuille technique alimentée depuis la base.
"""
from __future__ import annotations

import copy
import re
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.formula.translate import Translator

from .xlsx_patch import Sheet, col_to_idx, idx_to_col
from .xlsx_repair import lire_parts

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
Q = lambda t: f"{{{S}}}{t}"          # noqa: E731
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

LISTES_DISTINCTES = ["AJ", "AL", "AN", "AP", "AR"]      # listes de valeurs distinctes (région, département, entreprise, nature, année)
LISTES_DEROULANTES = ["CQ", "CU", "CY", "DC"]           # listes des filtres du tableau de bord
CONSTANTES = {"A": lambda r: r - 4, "CM": lambda r: r - 5}
IGNOREES = {"BJ", "BK"}                                  # valeurs constantes de mise en page, sans lien avec les activités
RANGE_FIN = re.compile(r"(\$?[A-Z]{1,3}\$5:\$?[A-Z]{1,3}\$)(\d+)\b")


def capacite_modele(parts_or_path) -> int:
    """Nombre d'activités gérées = plus grande valeur de la colonne A de CALCULS (compteur de lignes)."""
    parts = parts_or_path if isinstance(parts_or_path, dict) else lire_parts(parts_or_path)
    fichier = _fichier_feuille(parts, "CALCULS")
    xml = parts[fichier].decode("utf-8")
    valeurs = [int(v) for v in re.findall(r'<c r="A\d+"[^>]*><v>(\d+)</v>', xml)]
    return max(valeurs) if valeurs else 0


def _fichier_feuille(parts, nom):
    wb = etree.fromstring(parts["xl/workbook.xml"])
    rels = etree.fromstring(parts["xl/_rels/workbook.xml.rels"])
    cible = {r.get("Id"): r.get("Target") for r in rels}
    for s in wb.find(Q("sheets")):
        if s.get("name") == nom:
            t = cible[s.get(f"{{{RNS}}}id")].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise KeyError(nom)


def _traduire(formule: str, origine: str, cible: str) -> str:
    return Translator("=" + formule, origin=origine).translate_formula(cible)[1:]


def _agrandir_plages(formule: str, ancienne_fin: int, nouvelle_fin: int) -> str:
    return RANGE_FIN.sub(lambda m: m.group(1) + (str(nouvelle_fin) if int(m.group(2)) == ancienne_fin else m.group(2)), formule)


def _set_formule(sh: Sheet, r: int, col: str, formule: str, style):
    cell = sh._cell(r, col_to_idx(col), create=True)
    for ch in list(cell):
        cell.remove(ch)
    if "t" in cell.attrib:
        del cell.attrib["t"]
    if style is not None:
        cell.set("s", style)
    etree.SubElement(cell, Q("f")).text = formule
    etree.SubElement(cell, Q("v"))
    sh.dirty = True


def _set_constante(sh: Sheet, r: int, col: str, valeur: int, style):
    cell = sh._cell(r, col_to_idx(col), create=True)
    for ch in list(cell):
        cell.remove(ch)
    cell.set("t", "n")
    if style is not None:
        cell.set("s", style)
    etree.SubElement(cell, Q("v")).text = str(valeur)
    sh.dirty = True


def _etendre_calculs(xml: bytes, ancienne_cap: int, cap: int) -> bytes:
    sh = Sheet(xml)
    vieille_fin, fin = 4 + ancienne_cap, 4 + cap                    # dernière ligne d'activité (49 -> 504)
    # 1. toutes les formules : plages agrandies
    for f in sh.root.iter(Q("f")):
        if f.text:
            f.text = _agrandir_plages(f.text, vieille_fin, fin)
    # 2. colonnes « une formule par activité » (modèle = ligne 49)
    cellules = {}
    for c in sh.rows[vieille_fin].findall(Q("c")):
        col = re.match(r"[A-Z]+", c.get("r")).group(0)
        f = c.find(Q("f"))
        cellules[col] = (f.text if f is not None else None, c.get("s"))
    par_activite = [c for c, (f, _) in cellules.items()
                    if f and c not in LISTES_DISTINCTES + LISTES_DEROULANTES and c not in CONSTANTES and c not in IGNOREES]
    for r in range(vieille_fin + 1, fin + 1):
        for col in par_activite:
            f, st = cellules[col]
            _set_formule(sh, r, col, _traduire(f, f"{col}{vieille_fin}", f"{col}{r}"), st)
        for col, fn in CONSTANTES.items():
            if col == "A":
                _set_constante(sh, r, col, fn(r), cellules[col][1])
    # 3. listes de valeurs distinctes : MATCH(k, …) avec k = n° de ligne - 5, jusqu'à fin + 5
    derniere_liste = vieille_fin + 5
    for col in LISTES_DISTINCTES:
        modele = sh.rows[derniere_liste].find(f"./{Q('c')}[@r='{col}{derniere_liste}']").find(Q("f")).text
        st = cellules[col][1]
        for r in range(derniere_liste + 1, fin + 5 + 1):
            _set_formule(sh, r, col, re.sub(r"MATCH\(\d+,", f"MATCH({r - 5},", modele, count=1), st)
    # 4. listes déroulantes (CM = rang constant) : lignes 50 -> fin + 1
    derniere = vieille_fin + 1
    for r in range(derniere + 1, fin + 1 + 1):
        _set_constante(sh, r, "CM", r - 5, cellules["CM"][1])
        for col in LISTES_DEROULANTES:
            base = sh.rows[derniere].find(f"./{Q('c')}[@r='{col}{derniere}']")
            _set_formule(sh, r, col, _traduire(base.find(Q("f")).text, f"{col}{derniere}", f"{col}{r}"), base.get("s"))
    # 5. dimension
    dim = sh.root.find(Q("dimension"))
    if dim is not None:
        m = re.fullmatch(r"([A-Z]+\d+):([A-Z]+)(\d+)", dim.get("ref"))
        if m:
            dim.set("ref", f"{m.group(1)}:{m.group(2)}{max(int(m.group(3)), fin + 5)}")
    return sh.tobytes()


def _etendre_alertes(parts, fichier: str, cap: int):
    """DÉTAIL DES ALERTES : table de 45 lignes -> cap lignes ; note de règle déplacée sous la table."""
    root = etree.fromstring(parts[fichier])
    data = root.find(Q("sheetData"))
    rows = {int(r.get("r")): r for r in data.findall(Q("row"))}
    ancienne_fin, fin = 51, 6 + cap
    note = copy.deepcopy(rows[53])
    for n in [k for k in rows if k > ancienne_fin]:
        data.remove(rows[n])
    modele = rows[ancienne_fin]
    dernier = modele
    for r in range(ancienne_fin + 1, fin + 1):
        nouvelle = copy.deepcopy(modele)
        nouvelle.set("r", str(r))
        for c in nouvelle.findall(Q("c")):
            col = re.match(r"[A-Z]+", c.get("r")).group(0)
            c.set("r", f"{col}{r}")
            f = c.find(Q("f"))
            if f is not None and f.text:
                f.text = _traduire(f.text, f"{col}{ancienne_fin}", f"{col}{r}")
            elif col == "A" and c.find(Q("v")) is not None:
                c.find(Q("v")).text = str(r - 6)
        dernier.addnext(nouvelle)
        dernier = nouvelle
    debut_note = fin + 2
    note.set("r", str(debut_note))
    for c in note.findall(Q("c")):
        col = re.match(r"[A-Z]+", c.get("r")).group(0)
        c.set("r", f"{col}{debut_note}")
    dernier.addnext(note)
    # fusion de la note, mise en forme conditionnelle, dimension
    for m in root.iter(Q("mergeCell")):
        if m.get("ref") == "B53:H54":
            m.set("ref", f"B{debut_note}:H{debut_note + 1}")
    for cf in root.iter(Q("conditionalFormatting")):
        if cf.get("sqref") == "B7:H51":
            cf.set("sqref", f"B7:H{fin}")
    dim = root.find(Q("dimension"))
    dim.set("ref", f"A1:I{debut_note + 1}")
    parts[fichier] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    for nom in [n for n in parts if re.fullmatch(r"xl/tables/table\d+\.xml", n)]:
        t = parts[nom].decode("utf-8")
        if 'name="tblAlertes"' in t:
            parts[nom] = t.replace('B6:H51', f'B6:H{fin}').encode("utf-8")


def etendre_modele(src, dst, capacite: int = 500) -> dict:
    parts = lire_parts(src)
    ancienne = capacite_modele(parts)
    if capacite <= ancienne:
        raise ValueError(f"Le modèle gère déjà {ancienne} activités (≥ {capacite}).")
    vieille_fin, fin = 4 + ancienne, 4 + capacite
    parts[_fichier_feuille(parts, "CALCULS")] = _etendre_calculs(parts[_fichier_feuille(parts, "CALCULS")], ancienne, capacite)
    # noms définis
    wb = parts["xl/workbook.xml"].decode("utf-8")
    wb = re.sub(r"(CALCULS!\$[A-Z]+\$5:\$[A-Z]+\$)%d\b" % vieille_fin, lambda m: m.group(1) + str(fin), wb)
    # tableau de bord : plages en dur
    f1 = _fichier_feuille(parts, "TABLEAU DE BORD")
    root = etree.fromstring(parts[f1])
    for f in root.iter(Q("f")):
        if f.text:
            f.text = _agrandir_plages(f.text, vieille_fin, fin)
    parts[f1] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    # alertes
    _etendre_alertes(parts, _fichier_feuille(parts, "DÉTAIL DES ALERTES"), capacite)
    # matrice : fusions supprimées, feuille masquée
    f3 = _fichier_feuille(parts, "MATRICE DE SUIVI")
    root = etree.fromstring(parts[f3])
    mc = root.find(Q("mergeCells"))
    if mc is not None:
        root.remove(mc)
    parts[f3] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    wb = re.sub(r'(<sheet [^>]*name="MATRICE DE SUIVI"[^>]*?)state="visible"', r'\1state="hidden"', wb)
    parts["xl/workbook.xml"] = wb.encode("utf-8")
    # écriture (ordre d'origine conservé)
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, parts[i.filename])
    return {"ancienne_capacite": ancienne, "capacite": capacite, "fichier": str(dst)}
