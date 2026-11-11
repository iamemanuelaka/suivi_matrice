"""Détection et réparation des défauts du classeur qui font afficher à Excel « Nous avons trouvé un problème… ».

Défauts traités (tous constatés dans le modèle `MATRICE_DE_SUIVI_DES_ACTIVITES_modele.xlsx`) :

1. DESSINS — « Dessin dans /xl/drawings/drawingN.xml (Forme de dessin) »
   Dans chaque image, la géométrie `<a:prstGeom prst="rect">` contenait `<avLst/>` SANS préfixe `a:`. Dans le document de dessin
   (espace de noms par défaut = spreadsheetDrawing) l'élément se retrouvait dans le MAUVAIS espace de noms
   (`xdr:avLst` au lieu de `a:avLst`) : XML bien formé mais invalide au regard de la norme, et Excel répare en supprimant la forme.
2. FORMULES — « Formule dans /xl/worksheets/sheetN.xml »
   Excel refuse une constante TEXTE de plus de 255 caractères dans une formule. La formule de la règle d'alerte
   (feuille « DÉTAIL DES ALERTES », B53) contenait un texte de 261 caractères : Excel supprime alors la formule.
   Correction : le texte est découpé en morceaux ≤ 200 caractères concaténés avec & (résultat identique).

Les contrôles ci-dessous sont aussi appliqués au fichier GÉNÉRÉ : un classeur qui présenterait l'un de ces défauts n'est jamais livré.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/package/2006/relationships"

MAX_LITTERAL = 255          # limite Excel : constante texte dans une formule
MAX_FORMULE = 8192          # limite Excel : longueur d'une formule
MAX_IMBRICATION = 64        # limite Excel : niveaux de parenthèses
MORCEAU = 200

_LIT = re.compile(r'"((?:[^"]|"")*)"')
_UNITES = re.compile(r'""|.', re.S)


def lire_parts_brutes(chemin) -> dict[str, bytes]:
    """Lit le XLSX sans modifier son contenu. Utilisé pour les diagnostics de défauts."""
    with zipfile.ZipFile(chemin) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


# ---------------------------------------------------------------- formules
def decouper_litteraux(formule: str) -> str:
    """Découpe toute constante texte > 255 caractères en `"a"&"b"&…` (sémantique inchangée)."""
    def rep(m):
        unites = _UNITES.findall(m.group(1))
        if len(unites) <= MAX_LITTERAL:
            return m.group(0)
        morceaux = ["".join(unites[i:i + MORCEAU]) for i in range(0, len(unites), MORCEAU)]
        return "&".join(f'"{p}"' for p in morceaux)
    return _LIT.sub(rep, formule)


def _imbrication(f: str) -> int:
    f = _LIT.sub('""', f)
    d = m = 0
    for ch in f:
        if ch == "(":
            d += 1
            m = max(m, d)
        elif ch == ")":
            d -= 1
    return m


def _serialiser(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def reparer_dessin(xml: bytes) -> tuple[bytes, int]:
    """Replace `avLst` dans l'espace de noms `a:` ; retourne (xml, nombre d'éléments corrigés)."""
    root = etree.fromstring(xml)
    n = 0
    for el in root.iter(f"{{{XDR}}}avLst"):
        el.tag = f"{{{A}}}avLst"
        n += 1
    return (_serialiser(root) if n else xml), n


def reparer_feuille(xml: bytes) -> tuple[bytes, list[str]]:
    root = etree.fromstring(xml)
    faits = []
    for c in root.iter(f"{{{S}}}c"):
        f = c.find(f"{{{S}}}f")
        if f is not None and f.text:
            nouveau = decouper_litteraux(f.text)
            if nouveau != f.text:
                f.text = nouveau
                faits.append(c.get("r"))
    return (_serialiser(root) if faits else xml), faits


def reparer_parts(parts: dict[str, bytes], noms_feuilles: dict[str, str] | None = None) -> list[str]:
    """Répare `parts` (nom → octets) sur place ; retourne la liste lisible des corrections."""
    notes = []
    for nom in sorted(parts):
        if re.fullmatch(r"xl/drawings/drawing\d+\.xml", nom):
            parts[nom], n = reparer_dessin(parts[nom])
            if n:
                notes.append(f"{nom} : {n} élément(s) avLst remis dans l'espace de noms a:")
        elif re.fullmatch(r"xl/worksheets/sheet\d+\.xml", nom):
            parts[nom], cells = reparer_feuille(parts[nom])
            if cells:
                libelle = (noms_feuilles or {}).get(nom, nom)
                notes.append(f"{libelle} : constante texte > {MAX_LITTERAL} caractères découpée en {', '.join(cells)}")
    return notes


# ---------------------------------------------------------------- diagnostic (sans Excel)
def diagnostiquer(parts: dict[str, bytes]) -> list[str]:
    """Contrôles statiques des causes connues de réparation par Excel. Liste vide = rien à signaler."""
    pb: list[str] = []
    for nom in sorted(parts):
        if re.fullmatch(r"xl/drawings/drawing\d+\.xml", nom):
            root = etree.fromstring(parts[nom])
            if next(root.iter(f"{{{XDR}}}avLst"), None) is not None:
                pb.append(f"{nom} : élément avLst dans le mauvais espace de noms (Excel répare « Forme de dessin »)")
            ids = [e.get("id") for e in root.iter(f"{{{XDR}}}cNvPr")]
            if len(ids) != len(set(ids)):
                pb.append(f"{nom} : identifiants de formes en double {ids}")
            for tag in ("ext",):
                for e in root.iter(f"{{{XDR}}}{tag}"):
                    if e.get("cx") in (None, "0") or e.get("cy") in (None, "0"):
                        pb.append(f"{nom} : forme de taille nulle")
        elif re.fullmatch(r"xl/worksheets/sheet\d+\.xml", nom):
            root = etree.fromstring(parts[nom])
            for c in root.iter(f"{{{S}}}c"):
                f = c.find(f"{{{S}}}f")
                if f is None or not f.text:
                    continue
                longs = [len(_UNITES.findall(m.group(1))) for m in _LIT.finditer(f.text)]
                if longs and max(longs) > MAX_LITTERAL:
                    pb.append(f"{nom} {c.get('r')} : constante texte de {max(longs)} caractères (> {MAX_LITTERAL}) : Excel supprime la formule")
                if len(f.text) > MAX_FORMULE:
                    pb.append(f"{nom} {c.get('r')} : formule de {len(f.text)} caractères (> {MAX_FORMULE})")
                if _imbrication(f.text) > MAX_IMBRICATION:
                    pb.append(f"{nom} {c.get('r')} : plus de {MAX_IMBRICATION} niveaux d'imbrication")
        elif re.fullmatch(r"xl/tables/table\d+\.xml", nom):
            pb += _controle_table(nom, parts)
    pb += _controle_relations(parts)
    return pb


def _controle_table(nom: str, parts: dict[str, bytes]) -> list[str]:
    """Les cellules d'en-tête d'une table doivent être du texte identique aux noms de colonnes."""
    t = etree.fromstring(parts[nom])
    ref = t.get("ref")
    m = re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref or "")
    if not m:
        return [f"{nom} : plage de table illisible {ref!r}"]
    cible = f"xl/tables/{Path(nom).name}"
    feuille = None
    for rel_nom, data in parts.items():
        if rel_nom.startswith("xl/worksheets/_rels/") and cible.encode() in data.replace(b"/xl/", b"xl/"):
            feuille = "xl/worksheets/" + Path(rel_nom).name.removesuffix(".rels")
    if feuille is None or feuille not in parts:
        return []
    sh = etree.fromstring(parts[feuille])
    colnames = [c.get("name") for c in t.iter(f"{{{S}}}tableColumn")]
    deb = _col(m.group(1))
    entete = {}
    for c in sh.iter(f"{{{S}}}c"):
        mm = re.fullmatch(r"([A-Z]+)(\d+)", c.get("r"))
        if int(mm.group(2)) == int(m.group(2)):
            if c.find(f"{{{S}}}f") is not None:
                entete[_col(mm.group(1))] = None
            else:
                entete[_col(mm.group(1))] = "".join(c.itertext()).strip()
    pb = []
    for i, nm in enumerate(colnames):
        v = entete.get(deb + i, "")
        if v is None:
            pb.append(f"{nom} : l'en-tête {nm!r} contient une formule")
        elif v != nm:
            pb.append(f"{nom} : en-tête de cellule {v!r} ≠ nom de colonne {nm!r}")
    return pb


def _col(lettres: str) -> int:
    n = 0
    for ch in lettres:
        n = n * 26 + ord(ch) - 64
    return n


def _controle_relations(parts: dict[str, bytes]) -> list[str]:
    pb = []
    for nom, data in parts.items():
        if not nom.endswith(".rels"):
            continue
        base = Path(nom).parent.parent
        for r in etree.fromstring(data):
            if r.get("TargetMode") == "External":
                continue
            t = r.get("Target")
            chemin = t.lstrip("/") if t.startswith("/") else (base / t).as_posix()
            while "/../" in chemin:
                chemin = re.sub(r"[^/]+/\.\./", "", chemin, count=1)
            if chemin not in parts:
                pb.append(f"{nom} : cible introuvable {t}")
    return pb


def lire_parts(chemin) -> dict[str, bytes]:
    """Lit le classeur et corrige automatiquement les défauts connus avant diagnostic/usage."""
    parts = lire_parts_brutes(chemin)
    reparer_parts(parts)
    return parts
