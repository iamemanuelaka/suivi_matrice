"""Modification des CELLULES d'un classeur Excel directement dans son XML.

Pourquoi : openpyxl réécrit tout le classeur à l'enregistrement et peut altérer graphiques, images, carte et mise en forme du tableau de bord.
Ici seuls les fichiers de feuille modifiés (et workbook.xml pour forcer le recalcul) sont réécrits ; toutes les autres parties du classeur
(graphiques, dessins, images, styles, tableaux, thème) sont recopiées OCTET POUR OCTET.
Les cellules contenant une formule ne sont jamais modifiées.
 """
import re
import zipfile
from datetime import date, datetime
from pathlib import Path

from lxml import etree

from . import xlsx_repair as X

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
Q = lambda t: f"{{{NS}}}{t}"   # noqa: E731


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def idx_to_col(i: int) -> str:
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def serial(v) -> float:
    if isinstance(v, datetime):
        return (v - datetime(1899, 12, 30)).total_seconds() / 86400
    return float((v - date(1899, 12, 30)).days)


class Sheet:
    def __init__(self, xml: bytes):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(Q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data.findall(Q("row"))}
        self.dirty = False

    # --- accès ---------------------------------------------------------------
    def _row(self, r: int):
        if r not in self.rows:
            el = etree.Element(Q("row"), r=str(r))
            later = [k for k in self.rows if k > r]
            (self.rows[min(later)].addprevious(el) if later else self.data.append(el))
            self.rows[r] = el
        return self.rows[r]

    def _cell(self, r: int, c: int, create=False):
        row, ref = self._row(r) if create else self.rows.get(r), f"{idx_to_col(c)}{r}"
        if row is None:
            return None
        for cell in row.findall(Q("c")):
            if cell.get("r") == ref:
                return cell
        if not create:
            return None
        el = etree.Element(Q("c"), r=ref)
        later = [x for x in row.findall(Q("c")) if col_to_idx(re.match(r"[A-Z]+", x.get("r"))[0]) > c]
        (later[0].addprevious(el) if later else row.append(el))
        return el

    def get(self, r: int, c: int):
        cell = self._cell(r, c)
        if cell is None:
            return None
        t, v, is_ = cell.get("t"), cell.find(Q("v")), cell.find(Q("is"))
        if t == "inlineStr" and is_ is not None:
            return "".join(is_.itertext())
        if v is None or v.text in (None, ""):
            return None
        if t in (None, "n"):
            f = float(v.text)
            return int(f) if f.is_integer() else f
        return v.text

    def has_formula(self, r: int, c: int) -> bool:
        cell = self._cell(r, c)
        return cell is not None and cell.find(Q("f")) is not None

    def style(self, r: int, c: int):
        cell = self._cell(r, c)
        return cell.get("s") if cell is not None else None

    # --- écriture --------------------------------------------------------------
    def set(self, r: int, c: int, value, style_from: tuple | None = None) -> bool:
        """Écrit une valeur (None = vider). Retourne False si la cellule contient une formule (non modifiée)."""
        cell = self._cell(r, c, create=True)
        if cell.find(Q("f")) is not None:
            return False
        if style_from and (s := self.style(*style_from)) is not None:
            cell.set("s", s)
        for child in list(cell):
            cell.remove(child)
        if value is None or value == "":
            cell.set("t", "n")
        elif isinstance(value, (bool,)):
            raise TypeError("booléen non géré")
        elif isinstance(value, (int, float)):
            cell.set("t", "n")
            etree.SubElement(cell, Q("v")).text = repr(value)
        elif isinstance(value, (date, datetime)):
            cell.set("t", "n")
            etree.SubElement(cell, Q("v")).text = repr(serial(value))
        else:
            cell.set("t", "inlineStr")
            t = etree.SubElement(etree.SubElement(cell, Q("is")), Q("t"))
            t.text = str(value)
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        self.dirty = True
        return True

    def set_row_height(self, r: int, height: float) -> None:
        row = self._row(r)
        row.set("ht", str(height))
        row.set("customHeight", "1")
        self.dirty = True

    def merged(self) -> list[str]:
        mc = self.root.find(Q("mergeCells"))
        return [m.get("ref") for m in mc] if mc is not None else []

    def tobytes(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)


class Package:
    def __init__(self, path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path) as z:
            self.infos = z.infolist()
            self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        wb = etree.fromstring(self.parts["xl/workbook.xml"])
        rels = etree.fromstring(self.parts["xl/_rels/workbook.xml.rels"])
        target = {r.get("Id"): r.get("Target") for r in rels}
        self._sheet_files = {}
        for s in wb.find(Q("sheets")):
            t = target[s.get(f"{{{RNS}}}id")].lstrip("/")
            self._sheet_files[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
        self._sheets: dict[str, Sheet] = {}
        self.corrections = list(X.reparer_parts(self.parts, self._sheet_files))

    def has_sheet(self, name: str) -> bool:
        return name in self._sheet_files

    def sheet(self, name: str) -> Sheet:
        if name not in self._sheets:
            self._sheets[name] = Sheet(self.parts[self._sheet_files[name]])
        return self._sheets[name]

    def force_full_calc(self) -> None:
        wb = self.parts["xl/workbook.xml"].decode("utf-8")
        if "<calcPr" in wb:
            wb = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb)
        else:
            wb = wb.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        self.parts["xl/workbook.xml"] = wb.encode("utf-8")

    def save(self, out) -> None:
        for name, sh in self._sheets.items():
            if sh.dirty:
                self.parts[self._sheet_files[name]] = sh.tobytes()
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        tmp = out.with_suffix(".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for i in self.infos:
                z.writestr(i.filename, self.parts[i.filename])
        tmp.replace(out)
