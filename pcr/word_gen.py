"""Phase 5 : génération des matrices Word (une par catégorie) à partir de la base.

Structure reproduite (d'après les matrices fournies) : titres, tableau à 3 colonnes « Désignation Activité | Avancement physique au JJ/MM/AAAA |
Commentaires ». Colonne « Avancement » : les observations techniques en tirets, AU-DESSUS du taux d'avancement, du délai consommé et du total
facturé. Marchés ruraux : 3 lignes par activité (observations / taux / délai) avec colonnes 1 et 3 fusionnées.

IMPORTANT : les fichiers .docx fournis étaient des extraits texte. La mise en page (polices, largeurs, orientation) est donc une reconstruction ;
dès que les vrais .docx sont disponibles, ils peuvent servir de modèle sans changer le contenu généré.
"""
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from .calculs import indicateurs_delai
from .config import normalize
from .referentiel import Referentiel

FONT, SIZE = "Arial", 9
WIDTHS_CM = (10.2, 9.8, 6.7)
NUM_ID = 90


def money(n) -> str:
    return f"{n:,}".replace(",", " ")


def pct(v, decimals=None) -> str:
    if v is None:
        return ""
    return f"{v:.{decimals}f}%" if decimals is not None else f"{v:g}%"


def fdate(d) -> str:
    return f"{d:%d/%m/%Y}" if d else ""


# --------------------------------------------------------------------------
# Choix des données
# --------------------------------------------------------------------------
def latest_valid(ref: Referentiel) -> dict:
    best = {}
    for r in ref.base.rows("Suivi_Activites"):
        if r["Statut_Traitement"] != "VALIDE":
            continue
        key = (r["Date_Actualisation"] or date.min, str(r["Date_Importation"] or ""), str(r["ID_SOUMISSION_KOBO"]))
        if r["CODE_ACTIVITE"] not in best or key > best[r["CODE_ACTIVITE"]][0]:
            best[r["CODE_ACTIVITE"]] = (key, r)
    return {c: v[1] for c, v in best.items()}


def ordered_activities(ref: Referentiel, cat: str) -> list[dict]:
    """Activités actives de la catégorie : celles dont le titulaire est connu d'abord, puis les marchés pas encore attribués ;
    dans chaque groupe, régions par ordre alphabétique puis numéro d'ordre."""
    acts = [a for a in ref.activities() if a["Categorie"] == cat and a["Statut"] == "Actif"]
    return sorted(acts, key=lambda a: (0 if a["Entreprise"] else 1, normalize(ref.cfg.regions[a["Region"]]), a["CODE_ACTIVITE"]))


def build_view(ref: Referentiel, a: dict, s: dict | None, au: date) -> dict:
    """Valeurs affichées pour une activité (aucune règle de mise en page ici)."""
    ind = indicateurs_delai(a["Date_Debut"], a["Delai_Mois"], au, ref.cfg.conventions[a["Categorie"]])
    baseline = bool(s) and s["Source"] == "BASELINE_WORD"
    return {
        "code": a["CODE_ACTIVITE"], "region": normalize(ref.cfg.regions[a["Region"]]), "intitule": a["Designation"] or "",
        "montant": money(a["Montant_Marche"]) if a["Montant_Marche"] else None, "titulaire": a["Entreprise"] or "", "os": fdate(a["Date_OS"]),
        "demarrage": fdate(a["Date_Debut"]), "delai": f"{a['Delai_Mois']} mois" if a["Delai_Mois"] else "", "fin": fdate(a["Date_Fin_Previsionnelle"]),
        "bureau": a["Bureau_Controle"] or "",
        "obs": [l.strip() for l in (s["Observations_Techniques"] or "").splitlines() if l.strip()] if s else [],
        "taux": pct(s["Taux_Execution_Physique"]) if s else "", "delai_conso": pct(ind["consommation_pct"], 2) if ind else "",
        "facture": pct(s["Total_Facture"]) if s else "",
        "commentaires": (s["Commentaires"] or "") if s else "",
        # Kobo : « RAS » si le technicien n'a rien signalé ; suivi de référence : on reprend le Word tel quel
        "difficultes": ((s["Difficultes"] or ("" if baseline else "RAS")) if s else None),
        "recommandations": (s["Recommandations"] or "") if s else None,
    }


# --------------------------------------------------------------------------
# Mise en forme Word
# --------------------------------------------------------------------------
def _run(p, text, bold=False, size=SIZE, color=None):
    r = p.add_run(text)
    r.bold, r.font.size, r.font.name = bold, Pt(size), FONT
    r._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _para(cell):
    """Paragraphe suivant dans une cellule (réutilise le paragraphe vide initial)."""
    p = cell.paragraphs[0] if (len(cell.paragraphs) == 1 and not cell.paragraphs[0].text) else cell.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(0)
    return p


def _labelled(cell, label, value, bold_value=False):
    p = _para(cell)
    _run(p, label + " ", bold=True)
    if value:
        _run(p, value, bold=bold_value)
    return p


def _dash(cell, text):
    p = _para(cell)
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{NUM_ID}"/></w:numPr>'))
    _run(p, text)
    return p


def _setup_numbering(doc):
    numbering = doc.part.numbering_part.element
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{NUM_ID}"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/>'
        f'<w:numFmt w:val="bullet"/><w:lvlText w:val="–"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="284" w:hanging="284"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{FONT}" w:hAnsi="{FONT}" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>')
    first_num = numbering.find(qn("w:num"))
    (first_num.addprevious(abstract) if first_num is not None else numbering.append(abstract))
    numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{NUM_ID}"><w:abstractNumId w:val="{NUM_ID}"/></w:num>'))


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _page_footer(section, text):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _run(p, text + "  –  Page ", size=8)
    for kind, instr in (("begin", None), (None, " PAGE "), ("end", None)):
        r = p.add_run()
        r.font.size, r.font.name = Pt(8), FONT
        if kind:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), kind)
            r._r.append(f)
        else:
            t = OxmlElement("w:instrText")
            t.set(qn("xml:space"), "preserve")
            t.text = instr
            r._r.append(t)


def _cell1(cell, v, cat):
    _para(cell)
    _run(cell.paragraphs[0], v["region"], bold=True)
    _labelled(cell, "Intitulé :", v["intitule"])
    if v["montant"]:
        _labelled(cell, "Montant marché :", v["montant"])
    _labelled(cell, "Titulaire :", v["titulaire"], bold_value=True)
    _labelled(cell, "O.S. :", v["os"])
    if cat == "MR":
        _labelled(cell, "Date de démarrage :", v["demarrage"])
    _labelled(cell, "Délai :", v["delai"])
    _labelled(cell, "Fin prévue :", v["fin"])
    _labelled(cell, "Bureau de contrôle :", v["bureau"])


def _cell3(cell, v):
    if v["commentaires"]:
        _run(_para(cell), v["commentaires"])
    if v["difficultes"]:
        _labelled(cell, "Difficultés :", v["difficultes"])
    if v["recommandations"]:
        _labelled(cell, "Recommandations :", v["recommandations"])


def render(ref: Referentiel, cat: str, au: date, path) -> Path:
    cfg, last = ref.cfg, latest_valid(ref)
    conf = cfg.category(cat)
    doc = Document()
    _setup_numbering(doc)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(SIZE)
    sec = doc.sections[0]
    sec.orientation, sec.page_width, sec.page_height = WD_ORIENT.LANDSCAPE, Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    _page_footer(sec, conf["pied_de_page"])
    for text, size in ((conf["titre_1"], 12), (conf["titre_2"], 11)):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        _run(p, text, bold=True, size=size)

    table = doc.add_table(rows=1, cols=3)
    table.style, table.alignment, table.autofit = "Table Grid", WD_TABLE_ALIGNMENT.CENTER, False
    for c, (title, w) in enumerate(zip(("Désignation Activité", f"Avancement physique au {au:%d/%m/%Y}", "Commentaires"), WIDTHS_CM)):
        cell = table.rows[0].cells[c]
        cell.width = Cm(w)
        _shade(cell, "D9E2F3")
        _run(_para(cell), title, bold=True)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

    for a in ordered_activities(ref, cat):
        v = build_view(ref, a, last.get(a["CODE_ACTIVITE"]), au)
        n_rows = 3 if cat == "MR" else 1
        rows = [table.add_row() for _ in range(n_rows)]
        for r in rows:
            for c, w in enumerate(WIDTHS_CM):
                r.cells[c].width = Cm(w)
            if cat == "MR":
                r._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        first = rows[0]
        c1, c2, c3 = first.cells
        if n_rows == 3:
            c1, c3 = rows[0].cells[0].merge(rows[2].cells[0]), rows[0].cells[2].merge(rows[2].cells[2])
        _cell1(c1, v, cat)
        mid = rows[0].cells[1]
        for line in v["obs"]:
            _dash(mid, line)
        if cat == "BS":
            _labelled(mid, "Taux d’avancement :", v["taux"], bold_value=True)
            _labelled(mid, "Délai consommé :", v["delai_conso"], bold_value=True)
            if v["facture"]:
                _labelled(mid, "Total facturé :", v["facture"], bold_value=True)
        else:
            if v["facture"]:
                _labelled(mid, "Total facturé :", v["facture"], bold_value=True)
            _labelled(rows[1].cells[1], "Taux d’avancement :", v["taux"], bold_value=True)
            _labelled(rows[2].cells[1], "Délai consommé :", v["delai_conso"], bold_value=True)
        _cell3(c3, v)
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return path


def generate_all(ref: Referentiel, au: date, out_dir) -> list[Path]:
    out = []
    for cat in ref.cfg.categories:
        conf = ref.cfg.category(cat)
        if conf.get("fichier_word"):
            out.append(render(ref, cat, au, Path(out_dir) / f"{conf['fichier_word']}_{au:%d%m%Y}.docx"))
    return out
