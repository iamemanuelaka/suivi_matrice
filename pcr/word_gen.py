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
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor

from .calculs import indicateurs_delai
from .config import normalize
from .kobo_photos import CHAMPS_PHOTOS
from .referentiel import Referentiel

FONT, SIZE = "Arial", 9
WIDTHS_CM = (6.4, 6.4, 6.9, 7.0)          # Désignation, Avancement, Commentaires, Photo (total 26,7 cm)
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
        "photos": [s[c] for c in CHAMPS_PHOTOS if s.get(c)] if s else [],   # jusqu'à 4 photos du site, dans l'ordre du formulaire
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

def _add_photos(cell, noms: list) -> None:
    """Colonne « Photo » : une photo occupe toute la cellule ; 2 à 4 photos sont rangées en 2 × 2 (cellule divisée en 4).
    Vide si aucune photo. Les originaux restent dans data/photos/ si quelqu'un veut les voir en grand."""
    noms = [n for n in noms if n][:4]
    if len(noms) <= 1:
        _add_photo(cell, noms[0] if noms else None, boite=(WIDTHS_CM[3] - 0.5, 4.8))   # une photo : toute la largeur, 4,8 cm de haut au plus
        return
    case_l = (WIDTHS_CM[3] - 0.4) / 2                     # deux cases par ligne dans la colonne « Photo »
    boite = (case_l - 0.2, 2.6)                            # chaque photo tient dans 3,1 × 2,6 cm, proportions gardées
    grille = cell.add_table(rows=(len(noms) + 1) // 2, cols=2)
    _sans_bordure(grille, largeur_cm=case_l * 2)
    for col in grille.columns:
        col.width = Cm(case_l)
    for i, nom in enumerate(noms):
        sous = grille.cell(i // 2, i % 2)
        sous.width = Cm(case_l)
        sous.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _add_photo(sous, nom, boite=boite)
    if not cell.paragraphs[0].text and len(cell.paragraphs) > 1:                       # pas de ligne vide au-dessus de la grille
        p0 = cell.paragraphs[0]._element
        p0.getparent().remove(p0)


def _sans_bordure(table, largeur_cm: float) -> None:
    """Grille de photos : sans bordure, largeur fixe (ne déborde jamais de la cellule)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    table.autofit = False
    tblPr = table._tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(int(largeur_cm * 567)))
    tw.set(qn("w:type"), "dxa")
    tblPr.append(tw)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    borders = OxmlElement("w:tblBorders")
    for cote in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{cote}")
        b.set(qn("w:val"), "nil")
        borders.append(b)
    tblPr.append(borders)
    marges = OxmlElement("w:tblCellMar")
    for cote in ("left", "right"):
        m = OxmlElement(f"w:{cote}")
        m.set(qn("w:w"), "28")
        m.set(qn("w:type"), "dxa")
        marges.append(m)
    tblPr.append(marges)


def _add_photo(cell, photo_filename: str | None, boite: tuple | None = None) -> None:
    """Insère une photo du site dans la cellule (vide si aucune photo).
    Gère les formats courants (JPEG, PNG, GIF, BMP, TIFF, WebP…).
    Les photos manquantes ou illisibles ne bloquent jamais la génération.
    """
    if not photo_filename:
        return

    # Kobo peut renvoyer un chemin ; on ne garde que le nom de fichier.
    # Les photos téléchargées sont enregistrées en « <nom sans extension>.jpg ».
    nom = Path(photo_filename).name
    photo_path = next((c for c in (Path("data/photos") / (Path(nom).stem + ".jpg"), Path("data/photos") / nom) if c.is_file()),
                      Path("data/photos") / nom)

    if not photo_path.is_file():
        p = _para(cell)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, f"📷 Photo introuvable : {photo_path.name}")
        return

    # Formats nativement supportés par python-docx
    native = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif"}
    # Formats à convertir via Pillow (WebP, HEIC, AVIF…)
    to_convert = {".webp", ".heic", ".heif", ".avif", ".jfif"}

    try:
        suffix = photo_path.suffix.lower()
        path_to_insert = photo_path

        if suffix in to_convert or suffix not in native:
            try:
                from PIL import Image
                import tempfile
                img = Image.open(photo_path)
                # Convertir en RGB si nécessaire (transparence, CMYK…)
                if img.mode in ("RGBA", "P", "LA"):
                    background = Image.new("RGB", img.size, (255, 255, 255))
                    if img.mode == "P":
                        img = img.convert("RGBA")
                    background.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
                    img = background
                elif img.mode != "RGB":
                    img = img.convert("RGB")
                # Écrire un JPEG temporaire
                tmp = Path(tempfile.gettempdir()) / f"pcr_photo_{photo_path.stem}.jpg"
                img.save(tmp, "JPEG", quality=85, optimize=True)
                path_to_insert = tmp
            except Exception:
                # Si Pillow échoue, on tente quand même l'insertion directe
                path_to_insert = photo_path

        p = _para(cell)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        if boite:                                           # un quart de cellule : la photo tient dans la case
            from PIL import Image as _Im
            with _Im.open(path_to_insert) as im:
                w, h = im.size
            l, hmax = boite
            if h and w / h < l / hmax:
                run.add_picture(str(path_to_insert), height=Cm(hmax))
            else:
                run.add_picture(str(path_to_insert), width=Cm(l))
        else:
            run.add_picture(str(path_to_insert), width=Cm(WIDTHS_CM[3] - 0.5))

        # Nettoyage du temporaire si créé
        if path_to_insert != photo_path and path_to_insert.exists():
            path_to_insert.unlink(missing_ok=True)

    except Exception as e:
        p = _para(cell)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _run(p, f"📷 Impossible d'insérer la photo ({e})")

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

    table = doc.add_table(rows=1, cols=4)
    table.style, table.alignment, table.autofit = "Table Grid", WD_TABLE_ALIGNMENT.CENTER, False
    
    # CORRECTION : Calcul de la date réelle des données pour éviter la date trompeuse
    last_dates = [s["Date_Actualisation"] for s in last.values() if s.get("Date_Actualisation")]
    date_donnees_reelles = max(last_dates) if last_dates else au
    delta_jours = (au - date_donnees_reelles).days if date_donnees_reelles else 0
    suffixe_alerte = f" (Données arrêtées au {date_donnees_reelles:%d/%m/%Y})" if delta_jours > 1 else ""
    titre_colonne_2 = f"Avancement physique au {au:%d/%m/%Y}{suffixe_alerte}"
    
    for c, (title, w) in enumerate(zip(("Désignation Activité", titre_colonne_2, "Commentaires", "Photo"), WIDTHS_CM)):
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
        c1, c2, c3, c4 = first.cells
        if n_rows == 3:
            c1, c3 = rows[0].cells[0].merge(rows[2].cells[0]), rows[0].cells[2].merge(rows[2].cells[2])
            c4 = rows[0].cells[3].merge(rows[2].cells[3])
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
        c4.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _add_photos(c4, v.get("photos") or [])
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
