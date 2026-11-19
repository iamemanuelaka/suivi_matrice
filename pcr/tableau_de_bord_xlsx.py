"""Tableau de bord Excel du PCR-CI, généré de A à Z depuis la base (aucun modèle à maintenir).

Trois feuilles :
  - TABLEAU DE BORD : 4 filtres (listes déroulantes), 6 cartes, graphique « consommation des délais », aperçu des alertes ;
  - ALERTES         : toutes les activités en alerte (tableau Excel filtrable) ;
  - DONNÉES (masquée) : une ligne par activité. Tout ce qui ne dépend pas des filtres (délai consommé, alerte, ONG…) y est
    calculé par Python avec les règles de pcr/indicateurs.py (les mêmes que l'application) ; les formules ne font que filtrer,
    additionner et classer.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .indicateurs import STATUTS, alias_ong, lignes, ong_cle
from .referentiel import Referentiel

# --- charte (reprise de l'ancien tableau de bord) --------------------------------------------------------------
MARINE, TEXTE, VERT, BLEU, ROUGE, ORANGE, GRIS = "0D3B6E", "1B2B41", "3A9A3F", "1F5FA8", "D62828", "F2A33A", "6B7A90"
BANDEAU, FOND_FILTRE, TRAIT = "BDD7EE", "F7F9FC", "D5DDE8"
LOGO = Path(__file__).resolve().parent.parent / "modeles" / "logo_pcr_ci.png"
TB, AL, DO = "TABLEAU DE BORD", "ALERTES", "DONNÉES"
TOUTES = {"nature": "Toutes les natures", "region": "Toutes les régions", "entreprise": "Toutes les entreprises",
          "statut": "Tous les statuts"}
MAX_LIGNES = 2000                  # capacité des formules (une ligne par activité)
MAX_GROUPES = 24                   # barres du graphique (au plus ; ajusté aux données : régions ou entreprises d'une région)

# DONNÉES : colonnes écrites par Python (valeurs) puis colonnes de formules
COLS = ["N°", "Code", "Nature", "Région", "Département", "Entreprise", "Statut", "PGES", "ONG (clé)", "ONG sans nom",
        "Montant", "Décaissé", "Taux", "Début", "Délai consommé", "Évaluable", "Alerte", "Écart", "Écoulé (éval.)",
        "Durée (éval.)", "Montant (éval.)", "Montant×taux (éval.)", "Taux (éval.)",
        # formules
        "Dans le filtre", "ONG 1re occ.", "Rang alerte", "Groupe", "Groupe 1re occ.", "Groupe : nb évaluées",
        "Groupe : délai consommé", "Groupe : exécution", "Groupe : rang", "Groupe : libellé"]
C = {nom: get_column_letter(i) for i, nom in enumerate(COLS, 1)}
LISTES_COL = get_column_letter(len(COLS) + 2)          # listes des filtres, à droite des données
GRAPH_COL = get_column_letter(len(COLS) + 7)           # données du graphique, triées


def _rng(col: str) -> str:
    return f"'{DO}'!${C[col]}$2:${C[col]}${MAX_LIGNES + 1}"


def _fill(c):
    return PatternFill("solid", fgColor=c)


def _bord(cell, gauche=None, tout=None):
    fin = Side(style="thin", color=TRAIT)
    cell.border = Border(left=Side(style="thick", color=gauche) if gauche else fin, right=fin if tout else None,
                         top=fin if tout else None, bottom=fin if tout else None)


def generer(ref: Referentiel, chemin, au: date | None = None, maintenant: datetime | None = None) -> dict:
    """Écrit le tableau de bord ; retourne {fichier, activites, alertes}."""
    au = au or date.today()
    maintenant = (maintenant or datetime.now()).replace(microsecond=0)
    cfg = ref.cfg
    seuil_taux = cfg.data.get("controles", {}).get("seuil_alerte_taux", 30)
    seuil_conso = cfg.data.get("controles", {}).get("seuil_alerte_delai", 50)
    lg = lignes(ref, au)
    if len(lg) > MAX_LIGNES:
        raise ValueError(f"{len(lg)} activités : le tableau de bord est prévu pour {MAX_LIGNES} au maximum.")
    alias = alias_ong(cfg)
    par_region: dict = {}
    for l in lg:
        if l.entreprise:
            par_region.setdefault(l.region, set()).add(l.entreprise)
    n_groupes = min(MAX_GROUPES, max([len(cfg.regions)] + [len(v) for v in par_region.values()]))

    wb = Workbook()
    wb._n_groupes = n_groupes
    tb = wb.active
    tb.title = TB
    al = wb.create_sheet(AL)
    do = wb.create_sheet(DO)

    _donnees(wb, do, lg, alias, cfg)
    _tableau_de_bord(wb, tb, len(lg), maintenant, seuil_taux, seuil_conso)
    alertes = sorted((l for l in lg if l.alerte), key=lambda l: (-(l.conso - l.taux), l.lot))
    _alertes(al, alertes, maintenant, seuil_taux, seuil_conso)

    do.sheet_state = "hidden"
    wb.active = 0
    chemin = Path(chemin)
    chemin.parent.mkdir(parents=True, exist_ok=True)
    tmp = chemin.with_name(chemin.name + ".tmp")
    wb.save(tmp)
    _format_etiquettes(tmp)
    tmp.replace(chemin)
    return {"fichier": str(chemin), "activites": len(lg), "alertes": len(alertes)}


def _format_etiquettes(chemin: Path) -> None:
    """openpyxl écrit le format des étiquettes sans « sourceLinked » : Excel et LibreOffice l'ignorent alors (décimales, -1 visibles)."""
    import shutil
    import zipfile
    tmp = chemin.with_name(chemin.name + ".z")
    with zipfile.ZipFile(chemin) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart"):
                txt = data.decode("utf-8")
                debut = txt.find("<dLbls>")
                fin = txt.find("</dLbls>", debut)
                if debut >= 0:
                    bloc = txt[debut:fin].replace('"/><dLblPos', '" sourceLinked="0"/><dLblPos', 1)
                    txt = txt[:debut] + bloc + txt[fin:]
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, chemin)


# --- feuille DONNÉES -------------------------------------------------------------------------------------------
def _donnees(wb, ws, lg, alias, cfg):
    gras = Font(bold=True, color="FFFFFF")
    for i, nom in enumerate(COLS, 1):
        c = ws.cell(1, i, nom)
        c.font, c.fill = gras, _fill(MARINE)
        ws.column_dimensions[get_column_letter(i)].width = 14
    for n, l in enumerate(lg, 2):
        cle, anonyme = ong_cle(l.ong, alias)
        ev = l.evaluable
        valeurs = [l.lot, l.code, l.nature, l.region, l.departement.title(), l.entreprise or None, l.statut, l.pges, cle or None,
                   int(anonyme), l.montant, l.decaisse, l.taux, l.debut, round(l.conso, 6) if l.conso is not None else None,
                   int(ev), int(l.alerte), round(l.conso - l.taux, 6) if ev else None,
                   l.ecoule if ev else 0, l.duree if ev else 0, l.montant if ev else 0, l.montant * l.taux if ev else 0,
                   l.taux if ev else 0]
        for i, v in enumerate(valeurs, 1):
            ws.cell(n, i, v)
        ws.cell(n, COLS.index("Début") + 1).number_format = "DD/MM/YYYY"
        r = n
        f = {k: f"{C[k]}{r}" for k in C}
        filtre = " * ".join([
            f'OR(selNature="",selNature="{TOUTES["nature"]}",{f["Nature"]}=selNature)',
            f'OR(selRegion="",selRegion="{TOUTES["region"]}",{f["Région"]}=selRegion)',
            f'OR(selEntreprise="",selEntreprise="{TOUTES["entreprise"]}",{f["Entreprise"]}=selEntreprise)',
            f'OR(selStatut="",selStatut="{TOUTES["statut"]}",{f["Statut"]}=selStatut)'])
        col = lambda k: C[k]                                                  # noqa: E731
        ws[f["Dans le filtre"]] = f"=({filtre})*1"
        ws[f["ONG 1re occ."]] = (f'=IF(AND({f["Dans le filtre"]}=1,{f["ONG (clé)"]}<>""),'
                                 f'IF(COUNTIFS(${col("ONG (clé)")}$2:{f["ONG (clé)"]},{f["ONG (clé)"]},'
                                 f'${col("Dans le filtre")}$2:{f["Dans le filtre"]},1)=1,1,0),0)')
        ws[f["Rang alerte"]] = (f'=IF(AND({f["Dans le filtre"]}=1,{f["Alerte"]}=1),'
                                f'COUNTIFS(dFiltre,1,dAlerte,1,dEcart,">"&{f["Écart"]})'
                                f'+COUNTIFS(${col("Dans le filtre")}$2:{f["Dans le filtre"]},1,${col("Alerte")}$2:{f["Alerte"]},1,'
                                f'${col("Écart")}$2:{f["Écart"]},{f["Écart"]}),"")')
        # graphique : par région, ou par entreprise quand une région est choisie
        ws[f["Groupe"]] = f'=IF({f["Dans le filtre"]}=1,IF(parEntreprise,{f["Entreprise"]}&"",{f["Région"]}),"")'
        ws[f["Groupe 1re occ."]] = (f'=IF(AND({f["Groupe"]}<>"",COUNTIF(${col("Groupe")}$2:{f["Groupe"]},{f["Groupe"]})=1),1,0)')
        g = f["Groupe"]
        ws[f["Groupe : nb évaluées"]] = f'=IF({f["Groupe 1re occ."]}=1,COUNTIFS(dGroupe,{g},dEvaluable,1),"")'
        ws[f["Groupe : délai consommé"]] = (f'=IF({f["Groupe 1re occ."]}=1,IF(SUMIFS(dDuree,dGroupe,{g})>0,'
                                            f'SUMIFS(dEcoule,dGroupe,{g})/SUMIFS(dDuree,dGroupe,{g})*100,-1),"")')
        nb = f["Groupe : nb évaluées"]
        ws[f["Groupe : exécution"]] = (f'=IF({f["Groupe 1re occ."]}=1,IF({nb}=0,-1,IF(SUMIFS(dMontEval,dGroupe,{g})>0,'
                                       f'SUMIFS(dMontTaux,dGroupe,{g})/SUMIFS(dMontEval,dGroupe,{g}),'
                                       f'SUMIFS(dTauxEval,dGroupe,{g})/{nb})),"")')
        cons = f["Groupe : délai consommé"]
        ws[f["Groupe : rang"]] = (f'=IF({f["Groupe 1re occ."]}=1,COUNTIFS(dPremier,1,dGConso,">"&{cons})'
                                  f'+COUNTIFS(${col("Groupe 1re occ.")}$2:{f["Groupe 1re occ."]},1,'
                                  f'${col("Groupe : délai consommé")}$2:{cons},{cons}),"")')
        ws[f["Groupe : libellé"]] = (f'=IF({f["Groupe 1re occ."]}=1,IF(LEN({g})>24,LEFT({g},23)&"…",{g})'
                                     f'&IF({cons}<0,"  – n.d.",""),"")')

    noms = {"dFiltre": "Dans le filtre", "dAlerte": "Alerte", "dEcart": "Écart", "dGroupe": "Groupe", "dEvaluable": "Évaluable",
            "dDuree": "Durée (éval.)", "dEcoule": "Écoulé (éval.)", "dMontEval": "Montant (éval.)", "dMontTaux": "Montant×taux (éval.)",
            "dTauxEval": "Taux (éval.)", "dPremier": "Groupe 1re occ.", "dGConso": "Groupe : délai consommé", "dRang": "Rang alerte",
            "dNum": "N°", "dCode": "Code", "dNature": "Nature", "dRegion": "Région", "dEntreprise": "Entreprise", "dTaux": "Taux",
            "dPges": "PGES", "dOngPremier": "ONG 1re occ.", "dOngSansNom": "ONG sans nom", "dMontant": "Montant", "dDecaisse": "Décaissé",
            "dGRang": "Groupe : rang", "dGLibelle": "Groupe : libellé", "dGExec": "Groupe : exécution"}
    for nom, colonne in noms.items():
        wb.defined_names[nom] = DefinedName(nom, attr_text=_rng(colonne))

    # listes des filtres
    lc = LISTES_COL
    i0 = COLS.index("Groupe : libellé") + 3
    listes = {"Nature": [TOUTES["nature"], *cfg.categories.values()],
              "Région": [TOUTES["region"], *sorted(cfg.regions.values(), key=str.casefold)],
              "Entreprise": [TOUTES["entreprise"], *sorted({l.entreprise for l in lg if l.entreprise}, key=str.casefold)],
              "Statut": [TOUTES["statut"], *STATUTS]}
    for j, (titre, valeurs) in enumerate(listes.items()):
        col = get_column_letter(i0 + j)
        ws[f"{col}1"] = f"Liste : {titre}"
        ws[f"{col}1"].font, ws[f"{col}1"].fill = gras, _fill(MARINE)
        ws.column_dimensions[col].width = 26
        for k, v in enumerate(valeurs, 2):
            ws[f"{col}{k}"] = v
        wb.defined_names[f"lst{titre.replace('é', 'e')}"] = DefinedName(
            f"lst{titre.replace('é', 'e')}", attr_text=f"'{DO}'!${col}$2:${col}${len(valeurs) + 1}")

    # données du graphique, triées par délai consommé décroissant
    g0 = i0 + len(listes) + 1
    gl, gc, ge, gp = (get_column_letter(g0 + k) for k in range(4))
    for col, titre in ((gl, "Graphique : libellé"), (gc, "Consommation du délai (%)"), (ge, "Exécution physique (%)"),
                       (gp, "Délai prévu atteint (100 %)")):
        ws[f"{col}1"] = titre
        ws[f"{col}1"].font, ws[f"{col}1"].fill = gras, _fill(MARINE)
        ws.column_dimensions[col].width = 26
    for k in range(1, wb._n_groupes + 1):
        r = k + 1
        ws[f"{gl}{r}"] = f'=IFERROR(INDEX(dGLibelle,MATCH({k},dGRang,0)),"")'
        ws[f"{gc}{r}"] = f'=IF({gl}{r}="",-1,INDEX(dGConso,MATCH({k},dGRang,0)))'
        ws[f"{ge}{r}"] = f'=IF({gl}{r}="",-1,INDEX(dGExec,MATCH({k},dGRang,0)))'
        ws[f"{gp}{r}"] = f'=IF({gl}{r}="",NA(),100)'
    wb._graphique = (ws, gl, gc, ge, gp)                                   # repris par la feuille TABLEAU DE BORD
    ws.freeze_panes = "C2"


# --- feuille TABLEAU DE BORD ------------------------------------------------------------------------------------
LARGEURS = {"A": 1.5, "B": 26, "C": 1.5}
CARTES = [("D", "F"), ("G", "I"), ("J", "L"), ("M", "O"), ("P", "R")]   # 5 cartes de 3 colonnes ; aperçu des alertes : S:W
APERCU = {"S": 11, "T": 14, "U": 30, "V": 16, "W": 8}


def _tableau_de_bord(wb, ws, n_act, maintenant, s_taux, s_conso):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    for col, w in LARGEURS.items():
        ws.column_dimensions[col].width = w
    for a, b in CARTES:
        for k in range(ord(a), ord(b) + 1):
            ws.column_dimensions[chr(k)].width = 11
    for col, w in APERCU.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["X"].width = 2

    # bandeau
    for r in range(1, 6):
        ws.row_dimensions[r].height = 18
        for k in range(1, 25):
            ws.cell(r, k).fill = _fill(BANDEAU)
    if LOGO.is_file():
        from openpyxl.drawing.image import Image
        img = Image(str(LOGO))
        img.width = img.height = 92
        ws.add_image(img, "B1")
    ws["D2"] = "TABLEAU DE BORD DE SUIVI DES ACTIVITÉS — PCR-CI"
    ws["D2"].font = Font(size=18, bold=True, color=TEXTE)
    ws["D4"] = "Vue d'ensemble de l'avancement physique, financier et des points d'attention."
    ws["D4"].font = Font(size=11, color=TEXTE)
    ws["S1"] = "DATE DE MAJ"
    ws["S1"].font = Font(size=14, bold=True, color=TEXTE)
    ws["S2"] = maintenant.strftime("%d/%m/%Y  %H:%M")
    ws["S2"].font = Font(size=14, bold=True, color=TEXTE)
    ws["S4"] = '=COUNTA(dCode)&" activités suivies"'
    ws["S4"].font = Font(size=14, bold=True, color=TEXTE)

    # filtres
    ws["B7"] = "▼  FILTRES"
    ws["B7"].font = Font(size=13, bold=True, color=MARINE)
    validations = []
    for (lib, cle, liste), r in zip((("Nature du marché", "nature", "lstNature"), ("Région", "region", "lstRegion"),
                                     ("Entreprise", "entreprise", "lstEntreprise"), ("Statut du chantier", "statut", "lstStatut")),
                                    (9, 12, 15, 18)):
        ws[f"B{r}"] = lib
        ws[f"B{r}"].font = Font(size=10, bold=True, color=TEXTE)
        cell = ws[f"B{r + 1}"]
        cell.value = TOUTES[cle]
        cell.font, cell.fill = Font(size=10, bold=True, color=BLEU), _fill(FOND_FILTRE)
        _bord(cell, tout=True)
        nom = {"nature": "selNature", "region": "selRegion", "entreprise": "selEntreprise", "statut": "selStatut"}[cle]
        wb.defined_names[nom] = DefinedName(nom, attr_text=f"'{TB}'!$B${r + 1}")
        dv = DataValidation(type="list", formula1=f"={liste}", allow_blank=True, showDropDown=False)
        dv.add(f"B{r + 1}")
        validations.append(dv)
    for dv in validations:
        ws.add_data_validation(dv)
    wb.defined_names["parEntreprise"] = DefinedName("parEntreprise", attr_text=f"AND(selRegion<>\"\",selRegion<>\"{TOUTES['region']}\")")
    ws["B21"] = '=IF(COUNTIF(dFiltre,1)=0,"⚠ Aucune activité ne correspond à ces filtres. Choisir « Toutes… » pour réinitialiser.","")'
    ws["B21"].font = Font(size=9, italic=True, color=ROUGE)
    ws["B21"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B21:B23")

    # cartes
    sel = 'COUNTIF(dFiltre,1)'
    cartes = [
        ("Activités sélectionnées / total", f'={sel}&" / "&COUNTA(dCode)', MARINE, VERT, 24, None, []),
        ("Taux d'exécution physique moyen", '=IFERROR(AVERAGEIFS(dTaux,dFiltre,1),"—")', ORANGE, BLEU, 24, '0.0"%"', []),
        ("Situation des PGES", f'=COUNTIFS(dFiltre,1,dPges,"Validé")&" Validés"', VERT, VERT, 20, None,
         ['="ONG mobilisées : "&SUM(dOngPremier)',
          '=IF(SUMIFS(dOngSansNom,dFiltre,1)>0,SUMIFS(dOngSansNom,dFiltre,1)&" activité(s) avec ONG non nommée","")']),
        ("Montant total des marchés", '=SUMIFS(dMontant,dFiltre,1)/1000000000', MARINE, BLEU, 20, '0.00" Mds FCFA"',
         ['="Montant décaissé : "&ROUND(SUMIFS(dDecaisse,dFiltre,1)/1000000,1)&" M FCFA"',
          '=IFERROR("Taux d\'exécution financière : "&ROUND(SUMIFS(dDecaisse,dFiltre,1)/SUMIFS(dMontant,dFiltre,1)*100,1)&" %","")']),
        ("Activités en alerte  ▸ détail", f'=SUMIFS(dAlerte,dFiltre,1)&" / "&{sel}', ROUGE, ROUGE, 24, None,
         [f'="Exécution < {s_taux} % ET délai consommé ≥ {s_conso} %"']),
    ]
    for (a, b), (titre, formule, couleur, bord, taille, fmt, sous) in zip(CARTES, cartes):
        ws.merge_cells(f"{a}7:{b}7")
        ws.merge_cells(f"{a}8:{b}9")
        ws[f"{a}7"] = titre
        ws[f"{a}7"].font = Font(size=10, bold=True, color=TEXTE)
        ws[f"{a}8"] = formule
        ws[f"{a}8"].font = Font(size=taille, bold=True, color=couleur)
        ws[f"{a}8"].alignment = Alignment(horizontal="left", vertical="center")
        if fmt:
            ws[f"{a}8"].number_format = fmt
        for k, f in enumerate(sous):
            fin = 11 if len(sous) == 1 else 10 + k
            ws.merge_cells(f"{a}{10 + k}:{b}{fin}")
            ws[f"{a}{10 + k}"] = f
            ws[f"{a}{10 + k}"].font = Font(size=9 if k else 10, bold=not k, color=TEXTE if k else couleur)
            ws[f"{a}{10 + k}"].alignment = Alignment(wrap_text=True, vertical="top")
        for r in range(7, 12):
            _bord(ws[f"{a}{r}"], gauche=bord)
    ws["D7"].hyperlink = None
    ws["P7"].hyperlink = f"#'{AL}'!A1"
    # couleur du taux : vert > 50 %, rouge < 10 %, orange entre les deux (seuils de l'ancien tableau de bord)
    ws.conditional_formatting.add("G8", CellIsRule(operator="greaterThan", formula=["50"], font=Font(color=VERT, bold=True, size=24)))
    ws.conditional_formatting.add("G8", CellIsRule(operator="lessThan", formula=["10"], font=Font(color=ROUGE, bold=True, size=24)))

    # aperçu des alertes
    ws.merge_cells("S7:W7")
    ws["S7"] = "Aperçu des alertes  ▸ voir tout"
    ws["S7"].hyperlink = f"#'{AL}'!A1"
    ws["S7"].font = Font(size=10, bold=True, color=TEXTE)
    ws.merge_cells("S8:W9")
    ws["S8"] = '=SUMIFS(dAlerte,dFiltre,1)&" activité(s) en alerte"'
    ws["S8"].font = Font(size=18, bold=True, color=ROUGE)
    ws["S8"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("S10:W11")
    ws["S10"] = f"Retard potentiel d'exécution : exécution < {s_taux} % ET délai consommé ≥ {s_conso} %"
    ws["S10"].font = Font(size=9, italic=True, color=GRIS)
    ws["S10"].alignment = Alignment(wrap_text=True, vertical="top")
    for r in range(7, 12):
        _bord(ws[f"S{r}"], gauche=ROUGE)

    entete = Font(size=10, bold=True, color="FFFFFF")
    for col, titre in zip("STUVW", ("N° activité", "Région", "Entreprise", "Nature", "Taux")):
        ws[f"{col}13"] = titre
        ws[f"{col}13"].font, ws[f"{col}13"].fill = entete, _fill(MARINE)
    for k in range(1, 6):
        r = 13 + k
        m = f"MATCH({k},dRang,0)"
        ws[f"S{r}"] = f'=IFERROR("P-"&TEXT(INDEX(dNum,{m}),"000"),"")'
        ws[f"T{r}"] = f'=IFERROR(INDEX(dRegion,{m}),"")'
        ws[f"U{r}"] = f'=IFERROR(INDEX(dEntreprise,{m})&"","")'
        ws[f"V{r}"] = f'=IFERROR(INDEX(dNature,{m}),"")'
        ws[f"W{r}"] = f'=IFERROR(INDEX(dTaux,{m})/100,"")'
        ws[f"W{r}"].number_format = "0%"
        for col in "STUVW":
            ws[f"{col}{r}"].font = Font(size=10, color=ROUGE if col == "W" else TEXTE)
            ws[f"{col}{r}"].border = Border(bottom=Side(style="thin", color=TRAIT))
    ws.merge_cells("S19:W19")
    ws["S19"] = ('=IF(SUMIFS(dAlerte,dFiltre,1)>5,"… et "&(SUMIFS(dAlerte,dFiltre,1)-5)&" autre(s) activité(s) en alerte : voir le détail ▸",'
                 'IF(SUMIFS(dAlerte,dFiltre,1)=0,"Aucune activité en alerte dans le périmètre.",""))')
    ws["S19"].hyperlink = f"#'{AL}'!A1"
    ws["S19"].font = Font(size=10, bold=True, underline="single", color=BLEU)
    ws.merge_cells("S21:W24")
    ws["S21"] = ('="Activités évaluées : "&SUMIFS(dEvaluable,dFiltre,1)&" sur "&' + sel +
                 '&" (taux d\'exécution ET dates disponibles). Les autres ne sont pas comptées dans les alertes."')
    ws["S21"].font = Font(size=9, italic=True, color=GRIS)
    ws["S21"].alignment = Alignment(wrap_text=True, vertical="top")

    # graphique
    ws.merge_cells("D13:R13")
    ws["D13"] = ('="Consommation des délais par "&IF(parEntreprise,"entreprise","région")'
                 '&"   (délai consommé vs exécution physique)"')
    ws["D13"].font, ws["D13"].fill = Font(size=11, bold=True, color="FFFFFF"), _fill(MARINE)
    do, gl, gc, ge, gp = wb._graphique
    bar = BarChart()
    bar.type, bar.grouping, bar.gapWidth, bar.overlap = "col", "clustered", 60, -5
    cats = Reference(do, range_string=f"'{DO}'!${gl}$2:${gl}${wb._n_groupes + 1}")
    for col, couleur in ((gc, ORANGE), (ge, BLEU)):
        bar.add_data(Reference(do, range_string=f"'{DO}'!${col}$1:${col}${wb._n_groupes + 1}"), titles_from_data=True)
        s = bar.series[-1]
        s.graphicalProperties = GraphicalProperties(solidFill=couleur)
        s.graphicalProperties.line.noFill = True
    bar.set_categories(cats)
    bar.dataLabels = DataLabelList()
    bar.dataLabels.showVal = True
    for k in ("showSerName", "showCatName", "showLegendKey", "showPercent", "showBubbleSize", "showLeaderLines"):
        setattr(bar.dataLabels, k, False)
    bar.dataLabels.position = "outEnd"
    bar.dataLabels.numFmt = '0"%";;0"%"'                                   # -1 (non calculable) : ni barre ni étiquette
    bar.y_axis.scaling.min = 0
    bar.y_axis.numFmt = '0"%"'
    bar.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E8EF"))
    bar.y_axis.delete = False
    bar.x_axis.delete = False
    bar.x_axis.tickLblPos = "low"
    line = LineChart()
    line.add_data(Reference(do, range_string=f"'{DO}'!${gp}$1:${gp}${wb._n_groupes + 1}"), titles_from_data=True)
    ls = line.series[0]
    ls.graphicalProperties.line.solidFill = GRIS
    ls.graphicalProperties.line.dashStyle = "dash"
    ls.graphicalProperties.line.width = 19050
    from openpyxl.chart.marker import Marker
    ls.marker = Marker(symbol="none")
    ls.smooth = False
    bar += line
    bar.legend.position = "t"
    bar.height, bar.width = 10.5, 31
    ws.add_chart(bar, "D14")

    ws.merge_cells("D36:W38")
    ws["D36"] = ("Méthode : consommation du délai = Σ temps écoulé ÷ Σ durées prévues (depuis la date de début des travaux) ; exécution "
                 "physique pondérée par les montants des marchés (moyenne simple si montants absents). Statut : celui saisi par "
                 "l'administrateur, sinon déduit du taux. Délai non calculable pour une activité achevée, sans date de début ou sans fin prévue.")
    ws["D36"].font = Font(size=9, italic=True, color=GRIS)
    ws["D36"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.print_area = "A1:W38"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# --- feuille ALERTES ----------------------------------------------------------------------------------------------
def _alertes(ws, alertes, maintenant, s_taux, s_conso):
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"Alertes et points d’attention — Voir tout ({len(alertes)})"
    ws["A1"].font = Font(size=18, bold=True, color=MARINE)
    ws["A2"] = (f"Périmètre : toutes les activités  ·  Règle : exécution < {s_taux} % ET délai consommé ≥ {s_conso} %  ·  "
                f"Mise à jour : {maintenant:%d/%m/%Y %H:%M}  ·  Filtrer avec les flèches des en-têtes")
    ws["A2"].font = Font(size=10, color=GRIS)
    ws["A3"] = "←  Retour au tableau de bord"
    ws["A3"].hyperlink = f"#'{TB}'!A1"
    ws["A3"].font = Font(size=11, bold=True, underline="single", color=BLEU)
    entetes = ["N° activité", "Région", "Département", "Entreprise", "Date de début des travaux", "Taux d’exécution",
               "Délai consommé", "Observation technique et remarque"]
    largeurs = [12, 14, 20, 32, 14, 11, 11, 90]
    for i, (t, w) in enumerate(zip(entetes, largeurs), 1):
        c = ws.cell(5, i, t)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), _fill(MARINE)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[5].height = 32
    for r, a in enumerate(alertes, 6):
        obs = "; ".join(x.strip(" -") for x in a.observations.splitlines() if x.strip(" -")) or "RAS"
        remarque = (f"Retard potentiel d'exécution — délai consommé : {a.conso:.0f} % pour {a.taux:.0f} % d'exécution. "
                    f"Observation : {obs}")
        vals = [a.numero, a.region, a.departement.title(), a.entreprise or "—", a.debut, a.taux / 100, a.conso / 100, remarque]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = Border(bottom=Side(style="thin", color=TRAIT))
        ws.cell(r, 5).number_format = "DD/MM/YYYY"
        ws.cell(r, 6).number_format = ws.cell(r, 7).number_format = "0%"
        ws.cell(r, 6).font = Font(color=ROUGE, bold=True)
        ws.row_dimensions[r].height = 15 * max(2, -(-len(remarque) // 95))   # hauteur selon la longueur du texte
    if alertes:
        t = Table(displayName="Alertes", ref=f"A5:H{5 + len(alertes)}")
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        ws.add_table(t)
    else:
        ws["A6"] = "Aucune activité en alerte."
    ws.freeze_panes = "A6"
    ws.print_title_rows = "5:5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
