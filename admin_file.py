"""Fichier d'administration (Excel) : l'administrateur y crée et corrige les activités.

  - ligne avec CODE vide  -> nouvelle activité (Categorie et Region obligatoires) : le code est généré à l'application ;
  - ligne avec CODE       -> modification des informations fixes (écart historisé) ;
  - cellule laissée vide  -> ignorée (jamais une suppression) ; pour effacer une valeur, écrire  #VIDER  dans la cellule.
"""
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from .referentiel import EDITABLE, FIXED, STATUTS, Referentiel, coerce

CLEAR = "#VIDER"
COLUMNS = ["CODE_ACTIVITE", "Categorie", "Region"] + FIXED + ["Date_Fin_Previsionnelle", "Statut"]
INFO_ONLY = {"CODE_ACTIVITE", "Date_Fin_Previsionnelle"}      # produits par le système, ignorés à la lecture
FILLS = {"fixe": PatternFill("solid", fgColor="1F4E78"), "info": PatternFill("solid", fgColor="7F7F7F")}
WIDTHS = {"Designation": 60, "Localite": 40, "Entreprise": 32, "Departement": 24, "Bureau_Controle": 30}


@dataclass
class Report:
    created: list = field(default_factory=list)      # [(ligne, code)]
    updated: list = field(default_factory=list)      # [(code, [(champ, ancien, nouveau)])]
    errors: list = field(default_factory=list)       # [(ligne, message)]

    def lines(self) -> list[str]:
        out = [f"{len(self.created)} activité(s) créée(s), {len(self.updated)} modifiée(s), {len(self.errors)} erreur(s)."]
        out += [f"  + ligne {r} : {c} créé" for r, c in self.created]
        out += [f"  ~ {c} : {f} : {o!r} -> {n!r}" for c, ch in self.updated for f, o, n in ch]
        out += [f"  ! ligne {r} : {m}" for r, m in self.errors]
        return out


def export_admin_file(ref: Referentiel, path) -> Path:
    """Écrit le fichier d'édition à partir de la base (à relancer après chaque application pour voir les codes générés)."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Activites"
    for c, h in enumerate(COLUMNS, 1):
        cell = ws.cell(1, c, h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = FILLS["info"] if h in INFO_ONLY else FILLS["fixe"]
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cell.column_letter].width = WIDTHS.get(h, 18)
    for r, a in enumerate(ref.activities(), 2):
        for c, h in enumerate(COLUMNS, 1):
            cell = ws.cell(r, c, a[h])
            cell.font = Font(name="Arial", size=10, color="7F7F7F" if h in INFO_ONLY else "000000")
            if h in ("Date_OS", "Date_Debut", "Date_Fin_Previsionnelle"):
                cell.number_format = "DD/MM/YYYY"
            elif h == "Montant_Marche":
                cell.number_format = "#,##0"
            if h in ("Designation", "Localite"):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{ws.cell(1, len(COLUMNS)).column_letter}{max(ws.max_row, 2)}"

    lst = wb.create_sheet("Listes")
    lst.append(["Catégories (code)", "Nom", "", "Régions (code)", "Nom", "", "Statuts"])
    cats, regs = list(ref.cfg.categories.items()), list(ref.cfg.regions.items())
    for i in range(max(len(cats), len(regs), len(STATUTS))):
        lst.append([*(cats[i] if i < len(cats) else ("", "")), "", *(regs[i] if i < len(regs) else ("", "")), "",
                    STATUTS[i] if i < len(STATUTS) else ""])
    for col, rng, n in (("Categorie", "A", len(cats)), ("Region", "D", len(regs)), ("Statut", "G", len(STATUTS))):
        dv = DataValidation(type="list", formula1=f"=Listes!${rng}$2:${rng}${n + 1}", allow_blank=True)
        ws.add_data_validation(dv)
        letter = ws.cell(1, COLUMNS.index(col) + 1).column_letter
        dv.add(f"{letter}2:{letter}1000")

    guide = wb.create_sheet("Mode_emploi")
    for line in [
        "MODE D'EMPLOI",
        "1. Nouvelle activité : ajouter une ligne en bas, laisser CODE_ACTIVITE vide, choisir Categorie et Region (listes déroulantes).",
        "   Les autres champs peuvent rester vides : ils pourront être complétés plus tard.",
        "2. Modifier une activité : changer la cellule sur sa ligne (ne jamais modifier CODE_ACTIVITE, Categorie ni Region).",
        "3. Une cellule laissée VIDE n'efface rien. Pour effacer une valeur, écrire  #VIDER  dans la cellule.",
        "4. Enregistrer le fichier puis lancer :  python -m pcr appliquer",
        "5. Le rapport indique les créations, les modifications (ancienne -> nouvelle valeur) et les erreurs ligne par ligne.",
        "Colonnes grises : calculées ou attribuées par le système, ignorées à la lecture.",
        "Dates au format JJ/MM/AAAA. Montant en FCFA, sans texte. Délai en mois entiers.",
    ]:
        guide.append([line])
    guide.column_dimensions["A"].width = 120
    guide["A1"].font = Font(name="Arial", bold=True)
    wb.save(path)
    return path


def apply_admin_file(ref: Referentiel, path, user: str | None = None) -> Report:
    """Applique le fichier d'édition à la base (en mémoire : l'appelant sauvegarde). Les erreurs sont ligne par ligne,
    les lignes valides sont appliquées."""
    rep = Report()
    ws = load_workbook(path, data_only=True)["Activites"]
    headers = [c.value for c in ws[1]]
    missing = [h for h in ("CODE_ACTIVITE", "Categorie", "Region") if h not in headers]
    if missing:
        raise ValueError(f"Colonnes obligatoires absentes du fichier : {missing}")
    idx = {h: i for i, h in enumerate(headers) if h}
    for r, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        row = {h: values[i] for h, i in idx.items() if i < len(values)}
        data = {h: v for h, v in row.items() if h not in INFO_ONLY and v is not None and str(v).strip() != ""}
        code = str(row.get("CODE_ACTIVITE") or "").strip().upper()
        if not data and not code:
            continue
        try:
            if not code:
                if "Categorie" not in data or "Region" not in data:
                    raise ValueError("Categorie et Region sont obligatoires pour créer une activité")
                fields = {k: v for k, v in data.items() if k in EDITABLE and str(v).strip() != CLEAR}
                rep.created.append((r, ref.create_activity(data["Categorie"], data["Region"], user, **fields)))
                continue
            act = ref.get(code)
            for col, key in (("Categorie", ref.cfg.category_code), ("Region", ref.cfg.region_code)):
                if col in data and key(data[col]) != act[col]:
                    raise ValueError(f"{col} non modifiable ({act[col]} -> {data[col]}) : elle fait partie du code {code}")
            changes = {k: v for k, v in data.items() if k in EDITABLE and str(v).strip() != CLEAR}
            clear = [k for k, v in data.items() if k in EDITABLE and str(v).strip() == CLEAR]
            for k, v in changes.items():
                coerce(k, v)                                   # validation complète avant toute écriture
            done = ref.update_activity(code, user, changes, clear)
            if done:
                rep.updated.append((code, done))
        except (ValueError, KeyError) as e:
            rep.errors.append((r, str(e.args[0]) if e.args else str(e)))
    return rep
