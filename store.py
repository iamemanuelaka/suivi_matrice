"""Base de données = un classeur Excel (data/base_pcr.xlsx), une feuille par table.

Choix : lisible et dépannable à la main, persistant dans le dépôt, aucun service à gérer.
Le fichier est écrit de façon atomique (fichier temporaire puis remplacement) et peut être sauvegardé avant écriture.
"""
import os
import shutil
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

SCHEMA = {
    "Activites": [
        "CODE_ACTIVITE", "Categorie", "Region", "Departement", "Localite", "Designation", "Montant_Marche", "Entreprise",
        "Numero_OS", "Date_OS", "Date_Debut", "Delai_Mois", "Bureau_Controle", "Date_Fin_Previsionnelle", "Statut",
        "Date_Creation", "Date_Modification",
    ],
    "Suivi_Activites": [
        "ID_SOUMISSION_KOBO", "CODE_ACTIVITE", "Date_Actualisation", "Observations_Techniques", "Taux_Execution_Physique",
        "Commentaires", "Difficultes", "Recommandations", "Date_Importation", "Statut_Traitement", "Motif_Rejet",
    ],
    "Historique_Modifications": [
        "ID_MODIFICATION", "CODE_ACTIVITE", "Champ_Modifie", "Ancienne_Valeur", "Nouvelle_Valeur", "Date_Modification", "Utilisateur",
    ],
    "Compteurs_Codes": ["Categorie", "Region", "Dernier_Numero"],
    "Journal_Executions": [
        "ID_Execution", "Date_Execution", "Declencheur", "Recuperees", "Traitees", "Ignorees", "Echouees", "Documents", "Email", "Resultat", "Detail",
    ],
}
DATE_COLS = {"Date_OS", "Date_Debut", "Date_Fin_Previsionnelle", "Date_Actualisation"}
DATETIME_COLS = {"Date_Creation", "Date_Modification", "Date_Importation", "Date_Execution"}
INT_COLS = {"Montant_Marche", "Delai_Mois", "Dernier_Numero", "ID_MODIFICATION", "ID_Execution", "Recuperees", "Traitees", "Ignorees", "Echouees"}
WIDTHS = {"Designation": 60, "Localite": 40, "Entreprise": 32, "Observations_Techniques": 60, "Commentaires": 40, "Difficultes": 40,
          "Recommandations": 40, "Ancienne_Valeur": 30, "Nouvelle_Valeur": 30, "Detail": 50, "Bureau_Controle": 30, "Departement": 24}
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")


def _read(col, v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if col in DATE_COLS and isinstance(v, datetime):
        return v.date()
    if col in INT_COLS and isinstance(v, float) and v.is_integer():
        return int(v)
    return v


class Base:
    def __init__(self, path):
        self.path = Path(path)
        if self.path.exists():
            self.wb = load_workbook(self.path)
        else:
            self.wb = Workbook()
            self.wb.remove(self.wb.active)
        for name, cols in SCHEMA.items():
            self._ensure(name, cols)

    def _ensure(self, name, cols):
        if name not in self.wb.sheetnames:
            ws = self.wb.create_sheet(name)
            for c, h in enumerate(cols, 1):
                cell = ws.cell(1, c, h)
                cell.font, cell.fill = Font(name="Arial", bold=True, color="FFFFFF"), HEADER_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                ws.column_dimensions[cell.column_letter].width = WIDTHS.get(h, 18)
            ws.freeze_panes = "B2" if name == "Activites" else "A2"
            return
        found = [c.value for c in self.wb[name][1]][: len(cols)]
        if found != cols:
            raise ValueError(f"Feuille {name} : en-têtes inattendus {found} (attendu {cols}). Structure modifiée à la main ?")

    # --- lecture -----------------------------------------------------------
    def rows(self, sheet: str) -> list[dict]:
        ws, cols, out = self.wb[sheet], SCHEMA[sheet], []
        for r, values in enumerate(ws.iter_rows(min_row=2, max_col=len(cols), values_only=True), start=2):
            row = {c: _read(c, v) for c, v in zip(cols, values)}
            if any(v is not None for v in row.values()):
                row["_r"] = r
                out.append(row)
        return out

    # --- écriture ----------------------------------------------------------
    def _put(self, ws, r, c, col, v):
        cell = ws.cell(r, c)
        cell.value = v                      # affectation explicite : ws.cell(r, c, None) n'effacerait pas la cellule
        cell.font = Font(name="Arial", size=10)
        if col in DATE_COLS and v is not None:
            cell.number_format = "DD/MM/YYYY"
        elif col in DATETIME_COLS and v is not None:
            cell.number_format = "DD/MM/YYYY HH:MM"
        elif col == "Montant_Marche":
            cell.number_format = "#,##0"
        if col in ("Designation", "Observations_Techniques", "Localite", "Commentaires", "Difficultes", "Recommandations"):
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    def append(self, sheet: str, row: dict) -> int:
        ws, cols = self.wb[sheet], SCHEMA[sheet]
        unknown = set(row) - set(cols)
        if unknown:
            raise KeyError(f"Colonnes inconnues pour {sheet} : {sorted(unknown)}")
        r = ws.max_row + 1
        for c, col in enumerate(cols, 1):
            self._put(ws, r, c, col, row.get(col))
        return r

    def update(self, sheet: str, r: int, values: dict) -> None:
        ws, cols = self.wb[sheet], SCHEMA[sheet]
        for col, v in values.items():
            self._put(ws, r, cols.index(col) + 1, col, v)

    def save(self, backup_dir=None) -> None:
        """Écriture atomique. Si backup_dir est fourni, l'ancienne version est copiée avant remplacement."""
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if backup_dir and self.path.exists():
            Path(backup_dir).mkdir(parents=True, exist_ok=True)
            shutil.copy2(self.path, Path(backup_dir) / f"{self.path.stem}_{datetime.now():%Y%m%d-%H%M%S}.xlsx")
        tmp = self.path.with_suffix(".tmp.xlsx")
        self.wb.save(tmp)
        os.replace(tmp, self.path)
