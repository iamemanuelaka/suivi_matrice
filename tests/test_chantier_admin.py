"""Informations gérées par l'administrateur : finances (décaissé), statut du chantier (4 valeurs), PGES, ONG, dates."""
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref  # noqa: F401,E402
from pcr.admin_file import CLEAR, COLUMNS, INFO_ONLY, apply_admin_file, export_admin_file
from pcr.pipeline import alertes, run
from pcr.referentiel import STATUTS_CHANTIER, coerce, taux_financier
from pcr.store import Base
from pcr.word_gen import ordered_activities


def ligne_de(ref, code):
    acts = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    return 3 + 1 + [a["CODE_ACTIVITE"] for a in acts].index(code)


def test_une_base_existante_sans_les_nouvelles_colonnes_est_migree(tmp_path):
    p = tmp_path / "ancienne.xlsx"
    b = Base(p)
    ws = b.wb["Activites"]
    ws.delete_cols(18, 6)                                     # ancien schéma : s'arrête à Date_Modification
    b.save()
    assert [c.value for c in load_workbook(p)["Activites"][1]][-1] == "Date_Modification"
    b2 = Base(p)
    assert [c.value for c in b2.wb["Activites"][1]][-6:] == ["Montant_Decaisse", "Date_Remise_Site", "Statut_Chantier", "PGES",
                                                              "ONG_Mobilisees", "Date_Reception_Provisoire"]


def test_les_quatre_statuts_seulement():
    assert STATUTS_CHANTIER == ("Non démarré", "En cours", "Achevé", "Suspendu")
    assert coerce("Statut_Chantier", " suspendu ") == "Suspendu" and coerce("Statut_Chantier", "ACHEVE") == "Achevé"
    for interdit in ("Abandonné", "OS à reprendre", "Démarrage autorisé", "consortium"):
        with pytest.raises(ValueError, match="statut du chantier invalide"):
            coerce("Statut_Chantier", interdit)
    assert coerce("PGES", "validé") == "Validé"
    with pytest.raises(ValueError):
        coerce("PGES", "peut-être")


def test_fichier_admin_champs_chantier(ref, tmp_path):
    f = tmp_path / "admin.xlsx"
    export_admin_file(ref, f)
    wb = load_workbook(f)
    ws = wb["Activites"]
    h = {c.value: i + 1 for i, c in enumerate(ws[1])}
    for col in ("Montant_Decaisse", "Taux_Financier", "Date_Remise_Site", "Statut_Chantier", "PGES", "ONG_Mobilisees", "Date_Reception_Provisoire"):
        assert col in h
    assert len(ws.data_validations.dataValidation) >= 5                                   # listes : catégorie, région, statut, statut chantier, PGES
    code = ws.cell(2, h["CODE_ACTIVITE"]).value
    montant = ref.get(code)["Montant_Marche"]
    ws.cell(2, h["Montant_Decaisse"]).value = montant // 4
    ws.cell(2, h["Statut_Chantier"]).value = "Suspendu"
    ws.cell(2, h["PGES"]).value = "Validé"
    ws.cell(2, h["ONG_Mobilisees"]).value = "ONG TEST"
    ws.cell(2, h["Date_Remise_Site"]).value = datetime(2026, 9, 1)
    ws.cell(2, h["Taux_Financier"]).value = 999                                           # colonne grise : ignorée
    ws.cell(3, h["Statut_Chantier"]).value = "Abandonné"                                  # refusé
    ws.cell(4, h["Montant_Decaisse"]).value = 10 ** 13                                    # dépasse le marché : refusé
    wb.save(f)
    rep = apply_admin_file(ref, f)
    a = ref.get(code)
    assert (a["Montant_Decaisse"], a["Statut_Chantier"], a["PGES"], a["ONG_Mobilisees"], a["Date_Remise_Site"]) == \
        (montant // 4, "Suspendu", "Validé", "ONG TEST", date(2026, 9, 1))
    assert taux_financier(a) == 25.0
    msgs = " | ".join(m for _, m in rep.errors)
    assert "statut du chantier invalide" in msgs and "dépasse le montant du marché" in msgs and len(rep.errors) == 2
    hist = [h for h in ref.history(code) if h["Champ_Modifie"] in ("Montant_Decaisse", "Statut_Chantier", "PGES")]
    assert len(hist) == 3 and all(h["Utilisateur"] and h["Date_Modification"] for h in hist)
    # #VIDER efface ; une cellule vide n'efface rien
    export_admin_file(ref, f)
    wb = load_workbook(f)
    ws = wb["Activites"]
    ws.cell(2, h["PGES"]).value = CLEAR
    ws.cell(2, h["ONG_Mobilisees"]).value = None
    wb.save(f)
    apply_admin_file(ref, f)
    a = ref.get(code)
    assert a["PGES"] is None and a["ONG_Mobilisees"] == "ONG TEST"


def test_les_alertes_python_ignorent_un_chantier_acheve_par_l_administrateur(ref):
    avant = alertes(ref, date(2026, 11, 6))
    assert avant
    code = avant[0].split(" ")[0]
    ref.update_activity(code, "admin", {"Statut_Chantier": "Achevé"})
    assert not any(x.startswith(code) for x in alertes(ref, date(2026, 11, 6)))             # comme le tableau de bord (activité achevée : non évaluée)
