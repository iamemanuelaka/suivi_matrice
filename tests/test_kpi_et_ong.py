"""Carte « Activités » (sélection / total) selon les filtres, et gestion des ONG par l'administrateur."""
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref, TEMPLATE  # noqa: F401,E402
from pcr.excel_gen import generate_dashboard
from pcr.migration_modele import migrer_valeurs_modele
from pcr.referentiel import cle_ong, coerce, ong_canonique
from pcr.word_gen import ordered_activities
from pcr.xlsx_patch import Package

pytestmark = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")


def calcule(ref, tmp_path, **filtres):
    out = tmp_path / f"d_{len(list(tmp_path.glob('d_*.xlsx')))}.xlsx"
    generate_dashboard(ref, TEMPLATE, out, datetime(2026, 11, 6, 8, 0))
    if filtres:
        pkg = Package(out)
        sh = pkg.sheet("TABLEAU DE BORD")
        for cell, v in filtres.items():
            sh.set(int(cell[1:]), 2, v)                      # B10 nature · B13 région · B16 entreprise · B19 statut
        pkg.save(out)
    lo = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(lo), str(out)], capture_output=True, timeout=250)
    return load_workbook(lo / out.name, data_only=True)


def test_carte_activites_selection_sur_total(ref, tmp_path):
    migrer_valeurs_modele(ref, TEMPLATE)                                    # donne un chantier « Suspendu »
    n = len(ref.activities())
    wb = calcule(ref, tmp_path)
    assert wb["TABLEAU DE BORD"]["G8"].value == "Activités sélectionnées / total"
    assert wb["TABLEAU DE BORD"]["G9"].value == f"{n} / {n}"                # sans filtre : tout, même « Suspendu »
    statuts = [wb["MATRICE DE SUIVI"].cell(r, 18).value for r in range(4, 4 + n)]
    for statut in ("En cours", "Suspendu", "Non démarré"):
        attendu = statuts.count(statut)
        got = calcule(ref, tmp_path, B19=statut)["TABLEAU DE BORD"]["G9"].value
        assert got == f"{attendu} / {n}", (statut, got, attendu)
    mr = len(ordered_activities(ref, "MR"))
    assert calcule(ref, tmp_path, B10="Marchés ruraux")["TABLEAU DE BORD"]["G9"].value == f"{mr} / {n}"


def test_ong_gerees_par_l_admin_et_noms_proches(ref, tmp_path):
    bs = [a["CODE_ACTIVITE"] for a in ordered_activities(ref, "BS")][:5]
    for code, ong in zip(bs, ["ASAPSU", "ASAPSU ONG", "RIDDEF", "Mobilisée", "Non mobilisée"]):
        ref.update_activity(code, "admin", {"ONG_Mobilisees": ong})
    liste = ref.ong_liste()
    proches = {g["cle"]: g["proches"] for g in liste}
    assert proches["ASAPSU"] == ["ASAPSU ONG"] and proches["RIDDEF"] == []        # doublon probable signalé, ONG distincte non signalée
    avant = calcule(ref, tmp_path)["TABLEAU DE BORD"]
    assert avant["Q11"].value == "ONG mobilisées : 3"                              # ASAPSU, ASAPSU ONG, RIDDEF : « ASAPSU ONG » compte à part
    assert "1 activité(s) avec ONG non nommée" in avant["Q12"].value               # « Mobilisée » sans nom ; « Non mobilisée » ignorée
    codes = ref.renommer_ong("ASAPSU ONG", "ASAPSU")
    assert codes == [bs[1]] and ref.get(bs[1])["ONG_Mobilisees"] == "ASAPSU"
    assert any(h["Champ_Modifie"] == "ONG_Mobilisees" and h["Nouvelle_Valeur"] == "ASAPSU" for h in ref.history(bs[1]))
    apres = calcule(ref, tmp_path)["TABLEAU DE BORD"]
    assert apres["Q11"].value == "ONG mobilisées : 2"
    assert cle_ong("Asapsu") == cle_ong("ASAPSU") == "ASAPSU" and ref.renommer_ong("RIDDEF", "RIDDEF") == []
    with pytest.raises(ValueError):
        ref.renommer_ong("RIDDEF", "  ")


def test_valeurs_generiques_d_ong_harmonisees():
    for v in ("mobilisée", "Mobilisee", "OUI", "à confirmer", "ONG mobilisée"):
        assert coerce("ONG_Mobilisees", v) == "Mobilisée", v
    for v in ("non mobilisée", "Non mobilisee", "Aucune", "néant", "sans ONG", "pas d'ONG"):
        assert coerce("ONG_Mobilisees", v) == "Non mobilisée", v
    assert coerce("ONG_Mobilisees", "  ONG  VOIX  DE FEMME ") == "ONG VOIX DE FEMME" and ong_canonique("Solidarité") == "Solidarité"
