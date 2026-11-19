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

from test_phase6_7 import ref  # noqa: F401,E402
from pcr.referentiel import cle_ong, coerce, ong_canonique
from pcr.tableau_de_bord_xlsx import generer
from pcr.word_gen import ordered_activities

LO = pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")


def calcule(ref, tmp_path, **filtres):
    """Génère le tableau de bord, choisit les filtres (B10 nature · B13 région · B16 entreprise · B19 statut), le fait recalculer
    par LibreOffice et retourne le classeur avec les valeurs calculées."""
    out = tmp_path / f"d_{len(list(tmp_path.glob('d_*.xlsx')))}.xlsx"
    generer(ref, out, maintenant=datetime(2026, 11, 6, 8, 0))
    if filtres:
        wb = load_workbook(out)
        for cell, v in filtres.items():
            wb["TABLEAU DE BORD"][cell] = v
        wb.save(out)
    lo = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(lo), str(out)], capture_output=True, timeout=250)
    return load_workbook(lo / out.name, data_only=True)


@LO
def test_carte_activites_selection_sur_total(ref, tmp_path):
    ref.update_activity(ref.activities()[3]["CODE_ACTIVITE"], "admin", {"Statut_Chantier": "Suspendu"})
    n = len(ref.activities())
    wb = calcule(ref, tmp_path)
    assert wb["TABLEAU DE BORD"]["D7"].value == "Activités sélectionnées / total"
    assert wb["TABLEAU DE BORD"]["D8"].value == f"{n} / {n}"                # sans filtre : tout, même « Suspendu »
    do = wb["DONNÉES"]
    col = [c.value for c in do[1]].index("Statut") + 1
    statuts = [do.cell(r, col).value for r in range(2, 2 + n)]
    for statut in ("En cours", "Suspendu", "Non démarré"):
        attendu = statuts.count(statut)
        got = calcule(ref, tmp_path, B19=statut)["TABLEAU DE BORD"]["D8"].value
        assert got == f"{attendu} / {n}", (statut, got, attendu)
    mr = len(ordered_activities(ref, "MR"))
    assert calcule(ref, tmp_path, B10="Marchés ruraux")["TABLEAU DE BORD"]["D8"].value == f"{mr} / {n}"


@LO
def test_ong_gerees_par_l_admin_et_noms_proches(ref, tmp_path):
    bs = [a["CODE_ACTIVITE"] for a in ordered_activities(ref, "BS")][:5]
    for code, ong in zip(bs, ["ASAPSU", "ASAPSU ONG", "RIDDEF", "Mobilisée", "Non mobilisée"]):
        ref.update_activity(code, "admin", {"ONG_Mobilisees": ong})
    liste = ref.ong_liste()
    proches = {g["cle"]: g["proches"] for g in liste}
    assert proches["ASAPSU"] == ["ASAPSU ONG"] and proches["RIDDEF"] == []        # doublon probable signalé, ONG distincte non signalée
    avant = calcule(ref, tmp_path)["TABLEAU DE BORD"]
    assert avant["J10"].value == "ONG mobilisées : 3"                              # ASAPSU, ASAPSU ONG, RIDDEF : « ASAPSU ONG » compte à part
    assert "1 activité(s) avec ONG non nommée" in avant["J11"].value               # « Mobilisée » sans nom ; « Non mobilisée » ignorée
    codes = ref.renommer_ong("ASAPSU ONG", "ASAPSU")
    assert codes == [bs[1]] and ref.get(bs[1])["ONG_Mobilisees"] == "ASAPSU"
    assert any(h["Champ_Modifie"] == "ONG_Mobilisees" and h["Nouvelle_Valeur"] == "ASAPSU" for h in ref.history(bs[1]))
    apres = calcule(ref, tmp_path)["TABLEAU DE BORD"]
    assert apres["J10"].value == "ONG mobilisées : 2"
    assert cle_ong("Asapsu") == cle_ong("ASAPSU") == "ASAPSU" and ref.renommer_ong("RIDDEF", "RIDDEF") == []
    with pytest.raises(ValueError):
        ref.renommer_ong("RIDDEF", "  ")


def test_valeurs_generiques_d_ong_harmonisees():
    for v in ("mobilisée", "Mobilisee", "OUI", "à confirmer", "ONG mobilisée"):
        assert coerce("ONG_Mobilisees", v) == "Mobilisée", v
    for v in ("non mobilisée", "Non mobilisee", "Aucune", "néant", "sans ONG", "pas d'ONG"):
        assert coerce("ONG_Mobilisees", v) == "Non mobilisée", v
    assert coerce("ONG_Mobilisees", "  ONG  VOIX  DE FEMME ") == "ONG VOIX DE FEMME" and ong_canonique("Solidarité") == "Solidarité"
