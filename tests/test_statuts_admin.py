"""Statuts des chantiers pilotés par l'administrateur, liés au tableau de bord et à ses filtres."""
import re
import shutil
import sys
import zipfile
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref, TEMPLATE  # noqa: F401,E402
from pcr.actions_masse import figer_statuts, statuts_a_verifier
from pcr.admin_file import COLUMNS, INFO_ONLY, apply_admin_file, export_admin_file
from pcr.pipeline import run
from pcr.referentiel import statut_deduit
from pcr.word_gen import latest_valid


def deduit(ref, code):
    s = latest_valid(ref).get(code)
    return statut_deduit(s["Taux_Execution_Physique"] if s else None)


def test_le_filtre_statut_propose_toujours_les_4_statuts():
    wb = zipfile.ZipFile(TEMPLATE).read("xl/workbook.xml").decode()
    assert re.search(r'<definedName name="lstStatuts"[^>]*>PARAMETRES!\$I\$12:\$I\$16<', wb)      # liste fixe, plus dépendante des données
    ws = load_workbook(TEMPLATE)["PARAMETRES"]
    assert [ws.cell(r, 9).value for r in range(12, 17)] == ["Tous les statuts", "Non démarré", "En cours", "Achevé", "Suspendu"]


def test_fichier_admin_colonne_suggestion_grise(ref, tmp_path):
    f = tmp_path / "a.xlsx"
    export_admin_file(ref, f)
    ws = load_workbook(f)["Activites"]
    h = {c.value: i + 1 for i, c in enumerate(ws[1])}
    assert h["Statut_Suggere"] == h["Statut_Chantier"] + 1 and "Statut_Suggere" in INFO_ONLY
    code = ws.cell(2, h["CODE_ACTIVITE"]).value
    assert ws.cell(2, h["Statut_Suggere"]).value == deduit(ref, code)
    ws.cell(2, h["Statut_Suggere"]).value = "Suspendu"                                          # colonne grise : ignorée à la lecture
    ws.parent.save(f)
    apply_admin_file(ref, f)
    assert ref.get(code)["Statut_Chantier"] is None


def test_figer_les_statuts_donne_la_main_a_l_admin(ref):
    avant = len(ref.history())
    simulation = figer_statuts(ref, dry_run=True)
    assert len(simulation) == len(ref.activities()) and all(a["Statut_Chantier"] is None for a in ref.activities()) and len(ref.history()) == avant
    faits = figer_statuts(ref)
    assert len(faits) == len(ref.activities())
    assert all(ref.get(c)["Statut_Chantier"] == deduit(ref, c) for c, _ in faits)
    assert figer_statuts(ref) == []                                                              # idempotent
    # l'admin change un statut : il n'est plus écrasé par le taux
    code = faits[0][0]
    ref.update_activity(code, "admin", {"Statut_Chantier": "Suspendu"})
    assert figer_statuts(ref) == [] and ref.get(code)["Statut_Chantier"] == "Suspendu"


def test_statuts_a_verifier_signale_les_ecarts_avec_le_terrain(ref):
    assert statuts_a_verifier(ref) == []                                                         # aucun statut saisi : rien à comparer
    figer_statuts(ref)
    assert statuts_a_verifier(ref) == []                                                         # statuts fidèles au taux
    suivis = latest_valid(ref)
    en_cours = next(c for c, s in suivis.items() if 0 < (s["Taux_Execution_Physique"] or 0) < 100)
    ref.update_activity(en_cours, "admin", {"Statut_Chantier": "Non démarré"})
    ref.update_activity(en_cours, "admin", {"Statut_Chantier": "Achevé"})
    lignes = statuts_a_verifier(ref)
    assert lignes[0].startswith("Statuts à vérifier (1)") and f"{en_cours} : statut « Achevé » mais taux" in lignes[1]
    ref.update_activity(en_cours, "admin", {"Statut_Chantier": "Suspendu"})                      # « Suspendu » : décision de l'admin, pas d'alerte
    assert statuts_a_verifier(ref) == []


def test_le_rapport_du_lundi_mentionne_les_statuts_a_verifier(ref):
    figer_statuts(ref)
    code = next(c for c, s in latest_valid(ref).items() if 0 < (s["Taux_Execution_Physique"] or 0) < 100)
    ref.update_activity(code, "admin", {"Statut_Chantier": "Achevé"})
    res = run(ref, [], date(2026, 11, 16), send_mail=False)
    assert res.ok and any("Statuts à vérifier" in a for a in res.avertissements)
    assert any(code in ligne for ligne in res.lines())


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")
def test_dashboard_et_filtres_suivent_les_statuts_de_l_admin(ref, tmp_path):
    from test_kpi_et_ong import calcule
    figer_statuts(ref)
    acts = [a["CODE_ACTIVITE"] for a in ref.activities()]
    n = len(acts)
    for c in acts[:3]:
        ref.update_activity(c, "admin", {"Statut_Chantier": "Achevé"})                          # même si le taux n'est pas à 100 %
    ref.update_activity(acts[3], "admin", {"Statut_Chantier": "Suspendu"})
    ref.update_activity(acts[4], "admin", {"Statut_Chantier": "En cours"})
    attendu = {}
    for c in acts:
        attendu[ref.get(c)["Statut_Chantier"]] = attendu.get(ref.get(c)["Statut_Chantier"], 0) + 1
    for statut in ("Achevé", "Suspendu", "En cours", "Non démarré"):
        wb = calcule(ref, tmp_path, B19=statut)
        assert wb["TABLEAU DE BORD"]["G9"].value == f"{attendu.get(statut, 0)} / {n}", statut
    assert calcule(ref, tmp_path)["TABLEAU DE BORD"]["G9"].value == f"{n} / {n}"
