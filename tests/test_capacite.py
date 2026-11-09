"""Capacité 500 activités et alimentation du tableau de bord depuis la base (feuille « MATRICE DE SUIVI » = feuille technique masquée)."""
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref, sub, TEMPLATE  # noqa: F401,E402
from pcr import xlsx_repair as X
from pcr.excel_gen import generate_dashboard
from pcr.ingest import ingest
from pcr.pipeline import alertes
from pcr.xlsx_capacity import capacite_modele, etendre_modele
from pcr.xlsx_patch import Package

ANCIEN = None


def test_le_modele_livre_gere_500_activites_et_reste_sain():
    parts = X.lire_parts(TEMPLATE)
    assert capacite_modele(parts) == 500
    assert X.diagnostiquer(parts) == []
    wb = parts["xl/workbook.xml"].decode()
    assert re.search(r'name="MATRICE DE SUIVI"[^>]*state="hidden"', wb)           # feuille technique masquée
    assert "$504" in wb and "$49" not in wb                                       # noms définis étendus


def test_extension_idempotente_et_refusee_si_deja_etendu(tmp_path):
    with pytest.raises(ValueError, match="déjà"):
        etendre_modele(TEMPLATE, tmp_path / "x.xlsx", 500)
    r = etendre_modele(TEMPLATE, tmp_path / "y.xlsx", 600)
    assert capacite_modele(X.lire_parts(tmp_path / "y.xlsx")) == 600 and X.diagnostiquer(X.lire_parts(tmp_path / "y.xlsx")) == []


def test_ajout_d_un_bloc_sanitaire_avec_entreprise_ne_casse_plus_la_generation(ref, tmp_path):
    """Avant : « le nombre d'activités par région a changé » dès qu'un BS recevait une entreprise dans une région existante."""
    ref.create_activity("BS", "FOL", Entreprise="NOUVELLE ENTREPRISE SARL")
    ref.create_activity("BS", "GON", Entreprise="AUTRE SARL")
    res = generate_dashboard(ref, TEMPLATE, tmp_path / "x.xlsx")
    assert res["lignes"] == 44 and X.diagnostiquer(X.lire_parts(tmp_path / "x.xlsx")) == []


def test_chaque_ligne_porte_sa_region_et_les_lignes_inutilisees_restent_vides(ref, tmp_path):
    generate_dashboard(ref, TEMPLATE, tmp_path / "x.xlsx")
    sh = Package(tmp_path / "x.xlsx").sheet("MATRICE DE SUIVI")
    regions = [sh.get(r, 3) for r in range(4, 46)]
    assert all(regions), "une région manque sur au moins une ligne"
    assert sh.get(46, 3) is None and sh.get(46, 2) is None and sh.merged() == []


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")
def test_500_activites_kpi_et_alertes_identiques_au_calcul_python(ref, tmp_path):
    import random
    from openpyxl import load_workbook
    random.seed(3)
    regs = list(ref.cfg.regions)
    subs, sid = [], 70000
    for _ in range(500 - len(ref.activities())):
        code = ref.create_activity(random.choice(["BS", "MR"]), random.choice(regs), Entreprise=f"ENT {random.randint(1, 50)}",
                                   Montant_Marche=random.randint(50, 300) * 10**6, Date_Debut=date(2026, 8, 1 + random.randint(0, 25)),
                                   Delai_Mois=random.choice([3, 4, 6]))
        sid += 1
        subs.append(sub(sid, code=code, taux=random.choice([0, 5, 10, 40, 80, 100]), jour="2026-11-03"))
    ingest(ref, subs, today=date(2026, 11, 6), log=False)
    out = tmp_path / "d.xlsx"
    generate_dashboard(ref, TEMPLATE, out, datetime(2026, 11, 6, 8, 0))
    lo = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(lo), str(out)],
                   capture_output=True, timeout=250)
    wb = load_workbook(lo / "d.xlsx", data_only=True)
    t = wb["TABLEAU DE BORD"]
    assert t["G9"].value.endswith("/ 500")
    assert t["AA9"].value == f"{len(alertes(ref, date.today()))} / 500"
    assert not [c.coordinate for s in wb for row in s.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith("#") and c.value != "#N/A"]
