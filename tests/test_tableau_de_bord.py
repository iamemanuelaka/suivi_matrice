"""Tableau de bord de l'application : mêmes chiffres que le classeur Excel (valeurs relevées dans le classeur recalculé)."""
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pcr.config import load_config  # noqa: E402
from pcr.referentiel import Referentiel  # noqa: E402
from pcr.store import Base  # noqa: E402
from pcr_app.tableau_de_bord import Filtres, calculer  # noqa: E402

BASE = Path(__file__).parent / "app_ui" / "base_reference.xlsx"
AU = date(2026, 11, 19)                     # date du classeur recalculé (fonction AUJOURDHUI)


@pytest.fixture(scope="module")
def ref(tmp_path_factory):
    d = tmp_path_factory.mktemp("tb")
    shutil.copy(BASE, d / "base.xlsx")
    return Referentiel(Base(d / "base.xlsx"), load_config())


def test_sans_filtre_identique_au_classeur(ref):
    t = calculer(ref, au=AU)
    assert (len(t.selection), t.total) == (42, 42)                                           # G9 « 42 / 42 »
    assert round(t.taux_moyen, 6) == 27.961538                                               # L9
    assert t.pges_valides == 27 and (t.ong_nommees, t.ong_sans_nom) == (11, 1)               # Q9, Q11, Q12
    assert round(t.montant_total / 1e9, 6) == 5.906245 and round(t.decaisse_total / 1e6, 1) == 30.0   # V9, V11
    assert round(t.taux_financier, 1) == 0.5                                                 # V12
    assert len(t.alertes) == 9 and [a.numero for a in t.alertes[:5]] == ["P-010", "P-031", "P-017", "P-002", "P-005"]  # AA9, AF16:AF20
    assert t.par == "région" and [(g["libelle"], g["n"], round(g["conso"], 4), round(g["execution"], 4)) for g in t.groupes[:3]] == [
        ("Bafing", 2, 77.0492, 31.7002), ("Hambol", 2, 74.7934, 38.545), ("Bagoué", 1, 71.3115, 25.0)]   # CALCULS!EL:EN
    assert sum(1 for l in t.selection if l.evaluable) == 21
    assert t.exclusions["sans date de début"] == 6 and t.exclusions["sans taux d'exécution"] == 15   # G39


def test_region_choisie_graphique_par_entreprise_et_listes_liees(ref):
    t = calculer(ref, Filtres(region="Tchologo"), au=AU)
    assert t.par == "entreprise" and len(t.selection) == 4
    assert [(g["libelle"], round(g["conso"], 1), round(g["execution"], 1)) for g in t.groupes if g["conso"] is not None] == [
        ("SABARI KAGNI", 68.9, 35.0), ("ECF", 58.3, 10.0), ("DMS", 55.0, 15.0)]
    assert set(t.options["entreprise"]) == {"SABARI KAGNI", "ECF", "DMS", "LIKA"}             # liste limitée à la région


def test_filtre_sans_activite_ignore_comme_dans_l_excel(ref):
    t = calculer(ref, Filtres(region="Bafing", entreprise="DMS"), au=AU)
    assert t.ignores == ["Entreprise"] and t.effectifs.entreprise == "" and len(t.selection) == 3    # 3 activités au Bafing


def test_statut_et_nature(ref):
    t = calculer(ref, Filtres(nature="Blocs sanitaires", statut="En cours"), au=AU)
    assert all(l.nature == "Blocs sanitaires" and l.statut == "En cours" for l in t.selection)
    assert "Suspendu" in calculer(ref, au=AU).options["statut"]


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent : comparaison directe avec le classeur impossible")
@pytest.mark.parametrize("cellules, filtres", [
    ({}, Filtres()),
    ({"B10": "Marchés ruraux"}, Filtres(nature="Marchés ruraux")),
    ({"B13": "Tchologo"}, Filtres(region="Tchologo")),                                    # graphique par entreprise
    ({"B10": "Blocs sanitaires", "B19": "En cours"}, Filtres(nature="Blocs sanitaires", statut="En cours")),
])
def test_le_classeur_recalcule_donne_les_memes_chiffres_que_l_application(ref, tmp_path, cellules, filtres):
    """Génère le tableau de bord, choisit les filtres, le fait recalculer par LibreOffice et compare à l'application."""
    import openpyxl
    from pcr.tableau_de_bord_xlsx import generer
    generer(ref, tmp_path / "tb.xlsx", AU)
    wb = openpyxl.load_workbook(tmp_path / "tb.xlsx")
    for c, v in cellules.items():
        wb["TABLEAU DE BORD"][c] = v
    wb.save(tmp_path / "tb.xlsx")
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(tmp_path / "r"),
                    str(tmp_path / "tb.xlsx")], capture_output=True, timeout=240)
    w = openpyxl.load_workbook(tmp_path / "r" / "tb.xlsx", data_only=True)
    tb, do = w["TABLEAU DE BORD"], w["DONNÉES"]
    t = calculer(ref, filtres, au=AU)
    n = len(t.selection)
    assert tb["D8"].value == f"{n} / {t.total}" and tb["P8"].value == f"{len(t.alertes)} / {n}"
    assert abs(tb["G8"].value - t.taux_moyen) < 1e-9 and tb["J8"].value == f"{t.pges_valides} Validés"
    assert tb["J10"].value == f"ONG mobilisées : {t.ong_nommees}" and abs(tb["M8"].value - t.montant_total / 1e9) < 1e-9
    assert [tb[f"S{r}"].value or "" for r in range(14, 19)] == ([a.numero for a in t.alertes[:5]] + [""] * 5)[:5]
    assert tb["D13"].value.startswith(f"Consommation des délais par {t.par}")
    e = [c.value for c in do[1]]
    gl = e.index("Graphique : libellé") + 1
    xl = [(do.cell(r, gl).value, do.cell(r, gl + 1).value, do.cell(r, gl + 2).value) for r in range(2, 26) if do.cell(r, gl).value]
    app = [(g["conso"] if g["conso"] is not None else -1, g["execution"] if g["execution"] is not None else -1) for g in t.groupes]
    assert [(round(a, 6), round(b, 6)) for _, a, b in xl] == [(round(a, 6), round(b, 6)) for a, b in app]
    assert all("(" not in lib for lib, _, _ in xl)                                       # libellés sans « (n) »
    assert [lib.split("  – ")[0].rstrip("…") for lib, _, _ in xl] == [g["libelle"][:23] if len(g["libelle"]) > 24 else g["libelle"] for g in t.groupes]
