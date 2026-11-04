import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr.config import load_config
from pcr.referentiel import Referentiel
from pcr.seed import seed
from pcr.store import Base

SRC = Path(__file__).resolve().parent.parent / "sources"


def charger(tmp_path):
    ref = Referentiel(Base(tmp_path / "b.xlsx"), load_config())
    rows = seed(ref, [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")])
    return ref, {r["code"]: r for r in rows}


def test_effectifs_et_integrite(tmp_path):
    ref, rows = charger(tmp_path)
    assert sum(c.startswith("BS") for c in rows) == 26 and sum(c.startswith("MR") for c in rows) == 16
    assert ref.verify() == []


def test_codes_dans_l_ordre_du_word(tmp_path):
    ref, _ = charger(tmp_path)
    assert ref.get("BSBAF001")["Departement"] == "Koro"
    assert ref.get("BSPOR001")["Departement"] == "Dikodougou" and ref.get("BSPOR001")["Entreprise"] == "EGTP"
    assert ref.get("MRPOR003")["Entreprise"] == "ETS MAGIC SERVICES"


def test_donnees_extraites(tmp_path):
    ref, _ = charger(tmp_path)
    a = ref.get("BSBAF001")
    assert a["Montant_Marche"] == 118356245 and a["Date_OS"] == date(2026, 8, 17) and a["Date_Debut"] == date(2026, 8, 17)
    assert a["Delai_Mois"] == 4 and a["Date_Fin_Previsionnelle"] == date(2026, 12, 17)      # « 17/012/2026 » du Word corrigée
    m = ref.get("MRBER001")
    assert m["Localite"] == "Dianra et Kpesso".replace(" et ", ", ") and m["Date_Debut"] == date(2026, 9, 16)


def test_activites_incompletes_et_dates_incoherentes(tmp_path):
    ref, _ = charger(tmp_path)
    m = ref.get("MRBAF001")                                        # marché pas encore attribué
    assert m["Entreprise"] is None and m["Montant_Marche"] is None and m["Date_Fin_Previsionnelle"] is None
    w = ref.get("BSWOR003")                                        # O.S. daté du 24/12/2026 = fin prévue : incohérent
    assert w["Date_OS"] == date(2026, 12, 24) and w["Date_Debut"] is None and w["Date_Fin_Previsionnelle"] is None


def test_chargement_non_rejouable(tmp_path):
    ref, _ = charger(tmp_path)
    try:
        seed(ref, [("BS", SRC / "matrice_BS_texte.md")])
        assert False
    except RuntimeError:
        pass
