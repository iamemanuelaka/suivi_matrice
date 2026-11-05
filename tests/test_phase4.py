import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr.baseline import import_baseline, parse_suivis
from pcr.config import load_config
from pcr.ingest import flatten, ingest, set_statut
from pcr.referentiel import Referentiel
from pcr.seed import seed
from pcr.store import Base

SRC = Path(__file__).resolve().parent.parent / "sources"
TODAY = date(2026, 11, 4)


@pytest.fixture
def ref(tmp_path):
    r = Referentiel(Base(tmp_path / "b.xlsx"), load_config())
    r.create_activity("BS", "BAF", Entreprise="X", Date_Debut="17/08/2026", Delai_Mois=4)
    r.create_activity("BS", "BAF")
    return r


def sub(i, code="BSBAF001", jour="2026-10-22", taux="40", obs="Fondation en cours", t="2026-10-22T10:00:00", **kw):
    d = {"_id": i, "_submission_time": t, "identification/nom_agent": "Kouadio A", "identification/date_actualisation": jour,
         "identification/code_activite": code, "avancement/observations_techniques": obs, "avancement/taux_execution": taux,
         "remarques/difficultes": "RAS"}
    d.update(kw)
    return d


def rows(ref):
    return ref.base.rows("Suivi_Activites")


def test_champs_de_groupes_aplatis():
    f = flatten(sub(1))
    assert f["code_activite"] == "BSBAF001" and f["taux_execution"] == "40" and f["_id"] == 1


def test_integration_nominale(ref):
    rep = ingest(ref, [sub(100)], TODAY)
    r = rows(ref)[0]
    assert rep.count("VALIDE") == 1 and (r["ID_SOUMISSION_KOBO"], r["Taux_Execution_Physique"], r["Statut_Traitement"], r["Source"]) == (100, 40, "VALIDE", "KOBO")
    assert r["Observations_Techniques"] == "Fondation en cours" and r["Difficultes"] == "RAS" and r["Commentaires"] is None


def test_idempotence_aucune_double_integration(ref):
    ingest(ref, [sub(100)], TODAY)
    rep = ingest(ref, [sub(100), sub(101, taux="50", t="2026-10-23T10:00:00")], TODAY)
    assert rep.ignorees == 1 and len(rep.nouvelles) == 1 and len(rows(ref)) == 2
    rep = ingest(ref, [sub(100), sub(101)], TODAY)
    assert rep.ignorees == 2 and not rep.nouvelles and len(rows(ref)) == 2


def test_nouvelle_soumission_n_ecrase_pas_l_historique(ref):
    ingest(ref, [sub(100, taux="40")], TODAY)
    ingest(ref, [sub(101, taux="55", jour="2026-10-29", t="2026-10-29T09:00:00")], TODAY)
    assert [r["Taux_Execution_Physique"] for r in rows(ref)] == [40, 55]


@pytest.mark.parametrize("modif,motif", [
    ({"identification/code_activite": "BSBAF999"}, "code d'activité inconnu"),
    ({"avancement/taux_execution": "150"}, "taux d'exécution invalide"),
    ({"avancement/taux_execution": "abc"}, "taux d'exécution invalide"),
    ({"avancement/observations_techniques": "  "}, "observations techniques vides"),
    ({"identification/date_actualisation": "2026-12-25"}, "dans le futur"),
    ({"identification/date_actualisation": "n'importe quoi"}, "date d'actualisation invalide"),
    ({"_validation_status": {"uid": "validation_status_not_approved"}}, "refusée dans Kobo"),
])
def test_soumissions_invalides_rejetees_mais_conservees(ref, modif, motif):
    rep = ingest(ref, [sub(100, **modif)], TODAY)
    r = rows(ref)[0]
    assert rep.count("REJETE") == 1 and r["Statut_Traitement"] == "REJETE" and motif in r["Motif_Rejet"]
    assert ingest(ref, [sub(100, **modif)], TODAY).ignorees == 1        # elle n'est pas retraitée à chaque exécution


def test_baisse_de_taux_a_verifier_puis_validation_par_l_administrateur(ref):
    ingest(ref, [sub(100, taux="60")], TODAY)
    rep = ingest(ref, [sub(101, taux="40", jour="2026-10-29", t="2026-10-29T09:00:00")], TODAY)
    assert rep.count("A_VERIFIER") == 1 and "taux en baisse" in rows(ref)[1]["Motif_Rejet"]
    set_statut(ref, "101", "VALIDE")
    assert rows(ref)[1]["Statut_Traitement"] == "VALIDE"
    with pytest.raises(KeyError):
        set_statut(ref, "999", "VALIDE")


def test_ordre_d_arrivee_different_de_l_ordre_des_visites(ref):
    rep = ingest(ref, [sub(102, taux="70", jour="2026-10-26", t="2026-10-27T08:00:00"), sub(101, taux="55", jour="2026-10-12", t="2026-10-28T08:00:00")], TODAY)
    assert rep.count("A_VERIFIER") == 0 and rep.count("VALIDE") == 2         # visite ancienne envoyée tardivement : pas de fausse alerte


def test_total_facture_reporte_depuis_le_dernier_suivi(ref):
    ref.base.append("Suivi_Activites", {"ID_SOUMISSION_KOBO": "INIT-BSBAF001", "CODE_ACTIVITE": "BSBAF001", "Date_Actualisation": date(2026, 10, 1),
                                        "Taux_Execution_Physique": 10, "Statut_Traitement": "VALIDE", "Total_Facture": 12})
    ingest(ref, [sub(100, taux="40")], TODAY)
    assert rows(ref)[1]["Total_Facture"] == 12


def test_journal_d_execution(ref):
    ingest(ref, [sub(100), sub(101, code="NOPE")], TODAY, declencheur="test")
    j = ref.base.rows("Journal_Executions")[0]
    assert (j["Recuperees"], j["Traitees"], j["Ignorees"], j["Echouees"], j["Declencheur"]) == (2, 2, 0, 1, "test")


def test_soumission_sans_identifiant_signalee(ref):
    rep = ingest(ref, [{"identification/code_activite": "BSBAF001"}], TODAY)
    assert rep.count("REJETE") == 1 and not rows(ref)


# ---------- suivi de référence importé des Word ----------
def test_baseline_importee_depuis_les_word(tmp_path):
    r = Referentiel(Base(tmp_path / "b.xlsx"), load_config())
    seed(r, [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")])
    warns = import_baseline(r, [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")], date(2026, 10, 25))
    assert len(rows(r)) == 42 and any("répète" in w for w in warns)
    s = {x["CODE_ACTIVITE"]: x for x in rows(r)}
    assert s["BSHAM001"]["Taux_Execution_Physique"] == 60 and s["BSHAM001"]["Date_Actualisation"] == date(2026, 8, 27)
    assert s["MRGON001"]["Taux_Execution_Physique"] == 45 and s["MRBAF001"]["Taux_Execution_Physique"] is None
    assert s["BSBAF001"]["Commentaires"] == "Personnel de la mission de contrôle mobilisé." and s["BSBAF001"]["Total_Facture"] == 0
    n = len(rows(r))
    import_baseline(r, [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")], date(2026, 10, 25))
    assert len(rows(r)) == n                                              # idempotent
