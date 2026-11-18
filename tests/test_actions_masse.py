"""PGES en masse : « tous les chantiers En cours ont un PGES validé », avec simulation, historique et reprise des exceptions."""
import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref  # noqa: F401,E402
from pcr.actions_masse import pges_en_masse, statut_affiche
from pcr.word_gen import latest_valid


def par_statut(ref, statut):
    suivis = latest_valid(ref)
    return [a["CODE_ACTIVITE"] for a in ref.activities() if a["Statut"] == "Actif" and statut_affiche(a, suivis.get(a["CODE_ACTIVITE"])) == statut]


def test_simulation_n_ecrit_rien(ref):
    avant = {a["CODE_ACTIVITE"]: a["PGES"] for a in ref.activities()}
    n_hist = len(ref.history())
    rap = pges_en_masse(ref, "En cours", "Validé", dry_run=True)
    assert len(rap["modifiees"]) == len(par_statut(ref, "En cours")) > 0
    assert {a["CODE_ACTIVITE"]: a["PGES"] for a in ref.activities()} == avant and len(ref.history()) == n_hist


def test_tous_les_chantiers_en_cours_passent_a_valide_les_autres_sont_intacts(ref):
    ref.update_activity(par_statut(ref, "En cours")[0], "admin", {"PGES": "Non transmis"})        # ancienne valeur : sera remplacée
    non_demarres = par_statut(ref, "Non démarré")
    if non_demarres:
        ref.update_activity(non_demarres[0], "admin", {"PGES": "Non transmis"})
    en_cours = par_statut(ref, "En cours")
    rap = pges_en_masse(ref, "en cours", "validé")                                                   # casse / accents tolérés
    assert sorted(c for c, _ in rap["modifiees"]) == sorted(en_cours)
    assert all(ref.get(c)["PGES"] == "Validé" for c in en_cours)
    if non_demarres:
        assert ref.get(non_demarres[0])["PGES"] == "Non transmis"
    h = [x for x in ref.history(en_cours[0]) if x["Champ_Modifie"] == "PGES"]
    assert h and h[-1]["Nouvelle_Valeur"] == "Validé" and h[-1]["Utilisateur"]
    assert pges_en_masse(ref, "En cours", "Validé")["modifiees"] == []                              # idempotent
    # une exception saisie ensuite par l'administrateur est respectée tant qu'on ne relance pas la commande
    ref.update_activity(en_cours[1], "admin", {"PGES": "En cours de validation"})
    assert ref.get(en_cours[1])["PGES"] == "En cours de validation"


def test_option_seulement_vides_et_archives(ref):
    en_cours = par_statut(ref, "En cours")
    ref.update_activity(en_cours[0], "admin", {"PGES": "Non transmis"})
    rap = pges_en_masse(ref, "En cours", "Validé", seulement_vides=True)
    assert rap["gardees"] == [(en_cours[0], "Non transmis")] and ref.get(en_cours[0])["PGES"] == "Non transmis"
    ref.update_activity(en_cours[2], "admin", {"Statut": "Archivé"})
    ref.update_activity(en_cours[2], "admin", {}, clear=["PGES"])
    pges_en_masse(ref, "En cours", "Validé")
    assert ref.get(en_cours[2])["PGES"] is None                                                      # activité archivée : non concernée


def test_valeurs_invalides_refusees(ref):
    with pytest.raises(ValueError, match="PGES invalide"):
        pges_en_masse(ref, "En cours", "Peut-être")
    with pytest.raises(ValueError, match="statut du chantier invalide"):
        pges_en_masse(ref, "Abandonné", "Validé")


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")
def test_le_tableau_de_bord_affiche_tous_les_pges_valides_des_chantiers_en_cours(ref, tmp_path):
    from test_kpi_et_ong import calcule
    pges_en_masse(ref, "En cours", "Validé")
    n = len(par_statut(ref, "En cours"))
    wb = calcule(ref, tmp_path, B19="En cours")                      # filtre statut « En cours »
    assert wb["TABLEAU DE BORD"]["Q9"].value == f"{n} Validés"
    assert wb["TABLEAU DE BORD"]["G9"].value.startswith(f"{n} /")
