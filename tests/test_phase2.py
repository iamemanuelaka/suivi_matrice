import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr.admin_file import CLEAR, apply_admin_file, export_admin_file
from pcr.calculs import add_months, fin_previsionnelle, indicateurs_delai
from pcr.config import load_config
from pcr.referentiel import Referentiel
from pcr.store import Base


@pytest.fixture
def ref(tmp_path):
    return Referentiel(Base(tmp_path / "base.xlsx"), load_config())


# ---------- codes ----------
def test_format_et_sequence_par_categorie_et_region(ref):
    assert ref.create_activity("BS", "BAF") == "BSBAF001"
    assert ref.create_activity("BS", "BAF") == "BSBAF002"
    assert ref.create_activity("BS", "BER") == "BSBER001"      # la séquence repart à 001 pour une autre région
    assert ref.create_activity("MR", "BAF") == "MRBAF001"      # ... et pour une autre catégorie
    assert ref.create_activity("MR", "POR") == "MRPOR001"


def test_noms_acceptes_a_la_place_des_codes(ref):
    assert ref.create_activity("Blocs sanitaires", "Bounkani") == "BSBUN001"
    assert ref.create_activity("marchés ruraux", "BÉRÉ") == "MRBER001"


def test_categorie_ou_region_inconnue_refusee(ref):
    with pytest.raises(ValueError):
        ref.create_activity("XX", "BAF")
    with pytest.raises(ValueError):
        ref.create_activity("BS", "ZZZ")


def test_code_jamais_reutilise_apres_archivage(ref):
    a, b = ref.create_activity("BS", "BAF"), ref.create_activity("BS", "BAF")
    ref.archive(b)
    assert ref.create_activity("BS", "BAF") == "BSBAF003"
    assert ref.get(b)["Statut"] == "Archivé"


def test_compteur_prime_meme_si_la_ligne_disparait(ref):
    ref.create_activity("BS", "BAF"), ref.create_activity("BS", "BAF")
    ws = ref.base.wb["Activites"]
    ws.delete_rows(3)                                           # suppression manuelle (accident) du 2e code
    assert ref.create_activity("BS", "BAF") == "BSBAF003"       # le compteur empêche la réattribution de 002


def test_persistance_apres_rechargement(ref, tmp_path):
    ref.create_activity("BS", "BAF")
    ref.base.save()
    ref2 = Referentiel(Base(tmp_path / "base.xlsx"), load_config())
    assert ref2.create_activity("BS", "BAF") == "BSBAF002"


# ---------- informations incomplètes ----------
def test_creation_avec_informations_incompletes(ref):
    code = ref.create_activity("MR", "BAF")                     # aucune information fixe
    a = ref.get(code)
    assert a["Entreprise"] is None and a["Montant_Marche"] is None and a["Date_Fin_Previsionnelle"] is None
    assert a["Statut"] == "Actif"


def test_fin_previsionnelle_calculee_seulement_si_donnees_completes(ref):
    c1 = ref.create_activity("BS", "BAF", Date_Debut="17/08/2026")                  # délai manquant
    c2 = ref.create_activity("BS", "BAF", Date_Debut="17/08/2026", Delai_Mois=4)
    assert ref.get(c1)["Date_Fin_Previsionnelle"] is None
    assert ref.get(c2)["Date_Fin_Previsionnelle"] == date(2026, 12, 17)
    ref.update_activity(c1, changes={"Delai_Mois": 4})                              # complétée plus tard
    assert ref.get(c1)["Date_Fin_Previsionnelle"] == date(2026, 12, 17)


# ---------- modifications tracées ----------
def test_modification_historisee(ref):
    c = ref.create_activity("BS", "BAF", Entreprise="ANCIENNE SARL", Montant_Marche="118 356 245 F CFA")
    done = ref.update_activity(c, "Admin Test", {"Entreprise": "NOUVELLE SARL"})
    assert done == [("Entreprise", "ANCIENNE SARL", "NOUVELLE SARL")]
    h = [x for x in ref.history(c) if x["Champ_Modifie"] == "Entreprise"]
    assert len(h) == 1
    assert (h[0]["Ancienne_Valeur"], h[0]["Nouvelle_Valeur"], h[0]["Utilisateur"]) == ("ANCIENNE SARL", "NOUVELLE SARL", "Admin Test")
    assert h[0]["Date_Modification"] is not None
    assert ref.get(c)["Montant_Marche"] == 118356245


def test_valeur_identique_ne_cree_pas_d_historique(ref):
    c = ref.create_activity("BS", "BAF", Entreprise="X")
    n = len(ref.history(c))
    assert ref.update_activity(c, changes={"Entreprise": "X"}) == []
    assert len(ref.history(c)) == n


def test_vide_n_efface_jamais(ref):
    c = ref.create_activity("BS", "BAF", Entreprise="X", Montant_Marche=100)
    assert ref.update_activity(c, changes={"Entreprise": "", "Montant_Marche": None}) == []
    a = ref.get(c)
    assert a["Entreprise"] == "X" and a["Montant_Marche"] == 100


def test_effacement_uniquement_explicite_et_trace(ref):
    c = ref.create_activity("BS", "BAF", Entreprise="X")
    ref.update_activity(c, clear=["Entreprise"])
    assert ref.get(c)["Entreprise"] is None
    assert [x for x in ref.history(c) if x["Champ_Modifie"] == "Entreprise"][-1]["Ancienne_Valeur"] == "X"


def test_categorie_region_non_modifiables(ref):
    c = ref.create_activity("BS", "BAF")
    for champ in ("Categorie", "Region", "CODE_ACTIVITE"):
        with pytest.raises(ValueError):
            ref.update_activity(c, changes={champ: "autre"})


def test_valeurs_invalides_refusees(ref):
    c = ref.create_activity("BS", "BAF")
    for champ, val in (("Date_OS", "31/02/2026"), ("Delai_Mois", 0), ("Montant_Marche", "abc"), ("Statut", "Foo")):
        with pytest.raises(ValueError):
            ref.update_activity(c, changes={champ: val})


# ---------- vérification d'unicité ----------
def test_verify_conforme_puis_detecte_un_doublon(ref):
    ref.create_activity("BS", "BAF"), ref.create_activity("BS", "BAF")
    assert ref.verify() == []
    ref.base.wb["Activites"].cell(3, 1).value = "BSBAF001"       # corruption : code en double
    assert any("double" in p for p in ref.verify())


# ---------- fichier d'administration ----------
def test_fichier_admin_creation_modification_et_regle_du_vide(ref, tmp_path):
    c = ref.create_activity("BS", "BAF", Entreprise="X", Delai_Mois=4)
    p = export_admin_file(ref, tmp_path / "admin.xlsx")
    wb = load_workbook(p)
    ws = wb["Activites"]
    cols = {c.value: c.column for c in ws[1]}
    ws.cell(2, cols["Entreprise"]).value = None                  # cellule vidée par erreur -> ne doit rien effacer
    ws.cell(2, cols["Montant_Marche"]).value = 5000
    ws.cell(2, cols["Date_Debut"]).value = "17/08/2026"
    ws.cell(3, cols["Categorie"]).value = "MR"                   # nouvelle ligne
    ws.cell(3, cols["Region"]).value = "Poro"
    wb.save(p)
    rep = apply_admin_file(ref, p)
    assert [x[1] for x in rep.created] == ["MRPOR001"] and not rep.errors
    a = ref.get(c)
    assert a["Entreprise"] == "X" and a["Montant_Marche"] == 5000 and a["Date_Fin_Previsionnelle"] == date(2026, 12, 17)


def test_fichier_admin_vider_et_erreurs_par_ligne(ref, tmp_path):
    c = ref.create_activity("BS", "BAF", Entreprise="X")
    ref.create_activity("BS", "BAF")
    p = export_admin_file(ref, tmp_path / "admin.xlsx")
    wb = load_workbook(p)
    ws = wb["Activites"]
    cols = {c.value: c.column for c in ws[1]}
    ws.cell(2, cols["Entreprise"]).value = CLEAR
    ws.cell(3, cols["Date_OS"]).value = "pas une date"
    ws.cell(4, cols["Categorie"]).value = "BS"                   # ligne sans région
    wb.save(p)
    rep = apply_admin_file(ref, p)
    assert ref.get(c)["Entreprise"] is None
    assert len(rep.errors) == 2 and {r for r, _ in rep.errors} == {3, 4}


def test_fichier_admin_changement_de_region_refuse(ref, tmp_path):
    ref.create_activity("BS", "BAF")
    p = export_admin_file(ref, tmp_path / "admin.xlsx")
    wb = load_workbook(p)
    wb["Activites"].cell(2, 3).value = "BER"
    wb.save(p)
    rep = apply_admin_file(ref, p)
    assert rep.errors and ref.get("BSBAF001")["Region"] == "BAF"


# ---------- calculs ----------
def test_ajout_de_mois_calendaires_comme_edate():
    assert add_months(date(2026, 8, 17), 4) == date(2026, 12, 17)
    assert add_months(date(2026, 10, 31), 4) == date(2027, 2, 28)       # jour inexistant -> dernier jour du mois
    assert add_months(date(2026, 11, 30), 3) == date(2027, 2, 28)
    assert add_months(date(2027, 1, 31), 1) == date(2027, 2, 28)


def test_indicateurs_coherents_avec_la_matrice_word():
    # Lot de Koro : début 17/08/2026, 4 mois -> le Word affiche 54,09 % de délai consommé (au 22/10/2026)
    i = indicateurs_delai(date(2026, 8, 17), 4, date(2026, 10, 22))
    assert i["fin_previsionnelle"] == date(2026, 12, 17) and i["duree_jours"] == 122 and i["ecoule_jours"] == 66
    assert abs(i["consommation_pct"] - 54.09) < 0.02 and i["depassement_jours"] == 0


def test_depassement_et_avant_demarrage():
    assert indicateurs_delai(date(2026, 8, 17), 4, date(2027, 1, 16))["depassement_jours"] == 30
    avant = indicateurs_delai(date(2026, 8, 17), 4, date(2026, 8, 1))
    assert avant["ecoule_jours"] == 0 and avant["consommation_pct"] == 0


@pytest.mark.parametrize("debut,delai", [(None, 4), (date(2026, 8, 17), None), (date(2026, 8, 17), 0), (date(2026, 8, 17), -2), (date(2026, 8, 17), 2.5)])
def test_pas_de_calcul_artificiel_si_donnees_manquantes_ou_incoherentes(debut, delai):
    assert fin_previsionnelle(debut, delai) is None and indicateurs_delai(debut, delai, date(2026, 10, 1)) is None
