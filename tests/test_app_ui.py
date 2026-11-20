"""Interface du logiciel : parcours d'un utilisateur simulé (pas de navigateur), sur une copie de la base."""
import asyncio
import os
import shutil
import sys
from pathlib import Path

import pytest

pytest.importorskip("nicegui", exc_type=ImportError)
from nicegui import ui  # noqa: E402
from nicegui.testing.user_simulation import user_simulation  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
MAIN = Path(__file__).parent / "app_ui" / "application.py"


@pytest.fixture
def ws(tmp_path, monkeypatch):
    monkeypatch.setenv("PCR_APP_HOME", str(tmp_path / "home"))
    w = tmp_path / "ws"
    (w / "data").mkdir(parents=True)
    shutil.copy(Path(__file__).parent / "app_ui" / "base_reference.xlsx", w / "data" / "base_pcr.xlsx")
    monkeypatch.setenv("PCR_TEST_WS", str(w))
    cwd = os.getcwd()
    yield w
    os.chdir(cwd)


def alertes_du_jour(w):
    from pcr.config import load_config
    from pcr.indicateurs import calculer
    from pcr.referentiel import Referentiel
    from pcr.store import Base
    return len(calculer(Referentiel(Base(w / "data" / "base_pcr.xlsx"), load_config())).alertes)


async def attendre(condition, delai=10.0):
    """La table est rechargée en arrière-plan après la notification : on attend qu'elle soit à jour."""
    for _ in range(int(delai / 0.05)):
        if condition():
            return
        await asyncio.sleep(0.05)
    raise AssertionError("condition non remplie à temps")


def scenario(coro):
    async def go():
        async with user_simulation(main_file=MAIN) as user:
            await coro(user)
    asyncio.run(go())


def test_tableau_de_bord_et_menu(ws):
    async def s(user):
        await user.open("/")
        await user.should_see("Activités sélectionnées / total")
        await user.should_see("42 / 42")
        await user.should_see(f"{alertes_du_jour(ws)} / 42")             # activités en alerte (dépend de la date du jour), comme le classeur
        await user.should_see("Consommation des délais par région", retries=50)
        for page in ("Lancer le suivi", "Suivis à valider", "Activités", "Paramètres"):
            await user.should_see(page)
    scenario(s)


def test_filtre_region_comme_dans_l_excel(ws):
    async def s(user):
        await user.open("/")
        await user.should_see("42 / 42")
        region = next(iter(user.find(marker="filtre-region").elements))
        assert "Tchologo" in region.options.values()
        region.value = "Tchologo"
        await user.should_see("4 / 42", retries=100)
        await user.should_see("Consommation des délais par entreprise")
        entreprises = next(iter(user.find(marker="filtre-entreprise").elements))
        assert set(entreprises.options) - {""} == {"SABARI KAGNI", "ECF", "DMS", "LIKA"}       # listes liées
        user.find("Réinitialiser").click()
        await user.should_see("42 / 42", retries=100)
    scenario(s)


def test_creer_une_activite_depuis_le_formulaire(ws):
    async def s(user):
        await user.open("/activites")
        table = next(iter(user.find(ui.table).elements))
        assert any(r["code"] == "BSBAF001" for r in table.rows)
        user.find(marker="nouvelle").click()
        await user.should_see("Nouvelle activité")
        user.find(marker="f-cat").elements.pop().value = "BS"
        user.find(marker="f-reg").elements.pop().value = "POR"
        user.find(marker="f-Departement").type("Korhogo")
        user.find(marker="f-Montant_Marche").type("12 500 000")
        user.find(marker="f-Delai_Mois").type("4")
        user.find(marker="enregistrer").click()
        await user.should_see("Activité créée : BSPOR003")
        await attendre(lambda: any(r["code"] == "BSPOR003" and r["dep"] == "Korhogo" for r in table.rows))
    scenario(s)


def test_volet_chantier_saisi_par_l_administrateur(ws):
    """Statut, PGES, ONG et finances : saisis dans l'application (les agents ne les renseignent pas dans Kobo)."""
    async def s(user):
        await user.open("/activites")
        table = next(iter(user.find(ui.table).elements))
        for listener in table._event_listeners.values():
            if listener.type == "rowClick":
                with table.parent_slot:
                    listener.handler(type("Ev", (), {"args": [{}, {"code": "BSBAF001"}, 0], "sender": table, "client": user.client})())
        await user.should_see("BSBAF001 – Koro")
        next(iter(user.find(marker="f-Statut_Chantier").elements)).value = "Suspendu"
        next(iter(user.find(marker="f-PGES").elements)).value = "Validé"
        next(iter(user.find(marker="f-ONG_Mobilisees").elements)).value = "ASAPSU"
        next(iter(user.find(marker="f-Montant_Decaisse").elements)).value = "11 835 624"
        await user.should_see("Taux d'exécution financière : 10,0 %")
        user.find(marker="enregistrer").click()
        await user.should_see("modification(s) enregistrée(s)")
    scenario(s)
    from pcr.config import load_config
    from pcr.referentiel import Referentiel
    from pcr.store import Base
    ref = Referentiel(Base(ws / "data" / "base_pcr.xlsx"), load_config(ws / "config" / "config.yaml"))
    a = ref.get("BSBAF001")
    assert (a["Statut_Chantier"], a["PGES"], a["ONG_Mobilisees"], a["Montant_Decaisse"]) == ("Suspendu", "Validé", "ASAPSU", 11_835_624)
    modifs = [h["Champ_Modifie"] for h in ref.history("BSBAF001") if h["Utilisateur"] == ref.cfg.admin and h["Champ_Modifie"] != "(création)"]
    assert sorted(modifs) == ["Montant_Decaisse", "ONG_Mobilisees", "PGES", "Statut_Chantier"], modifs   # rien d'autre n'a bougé


def test_suivis_a_verifier_et_rejet_sans_motif_refuse(ws):
    async def s(user):
        await user.open("/suivis")
        await user.should_see("À vérifier")
        table = next(iter(user.find(ui.table).elements))
        assert len(table.rows) == 2                                      # 2 fiches à vérifier dans la base de référence
        sid = table.rows[0]["id"]
        for listener in table._event_listeners.values():                 # simule un clic sur la première ligne
            if listener.type == "rowClick":
                with table.parent_slot:
                    listener.handler(type("Ev", (), {"args": [{}, table.rows[0], 0], "sender": table, "client": user.client})())
        await user.should_see("Fiche Kobo n° :")
        user.find(marker="rejeter").click()
        await user.should_see("Indiquez le motif du rejet.")
        user.find("Motif du rejet (obligatoire pour rejeter)").type("Photo floue, refaire la visite")
        user.find(marker="rejeter").click()
        await user.should_see("Suivi rejeté.")
        await attendre(lambda: len(table.rows) == 1 and table.rows[0]["id"] != sid)
    scenario(s)


def test_parametres_enregistres_dans_le_profil(ws):
    async def s(user):
        await user.open("/parametres")
        await user.should_see("KoboToolbox")
        user.find("Votre nom").type("Emmanuel AKA")
        user.find(marker="enregistrer-parametres").click()
        await user.should_see("Paramètres enregistrés.")
    scenario(s)
    import json
    data = json.loads((Path(os.environ["PCR_APP_HOME"]) / "PCR-CI" / "parametres.json").read_text(encoding="utf-8"))
    assert data["administrateur"] == "Emmanuel AKA"


def test_generer_les_documents_sans_kobo(ws):
    async def s(user):
        await user.open("/suivi")
        await user.should_see("Aucun document pour l'instant.")
        user.find(marker="lancer").click()
        await user.should_see("2 document(s) généré(s)", retries=200)          # matrices Word
        await user.should_see("MATRICE_DE_SUIVI_DES_BLOCS_SANITAIRES")
        user.find(marker="lancer-tableau").click()
        await user.should_see("Tableau de bord généré", retries=200)
        await user.should_see("TABLEAU_DE_BORD_PCR-CI_du_")
    scenario(s)
