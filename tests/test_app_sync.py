"""Application : synchronisation GitHub (faux serveur en mémoire) et opérations de base."""
import base64
import hashlib
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from pcr_app.github_sync import GitHubSync, SyncConflict, blob_sha


def sha(data: bytes) -> str:
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


class FakeGitHub:
    """Imite les 3 points d'API utilisés : arbre, blob brut, écriture d'un fichier."""

    def __init__(self, files: dict[str, bytes]):
        self.files, self.headers, self.puts = dict(files), {}, []

    def _resp(self, status, json=None, content=b"", text=""):
        return type("R", (), {"status_code": status, "json": lambda self=None: json, "content": content, "text": text})()

    def get(self, url, timeout=None, params=None, headers=None):
        if "/git/trees/" in url:
            return self._resp(200, {"tree": [{"path": p, "type": "blob", "sha": sha(d)} for p, d in self.files.items()]})
        if "/git/blobs/" in url:
            s = url.rsplit("/", 1)[1]
            return self._resp(200, content=next(d for d in self.files.values() if sha(d) == s))
        return self._resp(200, {"permissions": {"push": True}})

    def put(self, url, json=None, timeout=None):
        path = url.split("/contents/", 1)[1]
        if json.get("sha") != (sha(self.files[path]) if path in self.files else None):
            return self._resp(409, text="sha does not match")
        self.files[path] = base64.b64decode(json["content"])
        self.puts.append(path)
        return self._resp(200, {"content": {"sha": sha(self.files[path])}})


BASE = "data/base_pcr.xlsx"


@pytest.fixture
def gh(tmp_path):
    fake = FakeGitHub({BASE: b"v1", "config/config.yaml": b"cfg", "data/photos/a.jpg": b"img", "sorties/x.docx": b"ignored"})
    return fake, GitHubSync("o/r", "main", "tok", tmp_path, session=fake), tmp_path


def test_premiere_synchro_recupere_base_config_photos_mais_pas_les_sorties(gh):
    fake, g, root = gh
    rep = g.sync()
    assert sorted(rep.telecharges) == ["config/config.yaml", BASE, "data/photos/a.jpg"]
    assert (root / BASE).read_bytes() == b"v1" and not (root / "sorties").exists() and fake.puts == []


def test_modification_locale_envoyee(gh):
    fake, g, root = gh
    g.sync()
    (root / BASE).write_bytes(b"v2-local")
    rep = g.sync()
    assert rep.envoyes == [BASE] and fake.files[BASE] == b"v2-local"
    assert g.sync().text() == "Déjà à jour"


def test_modification_github_recuperee(gh):
    fake, g, root = gh
    g.sync()
    fake.files[BASE] = b"v2-lundi"
    assert g.sync().telecharges == [BASE] and (root / BASE).read_bytes() == b"v2-lundi"


def test_conflit_rien_n_est_ecrase_puis_choix(gh):
    fake, g, root = gh
    g.sync()
    fake.files[BASE] = b"v2-lundi"
    (root / BASE).write_bytes(b"v2-local")
    with pytest.raises(SyncConflict):
        g.sync()
    assert fake.files[BASE] == b"v2-lundi" and (root / BASE).read_bytes() == b"v2-local"
    g.sync(prefer="github")
    assert (root / BASE).read_bytes() == b"v2-lundi"
    assert any(p.name.startswith("base_pcr_avant_synchro_") for p in (root / "data" / "sauvegardes").iterdir())


def test_conflit_resolu_en_gardant_la_version_locale(gh):
    fake, g, root = gh
    g.sync()
    fake.files[BASE] = b"v2-lundi"
    (root / BASE).write_bytes(b"v2-local")
    g.sync(prefer="local")
    assert fake.files[BASE] == b"v2-local"


def test_empreinte_identique_a_git(tmp_path):
    (tmp_path / "f").write_bytes(b"hello\n")
    assert blob_sha(tmp_path / "f") == "ce013625030ba8dba906f756967f9e9ca394464a"      # git hash-object


@pytest.fixture
def services(tmp_path, monkeypatch):
    monkeypatch.setenv("PCR_APP_HOME", str(tmp_path / "home"))
    ws = tmp_path / "ws"
    (ws / "data").mkdir(parents=True)
    shutil.copy(ROOT / "tests" / "app_ui" / "base_reference.xlsx", ws / "data" / "base_pcr.xlsx")
    from pcr_app.services import Services
    from pcr_app.workspace import Settings
    cwd = Path.cwd()
    yield Services(Settings(workspace=str(ws), github_depot=""))
    import os
    os.chdir(cwd)


def test_services_creation_modification_historique_et_rejet(services):
    from pcr_app.services import AppError
    nb = len(services.activites())
    code, _ = services.creer_activite("Blocs sanitaires", "Poro", {"Departement": "Korhogo", "Montant_Marche": "12 500 000", "Delai_Mois": "4"})
    assert len(services.activites()) == nb + 1 and code.startswith("BSPOR")
    changes, _ = services.modifier_activite(code, {"Entreprise": "ETS TEST"})
    assert changes == [("Entreprise", None, "ETS TEST")]
    assert [h["Champ_Modifie"] for h in services.historique(code)] == ["Entreprise", "(création)"]
    with pytest.raises(AppError, match="décaissé"):
        services.modifier_activite(code, {"Montant_Decaisse": "99 000 000"})
    with pytest.raises(AppError, match="motif"):
        services.decider_suivi("123", valider=False, motif="")


def test_services_documents_sans_kobo(services):
    r = services.lancer(avec_kobo=False)
    assert r.ok and len(services.documents()) == 2                      # matrices Word (le tableau de bord est à part)


# --- traitement GitHub Actions lancé depuis l'application ------------------------------------------
class FakeGitHubActions(FakeGitHub):
    """Ajoute les points d'API des workflows : lancement, liste des exécutions, état, étapes.
    À la fin de l'exécution simulée, le « bot » modifie la base et ajoute des documents dans sorties/."""

    def __init__(self, files, fin_ok=True):
        super().__init__(files)
        self.runs, self.dispatches, self.appels_etat, self.fin_ok = [101], [], 0, fin_ok

    def post(self, url, json=None, timeout=None):
        assert url.endswith("/dispatches")
        self.workflow = url.split("/actions/workflows/")[1].split("/")[0]
        self.dispatches.append(json)
        self.runs.append(102)
        return self._resp(204)

    def get(self, url, timeout=None, params=None, headers=None):
        if url.endswith("/runs") and "/actions/workflows/" in url:
            return self._resp(200, {"workflow_runs": [{"id": i} for i in reversed(self.runs)]})
        if url.endswith("/jobs"):
            return self._resp(200, {"jobs": [{"steps": [{"name": "Tests", "status": "completed"},
                                                         {"name": "Cycle complet (Kobo -> base -> Word + Excel -> e-mail)", "status": "in_progress"}]}]})
        if "/actions/runs/" in url:
            self.appels_etat += 1
            if self.appels_etat < 3:
                return self._resp(200, {"status": "in_progress", "conclusion": None, "html_url": "https://github.com/o/r/actions/runs/102"})
            if self.fin_ok and "sorties/MATRICE_DE_SUIVI_DES_BLOCS_SANITAIRES_20112026.docx" not in self.files:
                self.files[BASE] = self.files[BASE] + b""                     # base inchangée ici : seul le contenu de sorties/ compte
                self.files["sorties/MATRICE_DE_SUIVI_DES_BLOCS_SANITAIRES_20112026.docx"] = b"docx-bs"
                self.files["sorties/MATRICE_DE_SUIVI_DES_MARCHES_RURAUX_20112026.docx"] = b"docx-mr"
            return self._resp(200, {"status": "completed", "conclusion": "success" if self.fin_ok else "failure",
                                    "html_url": "https://github.com/o/r/actions/runs/102"})
        return super().get(url, timeout=timeout, params=params, headers=headers)


@pytest.fixture
def services_gh(tmp_path, monkeypatch):
    monkeypatch.setenv("PCR_APP_HOME", str(tmp_path / "home"))
    base = (ROOT / "tests" / "app_ui" / "base_reference.xlsx").read_bytes()
    fake = FakeGitHubActions({BASE: base, "config/config.yaml": (ROOT / "config" / "config.yaml").read_bytes(),
                              "sorties/ANCIEN_19112026.docx": b"ancien"})
    from pcr_app.services import Services
    from pcr_app.workspace import Settings
    cwd = Path.cwd()
    sv = Services(Settings(workspace=str(tmp_path / "ws"), github_depot="o/r", github_jeton="tok"), session=fake)
    yield sv, fake
    import os
    os.chdir(cwd)


def test_lancer_le_workflow_suivre_puis_recuperer_les_documents(services_gh):
    sv, fake = services_gh
    etat = sv.lancer_sur_github(envoyer_mail=True)
    assert fake.dispatches == [{"ref": "main", "inputs": {"send_mail": "true"}}] and etat["run_id"] is None
    st = sv.suivre_github()
    assert st["status"] == "in_progress" and st["etape"].startswith("Cycle complet")
    assert sv.execution_en_cours()["run_id"] == 102                       # l'exécution 101 (antérieure) est ignorée
    while st["status"] != "completed":
        st = sv.suivre_github()
    assert st["conclusion"] == "success"
    res = sv.recuperer_github()
    assert sorted(res["documents"]) == ["MATRICE_DE_SUIVI_DES_BLOCS_SANITAIRES_20112026.docx",
                                        "MATRICE_DE_SUIVI_DES_MARCHES_RURAUX_20112026.docx"]       # pas l'ancien document
    assert (sv.ws / "sorties" / "MATRICE_DE_SUIVI_DES_MARCHES_RURAUX_20112026.docx").read_bytes() == b"docx-mr"
    assert not (sv.ws / "sorties" / "ANCIEN_19112026.docx").exists()
    assert sv.execution_en_cours() is None and sv.suivre_github() is None


def test_un_seul_traitement_a_la_fois_et_reprise_apres_fermeture(services_gh):
    from pcr_app.services import AppError, Services
    sv, fake = services_gh
    sv.lancer_sur_github()
    with pytest.raises(AppError, match="déjà en cours"):
        sv.lancer_sur_github()
    rouvert = Services(sv.settings, session=fake)                        # application fermée puis rouverte
    assert rouvert.execution_en_cours() is not None and rouvert.suivre_github()["status"] == "in_progress"


def test_jeton_obligatoire_pour_lancer(tmp_path, monkeypatch):
    from pcr_app.services import AppError, Services
    from pcr_app.workspace import Settings
    monkeypatch.setenv("PCR_APP_HOME", str(tmp_path / "home"))
    cwd = Path.cwd()
    sv = Services(Settings(workspace=str(tmp_path / "ws"), github_depot="o/r", github_jeton=""))
    with pytest.raises(AppError, match="jeton"):
        sv.lancer_sur_github()
    import os
    os.chdir(cwd)


def test_second_workflow_tableau_de_bord(services_gh):
    sv, fake = services_gh
    sv.lancer_sur_github(envoyer_mail=True, quoi="tableau")
    assert fake.workflow == "tableau-de-bord.yml" and fake.dispatches == [{"ref": "main", "inputs": {"send_mail": "true"}}]
    assert sv.execution_en_cours()["workflow"] == "tableau-de-bord.yml" and sv.suivre_github()["status"] == "in_progress"


def test_tableau_de_bord_genere_sur_l_ordinateur(services):
    r = services.generer_tableau_de_bord()
    assert r.ok and any(f.name.startswith("TABLEAU_DE_BORD_PCR-CI_du_") for f in services.documents())
