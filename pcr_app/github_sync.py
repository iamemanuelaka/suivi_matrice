"""Synchronisation de l'espace de travail avec le dépôt GitHub, sans git installé (API REST GitHub).

Principe (comparaison des empreintes « blob » de git, calculables en local) :
  - pour chaque fichier suivi, on mémorise l'empreinte de la dernière version synchronisée ;
  - local inchangé, GitHub modifié (ex. traitement automatique du lundi)  -> on télécharge ;
  - local modifié, GitHub inchangé                                       -> on envoie ;
  - les deux modifiés                                                    -> CONFLIT : rien n'est écrasé, l'utilisateur choisit.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import requests

API = "https://api.github.com"
# Ce qui est récupéré depuis GitHub : configuration, base, photos, modèles, formulaire Kobo.
PULL_PREFIXES = ("config/", "data/", "modeles/", "kobo/")
PULL_EXCLUDE = ("data/sauvegardes/",)
# Ce que le logiciel renvoie vers GitHub : la base (seul fichier modifié par l'utilisateur).
PUSH_FILES = ("data/base_pcr.xlsx",)
STATE_FILE = ".synchro_github.json"


class SyncError(Exception):
    pass


class SyncConflict(SyncError):
    pass


def blob_sha(path: Path) -> str | None:
    """Empreinte git d'un fichier local (identique à celle que donne l'API GitHub)."""
    if not path.is_file():
        return None
    data = path.read_bytes()
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


@dataclass
class SyncReport:
    telecharges: list = field(default_factory=list)
    envoyes: list = field(default_factory=list)
    conflits: list = field(default_factory=list)

    def text(self) -> str:
        parts = []
        if self.telecharges:
            parts.append(f"{len(self.telecharges)} fichier(s) récupéré(s) depuis GitHub")
        if self.envoyes:
            parts.append(f"{len(self.envoyes)} fichier(s) envoyé(s) vers GitHub")
        if self.conflits:
            parts.append(f"CONFLIT sur {', '.join(self.conflits)}")
        return " ; ".join(parts) or "Déjà à jour"


class GitHubSync:
    def __init__(self, repo: str, branch: str, token: str, root: Path, session: requests.Session | None = None):
        self.repo, self.branch, self.root = repo.strip().strip("/"), branch or "main", Path(root)
        self.s = session or requests.Session()
        self.s.headers.update({"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28",
                               "User-Agent": "PCR-CI-application"})
        if token:
            self.s.headers["Authorization"] = f"Bearer {token.strip()}"
        self.state_path = self.root / STATE_FILE

    # --- état local -----------------------------------------------------------
    def _state(self) -> dict:
        try:
            return json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {"fichiers": {}}

    def _save_state(self, st: dict) -> None:
        st["derniere_synchro"] = datetime.now().isoformat(timespec="seconds")
        tmp = self.state_path.with_suffix(".tmp")
        tmp.write_text(json.dumps(st, indent=1, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, self.state_path)

    def last_sync(self) -> str | None:
        return self._state().get("derniere_synchro")

    # --- API ------------------------------------------------------------------
    def _get(self, url, **kw):
        try:
            r = self.s.get(url, timeout=60, **kw)
        except requests.RequestException as e:
            raise SyncError(f"GitHub injoignable : {e}") from e
        if r.status_code == 401:
            raise SyncError("Jeton GitHub invalide ou expiré (Paramètres → GitHub).")
        if r.status_code == 404:
            raise SyncError(f"Dépôt « {self.repo} » ou branche « {self.branch} » introuvable (ou dépôt privé sans jeton).")
        if r.status_code == 403:
            raise SyncError("Accès refusé par GitHub (limite de requêtes atteinte ou jeton sans droit de lecture).")
        if r.status_code != 200:
            raise SyncError(f"GitHub a répondu HTTP {r.status_code} : {r.text[:200]}")
        return r

    def remote_tree(self, prefixes: tuple = PULL_PREFIXES) -> dict[str, str]:
        """{chemin: empreinte} des fichiers du dépôt sous les préfixes donnés (par défaut : ceux que l'application suit)."""
        r = self._get(f"{API}/repos/{self.repo}/git/trees/{self.branch}", params={"recursive": "1"})
        out = {}
        for e in r.json().get("tree", []):
            p = e.get("path", "")
            if e.get("type") == "blob" and p.startswith(prefixes) and not p.startswith(PULL_EXCLUDE) and not p.endswith("/.gitkeep"):
                out[p] = e["sha"]
        return out

    # --- traitement automatique (GitHub Actions) ---------------------------------
    def _post(self, url, body, ok=(200, 201, 204)):
        try:
            r = self.s.post(url, json=body, timeout=60)
        except requests.RequestException as e:
            raise SyncError(f"GitHub injoignable : {e}") from e
        if r.status_code in (401, 403, 404):
            raise SyncError("Lancement refusé par GitHub : le jeton doit avoir la permission « Actions : Read and write » "
                            "(en plus de « Contents : Read and write »).")
        if r.status_code == 422:
            raise SyncError(f"Le workflow ne peut pas être lancé à la main (déclencheur workflow_dispatch absent ?) : {r.text[:200]}")
        if r.status_code not in ok:
            raise SyncError(f"GitHub a répondu HTTP {r.status_code} : {r.text[:200]}")
        return r

    def run_ids(self, workflow: str) -> set:
        r = self._get(f"{API}/repos/{self.repo}/actions/workflows/{workflow}/runs", params={"per_page": "20"})
        return {x["id"] for x in r.json().get("workflow_runs", [])}

    def dispatch(self, workflow: str, inputs: dict) -> None:
        self._post(f"{API}/repos/{self.repo}/actions/workflows/{workflow}/dispatches", {"ref": self.branch, "inputs": inputs})

    def new_run(self, workflow: str, avant: set) -> dict | None:
        """Exécution apparue depuis le lancement (None tant que GitHub ne l'a pas encore créée)."""
        r = self._get(f"{API}/repos/{self.repo}/actions/workflows/{workflow}/runs",
                      params={"per_page": "10", "event": "workflow_dispatch", "branch": self.branch})
        nouvelles = [x for x in r.json().get("workflow_runs", []) if x["id"] not in avant]
        return min(nouvelles, key=lambda x: x["id"]) if nouvelles else None

    def run_status(self, run_id: int) -> dict:
        """État d'une exécution et, si elle est en cours, l'étape en train de tourner."""
        run = self._get(f"{API}/repos/{self.repo}/actions/runs/{run_id}").json()
        etape = None
        if run.get("status") != "completed":
            try:
                jobs = self._get(f"{API}/repos/{self.repo}/actions/runs/{run_id}/jobs").json().get("jobs", [])
                for j in jobs:
                    for st in j.get("steps") or []:
                        if st.get("status") == "in_progress":
                            etape = st.get("name")
            except SyncError:
                pass
        return {"id": run_id, "status": run.get("status"), "conclusion": run.get("conclusion"),
                "url": run.get("html_url"), "etape": etape}

    def download_new(self, prefix: str, avant: dict) -> list[str]:
        """Télécharge les fichiers sous `prefix` créés ou modifiés sur GitHub depuis l'instantané `avant` ({chemin: empreinte})."""
        apres = self.remote_tree((prefix,))
        recus = []
        for path, sha in sorted(apres.items()):
            if avant.get(path) != sha and blob_sha(self.root / path) != sha:
                self._download(path, sha)
                recus.append(path)
        return recus

    def _download(self, path: str, sha: str) -> None:
        r = self._get(f"{API}/repos/{self.repo}/git/blobs/{sha}", headers={"Accept": "application/vnd.github.raw"})
        dest = self.root / path
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_name(dest.name + ".telechargement")
        tmp.write_bytes(r.content)
        if blob_sha(tmp) != sha:
            tmp.unlink(missing_ok=True)
            raise SyncError(f"Téléchargement incomplet : {path}")
        os.replace(tmp, dest)

    def _upload(self, path: str, remote_sha: str | None, message: str) -> str:
        body = {"message": message, "branch": self.branch,
                "content": base64.b64encode((self.root / path).read_bytes()).decode()}
        if remote_sha:
            body["sha"] = remote_sha
        try:
            r = self.s.put(f"{API}/repos/{self.repo}/contents/{path}", json=body, timeout=120)
        except requests.RequestException as e:
            raise SyncError(f"GitHub injoignable : {e}") from e
        if r.status_code in (409, 422) and "sha" in r.text:
            raise SyncConflict(f"{path} a été modifié sur GitHub pendant l'envoi : relancez la synchronisation.")
        if r.status_code in (401, 403, 404):
            raise SyncError("Envoi refusé par GitHub : le jeton doit avoir le droit « Contents : Read and write » sur le dépôt.")
        if r.status_code not in (200, 201):
            raise SyncError(f"Envoi de {path} refusé (HTTP {r.status_code}) : {r.text[:200]}")
        return r.json()["content"]["sha"]

    # --- synchronisation --------------------------------------------------------
    def sync(self, push: bool = True, message: str = "Mise à jour depuis l'application PCR-CI", prefer: str | None = None) -> SyncReport:
        """Récupère ce qui a changé sur GitHub, puis envoie la base si elle a été modifiée ici.
        prefer = "github" ou "local" pour résoudre un conflit (la version écartée est sauvegardée)."""
        st, rep = self._state(), SyncReport()
        known = st.setdefault("fichiers", {})
        remote = self.remote_tree()
        for path, rsha in sorted(remote.items()):
            lsha, last = blob_sha(self.root / path), known.get(path)
            if lsha == rsha:
                known[path] = rsha
                continue
            local_modifie = lsha is not None and lsha != last
            if path in PUSH_FILES and local_modifie and last is not None and rsha != last:      # les deux ont changé
                if prefer == "local":
                    continue                                         # traité à l'envoi ci-dessous
                if prefer != "github":
                    rep.conflits.append(path)
                    continue
            if path in PUSH_FILES and local_modifie and last is not None and rsha == last:
                continue                                             # modifié ici seulement : sera envoyé
            if lsha is not None:
                self._backup(path)
            self._download(path, rsha)
            known[path] = rsha
            rep.telecharges.append(path)
        if push and not rep.conflits:
            for path in PUSH_FILES:
                lsha = blob_sha(self.root / path)
                if lsha is None or lsha == remote.get(path):
                    continue
                if remote.get(path) != known.get(path) and prefer != "local":
                    continue                                         # version GitHub plus récente non récupérée : ne jamais l'écraser
                known[path] = self._upload(path, remote.get(path), message)
                rep.envoyes.append(path)
        self._save_state(st)
        if rep.conflits:
            raise SyncConflict("La base a été modifiée à la fois ici et sur GitHub (traitement automatique du lundi ?). "
                               "Choisissez la version à garder dans Paramètres → GitHub.")
        return rep

    def _backup(self, path: str) -> None:
        src = self.root / path
        if path in PUSH_FILES and src.is_file():
            d = self.root / "data" / "sauvegardes"
            d.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, d / f"{src.stem}_avant_synchro_{datetime.now():%Y%m%d_%H%M%S}{src.suffix}")

    def check(self) -> str:
        """Test de connexion (Paramètres)."""
        r = self._get(f"{API}/repos/{self.repo}")
        perms = r.json().get("permissions") or {}
        if perms.get("push"):
            return f"Connexion réussie à {self.repo} (lecture et écriture)."
        return f"Connexion réussie à {self.repo}, mais en LECTURE seule : la base ne pourra pas être envoyée."
