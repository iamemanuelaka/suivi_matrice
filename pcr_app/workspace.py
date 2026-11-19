"""Espace de travail du logiciel et paramètres de l'utilisateur.

- Espace de travail : un dossier qui a la même organisation que le dépôt (config/, data/, modeles/, kobo/, sorties/).
  Par défaut « Documents/PCR-CI » : l'utilisateur y retrouve facilement les matrices générées.
- Paramètres (jetons, mots de passe) : dans le profil Windows de l'utilisateur (%APPDATA%/PCR-CI/parametres.json),
  jamais dans l'espace de travail ni dans le dépôt.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path

APP_NAME = "PCR-CI"
TEMPLATE_DIRS = ("config", "modeles", "kobo")


def resource_root() -> Path:
    """Dossier des fichiers livrés avec le logiciel (racine du dépôt, ou dossier interne de l'exécutable)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent.parent


def settings_dir() -> Path:
    base = os.environ.get("PCR_APP_HOME") or os.environ.get("APPDATA") or (Path.home() / ".config")
    return Path(base) / APP_NAME


def default_workspace() -> Path:
    docs = Path.home() / "Documents"
    return (docs if docs.is_dir() else Path.home()) / APP_NAME


@dataclass
class Settings:
    workspace: str = ""
    administrateur: str = ""
    # GitHub (stockage de la base, en coulisse)
    github_depot: str = "iamemanuelaka/suivi_matrice"   # propriétaire/dépôt
    github_branche: str = "main"
    github_jeton: str = ""
    # KoboToolbox
    kobo_url: str = "https://kf.kobotoolbox.org"
    kobo_asset_uid: str = ""
    kobo_jeton: str = ""
    # E-mail
    smtp_hote: str = ""
    smtp_port: str = "587"
    smtp_utilisateur: str = ""
    smtp_mot_de_passe: str = ""
    mail_expediteur: str = ""
    mail_destinataires: str = ""                         # adresses séparées par des virgules
    extra: dict = field(default_factory=dict)

    # --- persistance ---------------------------------------------------------
    @classmethod
    def path(cls) -> Path:
        return settings_dir() / "parametres.json"

    @classmethod
    def load(cls) -> "Settings":
        p = cls.path()
        data = {}
        if p.is_file():
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                data = {}
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        s = cls(**known)
        s.workspace = s.workspace or str(default_workspace())
        return s

    def save(self) -> None:
        p = self.path()
        p.parent.mkdir(parents=True, exist_ok=True)
        tmp = p.with_suffix(".tmp")
        tmp.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, p)

    # --- utilisation -----------------------------------------------------------
    @property
    def ws(self) -> Path:
        return Path(self.workspace or default_workspace())

    @property
    def github_ok(self) -> bool:
        return "/" in (self.github_depot or "")

    @property
    def kobo_ok(self) -> bool:
        return bool(self.kobo_asset_uid and self.kobo_jeton)

    @property
    def mail_ok(self) -> bool:
        return bool(self.smtp_hote and self.smtp_utilisateur and self.smtp_mot_de_passe and self.mail_destinataires)

    def apply_env(self) -> None:
        """Le moteur pcr lit ses identifiants dans les variables d'environnement (comme sur GitHub Actions)."""
        mapping = {"KOBO_URL": self.kobo_url, "KOBO_ASSET_UID": self.kobo_asset_uid, "KOBO_TOKEN": self.kobo_jeton,
                   "SMTP_HOST": self.smtp_hote, "SMTP_PORT": self.smtp_port, "SMTP_USER": self.smtp_utilisateur,
                   "SMTP_PASSWORD": self.smtp_mot_de_passe, "MAIL_FROM": self.mail_expediteur, "MAIL_TO": self.mail_destinataires}
        for k, v in mapping.items():
            if v:
                os.environ[k] = str(v).strip()
            else:
                os.environ.pop(k, None)


def prepare_workspace(ws: Path) -> Path:
    """Crée l'espace de travail et y copie les fichiers livrés (configuration, modèles, formulaire Kobo) s'ils manquent."""
    ws.mkdir(parents=True, exist_ok=True)
    src = resource_root()
    for d in TEMPLATE_DIRS:
        if (src / d).is_dir():
            for f in (src / d).rglob("*"):
                if f.is_file():
                    dest = ws / f.relative_to(src)
                    if not dest.exists():
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(f, dest)
    for d in ("data", "data/photos", "data/sauvegardes", "sorties"):
        (ws / d).mkdir(parents=True, exist_ok=True)
    return ws
