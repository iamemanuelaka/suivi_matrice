"""Démarrage de l'application : `python -m pcr_app` ou l'exécutable PCR-CI.exe.

Options : --navigateur  ouvre l'interface dans le navigateur au lieu d'une fenêtre (repli si la fenêtre ne s'affiche pas).
"""
import multiprocessing
import os
import sys


def _journal() -> None:
    """Exécutable sans console : sorties redirigées vers un journal (utile en cas de problème)."""
    if sys.stdout is None or sys.stderr is None:
        from .workspace import settings_dir
        d = settings_dir()
        d.mkdir(parents=True, exist_ok=True)
        f = open(d / "journal.log", "a", encoding="utf-8", buffering=1)   # noqa: SIM115
        sys.stdout = sys.stdout or f
        sys.stderr = sys.stderr or f


def main() -> None:
    multiprocessing.freeze_support()                      # indispensable pour l'exécutable Windows
    _journal()
    from nicegui import app, ui

    from . import app as interface
    from .workspace import resource_root

    interface.demarrer()
    icone = resource_root() / "pcr_app" / "icone.png"
    navigateur = "--navigateur" in sys.argv or os.environ.get("PCR_NAVIGATEUR") == "1"
    if not navigateur:
        app.native.window_args["min_size"] = (1100, 700)
    ui.run(title="Suivi des activités PCR-CI", favicon=icone if icone.is_file() else "📋", language="fr",
           native=not navigateur, window_size=None if navigateur else (1360, 860), reload=False,
           show_welcome_message=False, port=8765 if navigateur else None)
