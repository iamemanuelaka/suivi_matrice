"""Fichier principal de l'interface pour les tests (simulation d'utilisateur NiceGUI, sans navigateur)."""
import importlib
import os

from nicegui import ui

import pcr_app.app as interface
from pcr_app.services import Services
from pcr_app.workspace import Settings

interface = importlib.reload(interface)                       # ré-enregistre les pages après la remise à zéro de NiceGUI
interface.demarrer(Services(Settings(workspace=os.environ["PCR_TEST_WS"], github_depot="", kobo_asset_uid="")))
ui.run(storage_secret="test", reload=False)
