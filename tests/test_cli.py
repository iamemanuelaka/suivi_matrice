"""Garde-fou : le point d'entrée `python -m pcr` doit toujours pouvoir être importé.

Un import cassé dans cli.py (fonction supprimée, nom changé) n'était détecté par aucun test et
ne se voyait qu'au lancement du workflow GitHub.
"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def test_cli_importable_et_commandes_declarees(capsys):
    from pcr import cli
    with pytest.raises(SystemExit) as e:
        cli.main(["--help"])
    assert e.value.code == 0
    sortie = capsys.readouterr().out
    for commande in ("lancer", "kobo-sync", "charger-reference", "generer-word", "generer-dashboard"):
        assert commande in sortie


def test_toutes_les_fonctions_importees_par_cli_existent():
    import ast
    import importlib
    arbre = ast.parse((Path(__file__).resolve().parent.parent / "pcr" / "cli.py").read_text(encoding="utf-8"))
    for noeud in ast.walk(arbre):
        if isinstance(noeud, ast.ImportFrom) and noeud.level == 1:
            module = importlib.import_module(f"pcr.{noeud.module}")
            for nom in noeud.names:
                assert hasattr(module, nom.name), f"pcr.{noeud.module}.{nom.name} introuvable"
