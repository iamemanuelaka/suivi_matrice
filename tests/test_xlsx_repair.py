"""Garde-fous « Excel ouvre sans message de réparation ».

Contexte : Excel affichait « Enregistrements supprimés : Formule (sheet2.xml) » et « Enregistrements réparés : Dessin
(drawing1.xml, drawing2.xml) ». Les trois défauts étaient dans le MODÈLE lui-même (le générateur recopie ces parties
octet pour octet) : constante texte de 261 caractères dans une formule (limite Excel : 255) et `avLst` hors de l'espace
de noms `a:` dans les images.
"""
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from test_phase6_7 import ref  # noqa: F401,E402  (fixture partagée)
from pcr import xlsx_repair as X
from pcr.excel_gen import generate_dashboard
from pcr.xlsx_patch import Package

TEMPLATE = Path(__file__).resolve().parent.parent / "modeles" / "MATRICE_DE_SUIVI_DES_ACTIVITES_modele.xlsx"


def test_le_modele_livre_est_sain():
    assert X.diagnostiquer(X.lire_parts(TEMPLATE)) == []


def test_decoupage_des_litteraux_longs_conserve_le_texte():
    long = "é" * 300 + '"" guillemets "" ' + "x" * 100
    f = f'"debut "&pAlTaux&"{long}"&" fin"'
    g = X.decouper_litteraux(f)
    assert f != g and max(len(X._UNITES.findall(m.group(1))) for m in X._LIT.finditer(g)) <= 255
    rebuilt = "".join(m.group(1) for m in X._LIT.finditer(g))
    assert rebuilt == "debut " + long + " fin"
    assert X.decouper_litteraux('IF(A1="court","a","b")') == 'IF(A1="court","a","b")'      # sans effet sur les formules saines


def _abimer(parts: dict) -> dict:
    """Reproduit les défauts d'origine sur une copie du modèle, sans dépendre de la formule actuelle en B53."""
    p = dict(parts)
    p["xl/drawings/drawing1.xml"] = p["xl/drawings/drawing1.xml"].replace(b"<a:avLst/>", b"<avLst/>")
    p["xl/drawings/drawing2.xml"] = p["xl/drawings/drawing2.xml"].replace(b"<a:avLst/>", b"<avLst/>")
    from lxml import etree
    root = etree.fromstring(p["xl/worksheets/sheet2.xml"])
    b53 = next(c for c in root.iter(f"{{{X.S}}}c") if c.get("r") == "B53").find(f"{{{X.S}}}f")
    # Injecte explicitement le défaut Excel (constante de 261 caractères) au lieu de
    # supposer que le modèle livré contient encore l'ancienne formule découpée.
    b53.text = (b53.text or "") + '&"' + ("x" * 261) + '"'
    p["xl/worksheets/sheet2.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return p


def test_le_diagnostic_reconnait_exactement_les_defauts_signales_par_excel():
    pb = X.diagnostiquer(_abimer(X.lire_parts(TEMPLATE)))
    assert any("drawing1.xml" in x for x in pb) and any("drawing2.xml" in x for x in pb)
    assert any("sheet2.xml B53" in x and "261" in x for x in pb)


def test_reparation_automatique_a_l_ouverture_du_modele(tmp_path):
    parts = _abimer(X.lire_parts(TEMPLATE))
    abime = tmp_path / "abime.xlsx"
    with zipfile.ZipFile(TEMPLATE) as src, zipfile.ZipFile(abime, "w", zipfile.ZIP_DEFLATED) as z:
        for i in src.infolist():
            z.writestr(i.filename, parts[i.filename])
    assert X.diagnostiquer(X.lire_parts(abime))
    pkg = Package(abime)
    assert len(pkg.corrections) == 3
    sortie = tmp_path / "ok.xlsx"
    pkg.save(sortie)
    assert X.diagnostiquer(X.lire_parts(sortie)) == []


def test_dashboard_genere_sans_defaut_et_deux_executions_identiques(ref, tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    r1 = generate_dashboard(ref, TEMPLATE, a, datetime(2026, 11, 6, 8, 0))
    r2 = generate_dashboard(ref, TEMPLATE, b, datetime(2026, 11, 6, 8, 0))
    assert X.diagnostiquer(X.lire_parts(a)) == [] and r1["corrections_modele"] == []
    pa, pb = X.lire_parts(a), X.lire_parts(b)
    assert pa == pb                                  # 2e exécution : aucun doublon, aucune dérive
    # le fichier généré est ré-ouvrable comme modèle : pas de dégradation cumulative
    c = tmp_path / "c.xlsx"
    generate_dashboard(ref, a, c, datetime(2026, 11, 6, 8, 0))
    assert X.lire_parts(c) == pa


def test_generation_refusee_si_le_classeur_serait_defectueux(ref, tmp_path, monkeypatch):
    monkeypatch.setattr(X, "reparer_parts", lambda parts, noms=None: [])       # simule l'absence de réparation automatique
    abime = tmp_path / "abime.xlsx"
    parts = _abimer(X.lire_parts(TEMPLATE))
    with zipfile.ZipFile(TEMPLATE) as src, zipfile.ZipFile(abime, "w", zipfile.ZIP_DEFLATED) as z:
        for i in src.infolist():
            z.writestr(i.filename, parts[i.filename])
    out = tmp_path / "x.xlsx"
    with pytest.raises(ValueError, match="défauts détectés"):
        generate_dashboard(ref, abime, out)
    assert not out.exists()
