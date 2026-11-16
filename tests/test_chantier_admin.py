"""Informations gérées par l'administrateur : finances (décaissé), statut du chantier (4 valeurs), PGES, ONG, dates ; reprise des valeurs du modèle."""
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from test_phase6_7 import ref, TEMPLATE  # noqa: F401,E402
from pcr import xlsx_repair as X
from pcr.admin_file import CLEAR, COLUMNS, INFO_ONLY, apply_admin_file, export_admin_file
from pcr.excel_gen import COL, SHEET, generate_dashboard
from pcr.migration_modele import USER, migrer_valeurs_modele, reprise_automatique
from pcr.pipeline import alertes, run
from pcr.referentiel import STATUTS_CHANTIER, coerce, taux_financier
from pcr.store import Base
from pcr.word_gen import ordered_activities
from pcr.xlsx_patch import Package


def ligne_de(ref, code):
    acts = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    return 3 + 1 + [a["CODE_ACTIVITE"] for a in acts].index(code)


def test_une_base_existante_sans_les_nouvelles_colonnes_est_migree(tmp_path):
    p = tmp_path / "ancienne.xlsx"
    b = Base(p)
    ws = b.wb["Activites"]
    ws.delete_cols(18, 6)                                     # ancien schéma : s'arrête à Date_Modification
    b.save()
    assert [c.value for c in load_workbook(p)["Activites"][1]][-1] == "Date_Modification"
    b2 = Base(p)
    assert [c.value for c in b2.wb["Activites"][1]][-6:] == ["Montant_Decaisse", "Date_Remise_Site", "Statut_Chantier", "PGES",
                                                              "ONG_Mobilisees", "Date_Reception_Provisoire"]


def test_les_quatre_statuts_seulement():
    assert STATUTS_CHANTIER == ("Non démarré", "En cours", "Achevé", "Suspendu")
    assert coerce("Statut_Chantier", " suspendu ") == "Suspendu" and coerce("Statut_Chantier", "ACHEVE") == "Achevé"
    for interdit in ("Abandonné", "OS à reprendre", "Démarrage autorisé", "consortium"):
        with pytest.raises(ValueError, match="statut du chantier invalide"):
            coerce("Statut_Chantier", interdit)
    assert coerce("PGES", "validé") == "Validé"
    with pytest.raises(ValueError):
        coerce("PGES", "peut-être")


def test_le_modele_ne_contient_plus_les_statuts_en_trop():
    parts = X.lire_parts(TEMPLATE)
    texte = b"".join(v for k, v in parts.items() if k.endswith(".xml"))
    for mot in ("Abandonn", "OS à reprendre", "Démarrage autorisé"):
        assert mot.encode() not in texte, mot
    ws = load_workbook(TEMPLATE)["PARAMETRES"]
    categories = {ws.cell(r, 2).value for r in range(13, 42)} - {None}
    assert categories == {"En cours", "Non démarré", "Achevé", "Suspendu"}
    assert [ws.cell(r, 9).value for r in range(12, 17)] == ["Tous les statuts", "Non démarré", "En cours", "Achevé", "Suspendu"]
    assert X.diagnostiquer(parts) == []


def test_fichier_admin_champs_chantier(ref, tmp_path):
    f = tmp_path / "admin.xlsx"
    export_admin_file(ref, f)
    wb = load_workbook(f)
    ws = wb["Activites"]
    h = {c.value: i + 1 for i, c in enumerate(ws[1])}
    for col in ("Montant_Decaisse", "Taux_Financier", "Date_Remise_Site", "Statut_Chantier", "PGES", "ONG_Mobilisees", "Date_Reception_Provisoire"):
        assert col in h
    assert len(ws.data_validations.dataValidation) >= 5                                   # listes : catégorie, région, statut, statut chantier, PGES
    code = ws.cell(2, h["CODE_ACTIVITE"]).value
    montant = ref.get(code)["Montant_Marche"]
    ws.cell(2, h["Montant_Decaisse"]).value = montant // 4
    ws.cell(2, h["Statut_Chantier"]).value = "Suspendu"
    ws.cell(2, h["PGES"]).value = "Validé"
    ws.cell(2, h["ONG_Mobilisees"]).value = "ONG TEST"
    ws.cell(2, h["Date_Remise_Site"]).value = datetime(2026, 9, 1)
    ws.cell(2, h["Taux_Financier"]).value = 999                                           # colonne grise : ignorée
    ws.cell(3, h["Statut_Chantier"]).value = "Abandonné"                                  # refusé
    ws.cell(4, h["Montant_Decaisse"]).value = 10 ** 13                                    # dépasse le marché : refusé
    wb.save(f)
    rep = apply_admin_file(ref, f)
    a = ref.get(code)
    assert (a["Montant_Decaisse"], a["Statut_Chantier"], a["PGES"], a["ONG_Mobilisees"], a["Date_Remise_Site"]) == \
        (montant // 4, "Suspendu", "Validé", "ONG TEST", date(2026, 9, 1))
    assert taux_financier(a) == 25.0
    msgs = " | ".join(m for _, m in rep.errors)
    assert "statut du chantier invalide" in msgs and "dépasse le montant du marché" in msgs and len(rep.errors) == 2
    hist = [h for h in ref.history(code) if h["Champ_Modifie"] in ("Montant_Decaisse", "Statut_Chantier", "PGES")]
    assert len(hist) == 3 and all(h["Utilisateur"] and h["Date_Modification"] for h in hist)
    # #VIDER efface ; une cellule vide n'efface rien
    export_admin_file(ref, f)
    wb = load_workbook(f)
    ws = wb["Activites"]
    ws.cell(2, h["PGES"]).value = CLEAR
    ws.cell(2, h["ONG_Mobilisees"]).value = None
    wb.save(f)
    apply_admin_file(ref, f)
    a = ref.get(code)
    assert a["PGES"] is None and a["ONG_Mobilisees"] == "ONG TEST"


def test_le_dashboard_ecrit_les_champs_depuis_la_base_et_plus_de_valeur_figee(ref, tmp_path):
    code = ordered_activities(ref, "BS")[0]["CODE_ACTIVITE"]
    autre = ordered_activities(ref, "MR")[0]["CODE_ACTIVITE"]
    m = ref.get(code)["Montant_Marche"] or 100_000_000
    ref.update_activity(code, "t", {"Montant_Marche": m, "Montant_Decaisse": m // 2, "PGES": "En cours de validation", "ONG_Mobilisees": "ASAPSU",
                                    "Statut_Chantier": "Suspendu", "Date_Reception_Provisoire": date(2026, 10, 5)})
    out = tmp_path / "d.xlsx"
    generate_dashboard(ref, TEMPLATE, out, datetime(2026, 11, 6, 8, 0))
    sh = Package(out).sheet(SHEET)
    r = ligne_de(ref, code)
    assert sh.get(r, COL["decaisse"]) == m // 2 and sh.get(r, COL["taux_fin"]) == 50.0
    assert (sh.get(r, COL["pges"]), sh.get(r, COL["ong"]), sh.get(r, COL["statut"])) == ("En cours de validation", "ASAPSU", "Suspendu")
    # une activité sans donnée administrateur n'hérite RIEN du modèle d'origine (30 000 000, « Validé », « ASAPSU »… étaient figés par position)
    r2 = ligne_de(ref, autre)
    for k in ("decaisse", "taux_fin", "remise", "pges", "ong", "reception", "photo", "resultats"):
        assert sh.get(r2, COL[k]) in (None, ""), k
    autres = [x for x in ordered_activities(ref, "BS") if x["CODE_ACTIVITE"] != code]
    assert all(sh.get(ligne_de(ref, x["CODE_ACTIVITE"]), COL["decaisse"]) in (None, "") for x in autres)


def test_reprise_des_valeurs_du_modele(ref):
    rap = migrer_valeurs_modele(ref, TEMPLATE, dry_run=True)
    assert len(rap["reprises"]) == 26 and rap["non_rapprochees"] == [] and all(a["Montant_Decaisse"] is None for a in ref.activities())
    code = rap["reprises"][0][0]
    ref.update_activity(code, "admin", {"PGES": "Sans PGES"})                              # saisie administrateur : ne sera jamais écrasée
    rap = migrer_valeurs_modele(ref, TEMPLATE)
    assert len(rap["reprises"]) >= 25 and ref.get(code)["PGES"] == "Sans PGES"
    assert [c for c, ch in rap["reprises"] if ch.get("Statut_Chantier")] == ["BSBAG002"]   # seul « Suspendu » est repris (les autres se déduisent du taux)
    assert any(a["Montant_Decaisse"] == 30_000_000 for a in ref.activities())
    assert migrer_valeurs_modele(ref, TEMPLATE)["reprises"] == []                         # idempotent
    assert sum(1 for h in ref.history() if h["Utilisateur"] == USER) > 50


def test_reprise_automatique_une_seule_fois(ref, monkeypatch):
    infos = reprise_automatique(ref, TEMPLATE)
    assert infos and "26" in infos[0]
    code = next(a["CODE_ACTIVITE"] for a in ref.activities() if a["PGES"])
    ref.update_activity(code, "admin", {}, clear=["PGES"])                                  # l'administrateur efface volontairement
    assert reprise_automatique(ref, TEMPLATE) == [] and ref.get(code)["PGES"] is None      # ... et la reprise ne se rejoue pas


def test_cycle_complet_reprend_puis_genere(ref):
    res = run(ref, [], date(2026, 11, 9), send_mail=False)
    assert res.ok and any("Reprise unique" in a for a in res.avertissements)
    res2 = run(ref, [], date(2026, 11, 9), send_mail=False)
    assert res2.ok and res2.avertissements == []


def test_les_alertes_python_ignorent_un_chantier_acheve_par_l_administrateur(ref):
    avant = alertes(ref, date(2026, 11, 6))
    assert avant
    code = avant[0].split(" ")[0]
    ref.update_activity(code, "admin", {"Statut_Chantier": "Achevé"})
    assert not any(x.startswith(code) for x in alertes(ref, date(2026, 11, 6)))             # comme le tableau de bord (activité achevée : non évaluée)


@pytest.mark.skipif(not shutil.which("soffice"), reason="LibreOffice absent")
def test_dashboard_recalcule_statuts_pges_ong_et_alertes(ref, tmp_path):
    ref.base.wb  # noqa: B018
    migrer_valeurs_modele(ref, TEMPLATE)
    suspendu = next(a["CODE_ACTIVITE"] for a in ref.activities() if a["Statut_Chantier"] == "Suspendu")
    out = tmp_path / "d.xlsx"
    generate_dashboard(ref, TEMPLATE, out, datetime(2026, 11, 6, 8, 0))
    lo = tmp_path / "lo"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(lo), str(out)], capture_output=True, timeout=250)
    wb = load_workbook(lo / "d.xlsx", data_only=True)
    C, T = wb["CALCULS"], wb["TABLEAU DE BORD"]
    statuts = {C[f"DC{r}"].value for r in range(5, 40)} - {None, ""}
    assert statuts <= {"Tous les statuts", "Non démarré", "En cours", "Achevé", "Suspendu"} and "Suspendu" in statuts
    cats = {C[f"N{r}"].value for r in range(5, 60)} - {None, ""}
    assert cats <= {"Non démarré", "En cours", "Achevé", "Suspendu", "Non renseigné"}
    assert T["AA9"].value.startswith(f"{len(alertes(ref, date.today()))} /")
    assert not [c.coordinate for s in wb for row in s.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("#") and c.value != "#N/A"]
