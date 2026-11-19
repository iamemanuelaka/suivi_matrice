import smtplib
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr import mailer
from pcr.baseline import import_baseline
from pcr.config import load_config
from pcr.pipeline import alertes, run, run_tableau_de_bord
from pcr.referentiel import statut_deduit
from pcr.tableau_de_bord_xlsx import generer
from pcr.referentiel import Referentiel
from pcr.seed import seed
from pcr.store import Base
from pcr.word_gen import ordered_activities

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "sources"
SOURCES = [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")]


@pytest.fixture
def ref(tmp_path):
    cfg = load_config()
    cfg.data["chemins"] = {**cfg.data["chemins"], "base": str(tmp_path / "b.xlsx"), "sauvegardes": str(tmp_path / "sv"), "sorties": str(tmp_path / "out")}
    cfg.root = Path("/")
    r = Referentiel(Base(tmp_path / "b.xlsx"), cfg)
    seed(r, SOURCES)
    import_baseline(r, SOURCES, date(2026, 10, 25))
    r.base.save()                                       # la base existe sur disque comme en production
    return r


def sub(i, code="BSBAF001", taux=20, jour="2026-11-02", obs="Fondation en cours"):
    return {"_id": i, "_submission_time": f"{jour}T09:00:00", "identification/date_actualisation": jour, "identification/code_activite": code,
            "avancement/observations_techniques": obs, "avancement/taux_execution": str(taux)}


# ---------- tableau de bord Excel (généré de A à Z) ----------
def test_statut_derive_du_taux():
    assert [statut_deduit(t) for t in (None, 0, 3, 99, 100)] == ["Non démarré", "Non démarré", "En cours", "En cours", "Achevé"]


def test_tableau_de_bord_trois_feuilles_filtres_et_donnees(ref, tmp_path):
    res = generer(ref, tmp_path / "tb.xlsx", date(2026, 11, 5), datetime(2026, 11, 5, 8, 0))
    assert res["activites"] == 42
    wb = load_workbook(tmp_path / "tb.xlsx")
    assert wb.sheetnames == ["TABLEAU DE BORD", "ALERTES", "DONNÉES"] and wb["DONNÉES"].sheet_state == "hidden"
    tb = wb["TABLEAU DE BORD"]
    assert sorted(str(dv.sqref) for dv in tb.data_validations.dataValidation) == ["B10", "B13", "B16", "B19"]      # 4 filtres
    assert [tb[c].value for c in ("B10", "B13", "B16", "B19")] == ["Toutes les natures", "Toutes les régions",
                                                                    "Toutes les entreprises", "Tous les statuts"]
    assert len(tb._charts) == 1 and len(tb._images) == 1                          # graphique + logo
    assert {"selNature", "selRegion", "selEntreprise", "selStatut", "dFiltre", "lstRegion"} <= set(wb.defined_names)
    do = wb["DONNÉES"]
    entetes = [c.value for c in do[1]]
    codes = [r[1] for r in do.iter_rows(min_row=2, max_row=43, values_only=True)]
    assert len(codes) == 42 and codes[0] and do.cell(44, 2).value is None
    reg = entetes.index("Liste : Région") + 1
    assert [do.cell(r, reg).value for r in range(2, 14)] == ["Toutes les régions", *sorted(ref.cfg.regions.values(), key=str.casefold)]
    al = wb["ALERTES"]
    assert al["A1"].value.startswith("Alertes et points d’attention — Voir tout (") and "Alertes" in al.tables


def test_tableau_de_bord_suit_les_valeurs_de_l_administrateur(ref, tmp_path):
    code = ref.activities()[0]["CODE_ACTIVITE"]
    ref.update_activity(code, changes={"Statut_Chantier": "Suspendu", "PGES": "Validé", "ONG_Mobilisees": "ONG AGIR POUR LES ODD",
                                       "Montant_Decaisse": 1_000_000})
    generer(ref, tmp_path / "tb.xlsx", date(2026, 11, 5))
    do = load_workbook(tmp_path / "tb.xlsx")["DONNÉES"]
    e = [c.value for c in do[1]]
    ligne = next(r for r in do.iter_rows(min_row=2, values_only=True) if r[1] == code)
    assert ligne[e.index("Statut")] == "Suspendu" and ligne[e.index("PGES")] == "Validé"
    assert ligne[e.index("ONG (clé)")] == "AGIR POUR ODD" and ligne[e.index("Décaissé")] == 1_000_000     # variante d'ONG ramenée au nom retenu


def test_second_traitement_envoie_le_tableau_de_bord(ref, monkeypatch):
    envoyes = []
    monkeypatch.setattr("pcr.pipeline.send", lambda m: envoyes.append(m))
    monkeypatch.setenv("MAIL_TO", "a@b.org")
    res = run_tableau_de_bord(ref, date(2026, 11, 5), send_mail=True, declencheur="test")
    assert res.ok and [Path(d).name for d in res.documents] == ["TABLEAU_DE_BORD_PCR-CI_du_05-11-26.xlsx"]
    m = envoyes[0]
    assert [p.get_filename() for p in m.iter_attachments()] == ["TABLEAU_DE_BORD_PCR-CI_du_05-11-26.xlsx"]
    assert m["Subject"] == "Tableau de bord PCR-CI – situation au 05/11/2026" and "Activités en alerte" in m.get_body().get_content()
    assert ref.base.rows("Journal_Executions")[-1]["Declencheur"] == "tableau de bord (test)"


# ---------- cycle complet ----------
def test_cycle_complet_sans_mail(ref):
    res = run(ref, [sub(1), sub(2, code="NOPE")], date(2026, 11, 5), declencheur="test")
    assert res.ok and res.report.count("VALIDE") == 1 and res.report.count("REJETE") == 1 and len(res.documents) == 2
    assert all(Path(d).exists() for d in res.documents)
    j = ref.base.rows("Journal_Executions")[-1]
    assert j["Resultat"] == "OK" and j["Traitees"] == 2 and "docx" in j["Documents"]
    assert list(Path(ref.cfg.path("sauvegardes")).glob("*.xlsx"))                      # sauvegarde faite avant écriture


def test_relance_idempotente(ref):
    run(ref, [sub(1)], date(2026, 11, 5))
    n = len(ref.base.rows("Suivi_Activites"))
    res = run(ref, [sub(1)], date(2026, 11, 5))
    assert res.report.ignorees == 1 and len(ref.base.rows("Suivi_Activites")) == n


def test_echec_journalise_et_aucun_envoi(ref, monkeypatch):
    envoyes = []
    monkeypatch.setattr("pcr.pipeline.send", lambda m: envoyes.append(m))
    monkeypatch.setenv("MAIL_TO", "a@b.org")
    monkeypatch.setattr("pcr.pipeline.generate_all", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("panne Word")))
    res = run(ref, [sub(1)], date(2026, 11, 5), send_mail=True)
    assert not res.ok and "panne Word" in res.erreur
    assert ref.base.rows("Journal_Executions")[-1]["Resultat"] == "ÉCHEC"
    assert len(envoyes) == 1 and envoyes[0]["Subject"].startswith("[ÉCHEC]") and not list(envoyes[0].iter_attachments())


def test_envoi_email_avec_pieces_jointes(ref, monkeypatch):
    envoyes = []
    monkeypatch.setattr("pcr.pipeline.send", lambda m: envoyes.append(m))
    monkeypatch.setenv("MAIL_TO", "a@b.org, c@d.org")
    res = run(ref, [sub(1)], date(2026, 11, 5), send_mail=True)
    m = envoyes[0]
    assert res.ok and res.email == "envoyé à a@b.org, c@d.org" and m["To"] == "a@b.org, c@d.org"
    # le traitement des matrices n'envoie que les 2 matrices Word (le tableau de bord a son propre traitement)
    assert sorted(p.get_filename() for p in m.iter_attachments()) == sorted(Path(d).name for d in res.documents)
    assert all(Path(d).suffix == ".docx" for d in res.documents) and len(res.documents) == 2
    corps = m.get_body().get_content()
    assert "05/11/2026" in m["Subject"] and "Activités en alerte" in corps and "tableau de bord" not in corps


def test_email_sans_destinataire_refuse(ref, monkeypatch):
    monkeypatch.delenv("MAIL_TO", raising=False)
    res = run(ref, [sub(1)], date(2026, 11, 5), send_mail=True)
    assert not res.ok and "destinataire" in res.erreur


def test_smtp_incomplet_erreur_claire(monkeypatch):
    for v in ("SMTP_HOST", "SMTP_USER", "SMTP_PASSWORD"):
        monkeypatch.delenv(v, raising=False)
    with pytest.raises(mailer.MailError, match="SMTP"):
        mailer.send(mailer.build_message("s", "b", [], "x@y", ["a@b"]))


def test_alertes_meme_regle_que_le_tableau_de_bord(ref):
    a = alertes(ref, date(2026, 11, 5))
    assert any(x.startswith("BSBAF002") for x in a)                                    # 5 % d'exécution pour 65 % du délai
    assert not any(x.startswith("MRGON001") for x in a)                                # 45 % d'exécution : hors alerte
    assert not any(x.startswith("BSWOR003") for x in a)                                # début absent : jamais calculé artificiellement
