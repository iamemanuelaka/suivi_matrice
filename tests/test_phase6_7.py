import smtplib
import sys
import zipfile
from datetime import date, datetime
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr import mailer
from pcr.baseline import import_baseline
from pcr.config import load_config
from pcr.excel_gen import generate_dashboard, statut_chantier, upper_noaccent
from pcr.pipeline import alertes, run
from pcr.referentiel import Referentiel
from pcr.seed import seed
from pcr.store import Base
from pcr.xlsx_patch import Package

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "sources"
SOURCES = [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")]
TEMPLATE = ROOT / "modeles" / "MATRICE_DE_SUIVI_DES_ACTIVITES_modele.xlsx"


@pytest.fixture
def ref(tmp_path):
    cfg = load_config()
    cfg.data["chemins"] = {**cfg.data["chemins"], "base": str(tmp_path / "b.xlsx"), "sauvegardes": str(tmp_path / "sv"), "sorties": str(tmp_path / "out"),
                           "dashboard_modele": str(TEMPLATE)}
    cfg.root = Path("/")
    r = Referentiel(Base(tmp_path / "b.xlsx"), cfg)
    seed(r, SOURCES)
    import_baseline(r, SOURCES, date(2026, 10, 25))
    r.base.save()                                       # la base existe sur disque comme en production
    return r


def sub(i, code="BSBAF001", taux=20, jour="2026-11-02", obs="Fondation en cours"):
    return {"_id": i, "_submission_time": f"{jour}T09:00:00", "identification/date_actualisation": jour, "identification/code_activite": code,
            "avancement/observations_techniques": obs, "avancement/taux_execution": str(taux)}


# ---------- Excel ----------
def test_statut_derive_du_taux():
    assert [statut_chantier(t) for t in (None, 0, 3, 99, 100)] == ["Non démarré", "Non démarré", "En cours", "En cours", "Achevé"]
    assert upper_noaccent("Séguéla") == "SEGUELA" and upper_noaccent("Koun-fao") == "KOUN-FAO"


def test_dashboard_parties_inchangees_et_lignes_ecrites(ref, tmp_path):
    out = tmp_path / "d.xlsx"
    res = generate_dashboard(ref, TEMPLATE, out, datetime(2026, 11, 5, 8, 0))
    assert res["lignes"] == 42 and not res["formules_ignorees"]
    a, b = zipfile.ZipFile(TEMPLATE), zipfile.ZipFile(out)
    assert a.namelist() == b.namelist()
    changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
    assert sorted(changed) == ["xl/worksheets/sheet3.xml", "xl/worksheets/sheet6.xml"]       # matrice + META seulement ; graphiques/images/styles identiques
    sh = Package(out).sheet("MATRICE DE SUIVI")
    assert (sh.get(4, 4), sh.get(4, 5), sh.get(4, 6)) == ("KORO", "ETS NATHA SARL", 118356245)
    assert sh.has_formula(4, 13) and sh.get(30, 9) == "Marchés ruraux" and sh.get(30, 3) == "BERE" and sh.get(30, 2) == 27


def test_dashboard_ne_touche_a_aucune_formule(ref, tmp_path):
    out = tmp_path / "d.xlsx"
    generate_dashboard(ref, TEMPLATE, out)
    import re
    f = lambda p, n: re.findall(r"<f>.*?</f>", zipfile.ZipFile(p).read(n).decode("utf-8"))   # noqa: E731
    for n in ("xl/worksheets/sheet3.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet5.xml"):
        assert f(TEMPLATE, n) == f(out, n)


def test_dashboard_refuse_un_depassement_de_capacite(ref, tmp_path):
    for i in range(4):
        ref.create_activity("BS", "BAF")
    with pytest.raises(ValueError, match="place"):
        generate_dashboard(ref, TEMPLATE, tmp_path / "x.xlsx")


# ---------- cycle complet ----------
def test_cycle_complet_sans_mail(ref):
    res = run(ref, [sub(1), sub(2, code="NOPE")], date(2026, 11, 5), declencheur="test")
    assert res.ok and res.report.count("VALIDE") == 1 and res.report.count("REJETE") == 1 and len(res.documents) == 3
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
    assert sorted(p.get_filename() for p in m.iter_attachments()) == sorted(Path(d).name for d in res.documents)
    assert "05/11/2026" in m["Subject"] and "Activités en alerte" in m.get_body().get_content()


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
