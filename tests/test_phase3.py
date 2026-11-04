import csv
import io
import os
import re
import shutil
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pyxform.xls2xform import convert

from pcr.config import load_config
from pcr.kobo_form import (CSV_COLUMNS, CSV_NAME, KoboClient, KoboError, SURVEY, activities_csv_rows, build_xlsform, csv_bytes,
                           sync_to_kobo, write_activities_csv)
from pcr.referentiel import Referentiel
from pcr.store import Base


@pytest.fixture
def ref(tmp_path):
    return Referentiel(Base(tmp_path / "b.xlsx"), load_config())


def suivi(ref, id_, code, jour, taux, statut="VALIDE"):
    ref.base.append("Suivi_Activites", {"ID_SOUMISSION_KOBO": id_, "CODE_ACTIVITE": code, "Date_Actualisation": jour, "Taux_Execution_Physique": taux,
                                        "Date_Importation": datetime(2026, 11, 1, 8, 0), "Statut_Traitement": statut})


# ---------------- liste des activités ----------------
def test_csv_colonnes_sans_bom_et_une_ligne_par_activite_active(ref):
    a = ref.create_activity("BS", "BAF", Entreprise="ETS NATHA SARL", Departement="Koro", Montant_Marche=118356245, Delai_Mois=4, Date_Debut="17/08/2026")
    b = ref.create_activity("BS", "BAF")
    ref.archive(b)
    data = csv_bytes(activities_csv_rows(ref))
    assert not data.startswith(b"\xef\xbb\xbf") and data.startswith(b"name,label,")
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8"))))
    assert [r["name"] for r in rows] == [a]                      # l'activité archivée n'est plus proposée
    assert rows[0]["label"] == "BSBAF001 · Bafing · Koro — ETS NATHA SARL"
    assert "118 356 245 FCFA" in rows[0]["ref_marche"] and "Fin prévue : 17/12/2026" in rows[0]["ref_marche"]


def test_activite_incomplete_proposee_avec_libelles_neutres(ref):
    ref.create_activity("MR", "BAF")
    r = activities_csv_rows(ref)[0]
    assert r["label"] == "MRBAF001 · Bafing · — — titulaire non renseigné"
    assert r["ref_marche"] == "Informations contractuelles non encore renseignées."
    assert r["ref_suivi"] == "Aucun suivi enregistré pour cette activité."


def test_csv_sans_retour_a_la_ligne_et_virgules_protegees(ref):
    ref.create_activity("BS", "BAG", Designation="Ligne 1\nLigne 2, avec virgule", Localite="A, B", Entreprise='SARL "X"')
    raw = csv_bytes(activities_csv_rows(ref)).decode("utf-8")
    assert len(raw.strip().splitlines()) == 2                    # en-tête + 1 ligne : aucun retour à la ligne dans les cellules
    row = next(csv.DictReader(io.StringIO(raw)))
    assert row["ref_titre"].endswith("Ligne 1 Ligne 2, avec virgule") and 'SARL "X"' in row["label"]


def test_dernier_suivi_valide_affiche_les_rejets_sont_ignores(ref):
    c = ref.create_activity("BS", "BAF")
    suivi(ref, 1, c, date(2026, 10, 5), 20)
    suivi(ref, 2, c, date(2026, 10, 20), 45)
    suivi(ref, 3, c, date(2026, 10, 12), 30)                     # visite plus ancienne reçue plus tard : ne doit pas l'emporter
    suivi(ref, 4, c, date(2026, 10, 28), 99, statut="REJETE")    # rejeté : jamais affiché
    assert activities_csv_rows(ref)[0]["ref_suivi"] == "Dernier suivi enregistré : 45 % au 20/10/2026"


def test_ecriture_du_fichier(ref, tmp_path):
    ref.create_activity("BS", "BAF")
    assert write_activities_csv(ref, tmp_path / "x" / CSV_NAME) == 1 and (tmp_path / "x" / CSV_NAME).exists()


# ---------------- formulaire ----------------
@pytest.fixture(scope="module")
def xform(tmp_path_factory):
    p = build_xlsform(tmp_path_factory.mktemp("f") / "q.xlsx", version="2026110401")
    res = convert(xlsform=str(p), warnings=[])
    assert res.warnings == []
    return res.xform


def test_champs_du_suivi_distincts_et_obligatoires(xform):
    for champ in ("code_activite", "date_actualisation", "observations_techniques", "taux_execution", "commentaires", "difficultes", "recommandations"):
        assert re.search(rf'nodeset="/data/\w+/{champ}"', xform), champ
    for champ in ("code_activite", "date_actualisation", "observations_techniques", "taux_execution", "nom_agent"):
        assert re.search(rf'nodeset="/data/\w+/{champ}"[^>]*required="true\(\)"', xform), f"{champ} doit être obligatoire"
    for champ in ("commentaires", "difficultes", "recommandations"):
        assert not re.search(rf'nodeset="/data/\w+/{champ}"[^>]*required=', xform), f"{champ} doit rester facultatif"


def test_aucune_information_fixe_n_est_redemandee(xform):
    saisies = {m for m in re.findall(r'nodeset="/data/(?:\w+/)?(\w+)"', xform)} - {"start", "end", "meta", "instanceID", "instanceName"}
    fixes = {"entreprise", "montant", "montant_marche", "region", "departement", "localite", "delai", "delai_mois", "date_os", "date_debut", "designation"}
    assert not (saisies & fixes)


def test_controles_de_saisie(xform):
    assert 'constraint=". &gt;= 0 and . &lt;= 100"' in xform                # taux 0..100
    assert 'constraint=". &lt;= today()"' in xform                         # pas de date future
    assert '<itemset nodeset="instance(\'activites\')/root/item">' in xform  # liste externe alimentée par le CSV
    assert "odk-instance-first-load" in xform and "today()" in xform


def test_colonnes_lues_par_pulldata_existent_dans_le_csv():
    fichier = CSV_NAME[:-4]
    lues = []
    for row in SURVEY:
        calc = row[9]
        if "pulldata" in calc:
            m = re.fullmatch(rf"pulldata\('{fichier}', '(\w+)', 'name', \$\{{code_activite\}}\)", calc)
            assert m, calc
            lues.append(m[1])
    assert lues and set(lues) <= set(CSV_COLUMNS)


@pytest.mark.skipif(not (shutil.which("java") and Path(os.environ.get("ODK_VALIDATE_JAR", "/tmp/validate.jar")).exists()), reason="ODK Validate non disponible")
def test_formulaire_valide_par_odk_validate(tmp_path):
    p = build_xlsform(tmp_path / "q.xlsx")
    out = tmp_path / "f.xml"
    out.write_text(convert(xlsform=str(p), warnings=[]).xform, encoding="utf-8")
    r = subprocess.run(["java", "-jar", os.environ.get("ODK_VALIDATE_JAR", "/tmp/validate.jar"), str(out)], capture_output=True, text=True)
    assert "Xform is valid" in r.stdout + r.stderr


# ---------------- synchronisation Kobo (serveur simulé) ----------------
class Resp:
    def __init__(self, status=200, body=None, text=""):
        self.status_code, self._b, self.text = status, body or {}, text

    def json(self):
        return self._b


class FakeSession:
    def __init__(self, script):
        self.headers, self.calls, self.script = {}, [], script          # script : {(méthode, fin_de_chemin): [réponses...]}

    def _answer(self, method, url, **kw):
        self.calls.append((method, url, kw))
        for (m, tail), queue in self.script.items():
            if m == method and url.split("?")[0].endswith(tail) and queue:
                return queue.pop(0) if len(queue) > 1 else queue[0]
        return Resp(404, text="non prévu")

    def request(self, method, url, **kw):
        return self._answer(method, url, **kw)

    def get(self, url, **kw):
        return self._answer("GET", url, **kw)


def client(script):
    s = FakeSession(script)
    return KoboClient("https://kf.kobotoolbox.org/", "aTEST", "secret", session=s), s


ROWS = [{"name": "BSBAF001", "label": "x", "ref_titre": "t", "ref_marche": "m", "ref_suivi": "s"}]
VIEUX = {"uid": "afOLD", "file_type": "form_media", "metadata": {"filename": CSV_NAME}}


def test_sync_premiere_fois(tmp_path):
    c, s = client({("GET", "/files/"): [Resp(200, {"results": []})], ("POST", "/files/"): [Resp(201)],
                   ("GET", "/aTEST/"): [Resp(200, {"version_id": "v42"})], ("PATCH", "/deployment/"): [Resp(200)]})
    res = sync_to_kobo(c, ROWS)
    assert res == {"activites": 1, "fichier": "ajouté", "redeploye": True, "version": "v42"}
    post = next(k for m, u, k in s.calls if m == "POST")
    assert post["files"]["content"][0] == CSV_NAME and post["files"]["content"][1].startswith(b"name,label")
    assert '"filename": "activites.csv"' in post["data"]["metadata"] and post["data"]["file_type"] == "form_media"
    assert s.headers["Authorization"] == "Token secret" and s.calls[-1][2]["json"] == {"active": True, "version_id": "v42"}


def test_sync_remplace_l_ancien_fichier_apres_succes_de_l_envoi():
    c, s = client({("GET", "/files/"): [Resp(200, {"results": [VIEUX]})], ("POST", "/files/"): [Resp(201)], ("DELETE", "/afOLD/"): [Resp(204)],
                   ("GET", "/aTEST/"): [Resp(200, {"version_id": "v1"})], ("PATCH", "/deployment/"): [Resp(200)]})
    assert sync_to_kobo(c, ROWS)["fichier"] == "remplacé"
    ordre = [m for m, _, _ in s.calls if m in ("POST", "DELETE")]
    assert ordre == ["POST", "DELETE"]                           # jamais de période sans liste côté Kobo


def test_sync_doublon_refuse_par_kobo_puis_suppression_et_nouvel_envoi():
    c, s = client({("GET", "/files/"): [Resp(200, {"results": [VIEUX]})], ("POST", "/files/"): [Resp(400, text="doublon"), Resp(201)],
                   ("DELETE", "/afOLD/"): [Resp(204)], ("GET", "/aTEST/"): [Resp(200, {"version_id": "v1"})], ("PATCH", "/deployment/"): [Resp(200)]})
    assert sync_to_kobo(c, ROWS)["fichier"] == "remplacé"
    assert [m for m, _, _ in s.calls if m in ("POST", "DELETE")] == ["POST", "DELETE", "POST"]


def test_sync_sans_redeploiement():
    c, s = client({("GET", "/files/"): [Resp(200, {"results": []})], ("POST", "/files/"): [Resp(201)]})
    assert sync_to_kobo(c, ROWS, redeploy=False)["redeploye"] is False and not [1 for m, _, _ in s.calls if m == "PATCH"]


def test_sync_jeton_invalide_message_clair():
    c, _ = client({("GET", "/files/"): [Resp(401, text="no")]})
    with pytest.raises(KoboError, match="jeton KOBO_TOKEN invalide"):
        sync_to_kobo(c, ROWS)


def test_sync_liste_vide_refusee_sans_appel_reseau():
    c, s = client({})
    with pytest.raises(KoboError, match="Aucune activité"):
        sync_to_kobo(c, [])
    assert s.calls == []


def test_sync_formulaire_non_deploye():
    c, _ = client({("GET", "/files/"): [Resp(200, {"results": []})], ("POST", "/files/"): [Resp(201)], ("GET", "/aTEST/"): [Resp(200, {"version_id": None})]})
    with pytest.raises(KoboError, match="déployez d'abord"):
        sync_to_kobo(c, ROWS)
