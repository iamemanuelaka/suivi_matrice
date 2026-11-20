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
from pcr.kobo_form import (CSV_COLUMNS, CSV_NAME, STRUCTURES_NAME, SURVEY, SURVEY_HEADER, TYPES_NAME, KoboClient, KoboError,
                           StructuresError, activities_csv_rows, build_xlsform, csv_bytes, load_structures, sync_to_kobo,
                           types_csv_rows, write_activities_csv)
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
    assert rows[0]["label"] == "Blocs sanitaires · Bafing · Koro — ETS NATHA SARL"          # pas de code : les agents ne le connaissent pas
    assert "118 356 245 FCFA" in rows[0]["ref_marche"] and "Fin prévue : 17/12/2026" in rows[0]["ref_marche"]


def test_activite_incomplete_proposee_avec_libelles_neutres(ref):
    ref.create_activity("MR", "BAF")
    r = activities_csv_rows(ref)[0]
    assert r["label"] == "Marchés ruraux · Bafing · — — titulaire non renseigné"
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
    assert "instance('activites')/root/item[" in xform                       # liste externe alimentée par le CSV (filtrée)
    assert "odk-instance-first-load" in xform and "today()" in xform


def test_colonnes_lues_par_pulldata_existent_dans_le_csv():
    fichier = CSV_NAME[:-4]
    lues = []
    for row in SURVEY:
        calc = row[SURVEY_HEADER.index("calculation")]
        if "pulldata" in calc:
            m = re.fullmatch(rf"pulldata\('{fichier}', '(\w+)', 'name', \$\{{code_activite\}}\)", calc)
            assert m, calc
            lues.append(m[1])
    assert lues and set(lues) <= set(CSV_COLUMNS)


import pyxform  # noqa: E402

VALIDATE_JAR = os.environ.get("ODK_VALIDATE_JAR") or str(Path(pyxform.__file__).parent / "validators" / "odk_validate" / "bin" / "ODK_Validate.jar")


@pytest.mark.skipif(not (shutil.which("java") and Path(VALIDATE_JAR).exists()), reason="ODK Validate (Java) non disponible")
def test_formulaire_valide_par_odk_validate(tmp_path):
    p = build_xlsform(tmp_path / "q.xlsx")
    out = tmp_path / "f.xml"
    out.write_text(convert(xlsform=str(p), warnings=[]).xform, encoding="utf-8")
    r = subprocess.run(["java", "-jar", VALIDATE_JAR, str(out)], capture_output=True, text=True)
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
    assert res == {"activites": 1, "fichier": "ajouté", "redeploye": True, "version": "v42", "structures": 0, "fichier_structures": None,
                   "types": 0}
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


# ---------------- structures et choix en cascade ----------------
def test_les_11_utr_couvrent_chacune_leur_region(ref):
    st = load_structures(ref.cfg)
    assert len(st) == 11 and {x["regions"] for x in st} == set(ref.cfg.regions)
    assert next(x for x in st if x["name"] == "UTR_BER")["label"] == "UTR du Béré"


@pytest.fixture
def ref_regions(ref):
    """Blocs sanitaires et marchés ruraux dans plusieurs régions."""
    for cat, reg in (("BS", "BER"), ("BS", "BER"), ("MR", "BER"), ("BS", "POR"), ("MR", "POR"), ("BS", "BAF"), ("MR", "TCH")):
        ref.create_activity(cat, reg, Departement="Test")
    return ref


def test_cascade_une_utr_ne_voit_que_les_activites_de_sa_region(ref_regions):
    ref = ref_regions
    rows = activities_csv_rows(ref)
    assert len(rows) == 7
    assert all(r["structures"] for r in rows)
    def proposees(structure):                       # ce que fait le choice_filter selected(structures, ${structure})
        return [r["name"] for r in rows if structure in r["structures"].split()]
    ber = proposees("UTR_BER")
    assert ber and all(c[2:5] == "BER" for c in ber)
    assert {c[:2] for c in ber} == {"BS", "MR"}                                   # blocs sanitaires ET marchés ruraux
    assert sum(len(proposees(x["name"])) for x in load_structures(ref.cfg)) == len(rows)


def _structures(tmp_path, ref, contenu):
    f = tmp_path / "structures.csv"
    f.write_text(contenu, encoding="utf-8")
    ref.cfg.data["chemins"]["kobo_structures"] = str(f)
    return f


def test_structure_ajoutee_au_csv_avec_plusieurs_regions_ou_toutes(ref_regions, tmp_path):
    ref = ref_regions
    base = (Path(__file__).resolve().parent.parent / "kobo" / "structures.csv").read_text(encoding="utf-8")
    _structures(tmp_path, ref, base + "DR_NORD,Direction régionale Nord,POR TCH BAG\nCOORD,Coordination PCR-CI,*\n")
    rows = {r["name"]: r["structures"].split() for r in activities_csv_rows(ref)}
    assert all("COORD" in v for v in rows.values())
    por = next(k for k in rows if k[2:5] == "POR")
    assert set(rows[por]) == {"UTR_POR", "DR_NORD", "COORD"}
    assert "DR_NORD" not in rows[next(k for k in rows if k[2:5] == "BER")]


@pytest.mark.parametrize("contenu, message", [
    ("name,label,regions\nUTR BER,UTR du Béré,BER\n", "invalide"),
    ("name,label,regions\nUTR_BER,UTR du Béré,XYZ\n", "XYZ"),
    ("name,label,regions\nUTR_BER,UTR du Béré,BER\nUTR_BER,Doublon,BER\n", "double"),
])
def test_liste_des_structures_controlee(ref, tmp_path, contenu, message):
    _structures(tmp_path, ref, contenu)
    with pytest.raises(StructuresError, match=message):
        load_structures(ref.cfg)


def test_activite_sans_structure_refusee_avant_envoi(ref_regions, tmp_path):
    ref = ref_regions
    _structures(tmp_path, ref, "name,label,regions\nUTR_BER,UTR du Béré,BER\n")
    with pytest.raises(StructuresError, match="Aucune structure ne couvre"):
        activities_csv_rows(ref)


def test_formulaire_structure_en_liste_deroulante_et_activite_filtree(xform):
    assert "instance('structures')/root/item" in xform
    assert re.search(r'ref="/data/identification/structure"[^>]*appearance="minimal"', xform)
    assert re.search(r"instance\('types_activite'\)/root/item\[selected\(\s*structures\s*,\s*/data/identification/structure\s*\)\]", xform)
    assert re.search(r"instance\('activites'\)/root/item\[selected\(\s*structures\s*,\s*/data/identification/structure\s*\)\s*and\s*type\s*=\s*"
                     r"/data/identification/type_activite\s*\]", xform)
    assert 'orx:max-pixels="1600"' in xform or 'max-pixels="1600"' in xform


def test_fichiers_medias_ecrits_ensemble(ref, tmp_path):
    write_activities_csv(ref, tmp_path / "activites.csv")
    act = (tmp_path / "activites.csv").read_text(encoding="utf-8").splitlines()
    st = (tmp_path / STRUCTURES_NAME).read_text(encoding="utf-8").splitlines()
    assert act[0].endswith(",structures,type") and st[0] == "name,label,regions" and len(st) == 12
    assert (tmp_path / "types_activite.csv").read_text(encoding="utf-8").splitlines()[0] == "name,label,structures"


def test_sync_envoie_les_structures_puis_les_activites():
    c, s = client({("GET", "/files/"): [Resp(200, {"results": []})], ("POST", "/files/"): [Resp(201)],
                   ("GET", "/aTEST/"): [Resp(200, {"version_id": "v42"})], ("PATCH", "/deployment/"): [Resp(200)]})
    res = sync_to_kobo(c, ROWS, structures=[{"name": "UTR_BAF", "label": "UTR du Bafing", "regions": "BAF"}])
    envoyes = [k["files"]["content"][0] for m, u, k in s.calls if m == "POST"]
    assert envoyes == [STRUCTURES_NAME, CSV_NAME] and res["structures"] == 1


def test_la_structure_choisie_est_enregistree_avec_le_suivi(ref_regions):
    from pcr.ingest import evaluate
    ref = ref_regions
    acts = {a["CODE_ACTIVITE"]: a for a in ref.activities()}
    sub = {"_id": 5, "identification/structure": "UTR_BER", "identification/code_activite": "BSBER001",
           "identification/date_actualisation": "2026-11-18", "avancement/taux_execution": "20", "avancement/observations_techniques": "Fouilles"}
    rec, _, _ = evaluate(sub, acts, [], date(2026, 11, 20))
    assert rec["Structure"] == "UTR_BER"
    ancien = {**sub, "identification/Structure": "UTR Béré (texte libre)"}
    del ancien["identification/structure"]
    assert evaluate(ancien, acts, [], date(2026, 11, 20))[0]["Structure"] == "UTR Béré (texte libre)"     # ancien champ texte


def test_les_agents_ne_voient_aucun_code_et_les_libelles_sont_distincts(ref_regions):
    ref = ref_regions
    a = ref.create_activity("BS", "BER", Departement="Mankono", Entreprise="ARI", Localite="DIALAKORO, DANGBASSO, LINGUEKORO")
    b = ref.create_activity("BS", "BER", Departement="Mankono", Entreprise="ARI", Localite="MISSIDOUGOU, KOUASSIDOUGOU")
    c = ref.create_activity("MR", "BER", Departement="Mankono", Entreprise="ARI")
    d = ref.create_activity("MR", "BER", Departement="Mankono", Entreprise="ARI")
    rows = {r["name"]: r for r in activities_csv_rows(ref)}
    visibles = " ".join(r[k] for r in rows.values() for k in ("label", "ref_titre", "ref_marche", "ref_suivi"))
    assert not re.search(r"\b(BS|MR)[A-Z]{3}\d{3}\b", visibles)
    assert len({r["label"] for r in rows.values()}) == len(rows)
    assert rows[a]["label"] == "Blocs sanitaires · Béré · Mankono — ARI · Dialakoro / Dangbasso…"
    assert rows[b]["label"] == "Blocs sanitaires · Béré · Mankono — ARI · Missidougou / Kouassidougou"
    assert not any("," in r["label"] for r in rows.values())                                    # aucune virgule à protéger dans le CSV
    assert rows[c]["label"].endswith("· lot 1") and rows[d]["label"].endswith("· lot 2")       # ni localité ni montant : numéro de lot


def test_activite_apres_la_structure_et_coherente_avec_elle(xform):
    bind = re.search(r'<bind nodeset="/data/identification/code_activite"[^>]*>', xform)[0]
    assert "relevant=\" /data/identification/structure  != ''\"" in bind or "/data/identification/structure" in bind
    assert "selected(pulldata('activites', 'structures', 'name', .)" in bind and "constraintMsg" in bind
    assert "pulldata('activites', 'type', 'name', .) =  /data/identification/type_activite" in bind
    assert "/data/identification/type_activite" in re.search(r'<bind nodeset="/data/identification/code_activite"[^>]*relevant="([^"]*)"', xform)[1]


def test_cascade_a_trois_niveaux_structure_type_activite(ref_regions, tmp_path):
    ref = ref_regions
    base = (Path(__file__).resolve().parent.parent / "kobo" / "structures.csv").read_text(encoding="utf-8")
    _structures(tmp_path, ref, base)
    st = load_structures(ref.cfg)
    rows = activities_csv_rows(ref, st)
    types = {t["name"]: t for t in types_csv_rows(ref, rows, st)}
    def types_de(structure):                                      # choice_filter des types
        return sorted(k for k, t in types.items() if structure in t["structures"].split())
    def activites_de(structure, typ):                             # choice_filter des activités
        return sorted(r["name"] for r in rows if structure in r["structures"].split() and r["type"] == typ)
    assert types["BS"]["label"] == "Blocs sanitaires" and types["MR"]["label"] == "Marchés ruraux"
    assert types_de("UTR_BER") == ["BS", "MR"] and types_de("UTR_BAF") == ["BS"] and types_de("UTR_TCH") == ["MR"]
    assert all(c[:2] == "BS" and c[2:5] == "BER" for c in activites_de("UTR_BER", "BS")) and len(activites_de("UTR_BER", "BS")) == 2
    assert activites_de("UTR_BER", "MR") == [r["name"] for r in rows if r["name"].startswith("MRBER")]
    assert types_de("UTR_GON") == []                               # aucune activité au Gontougo ici : aucun type proposé


def test_quatre_questions_photo_aux_noms_du_formulaire_deploye(xform):
    for nom in ("Photo_du_site", "Photo_du_site_001", "Photo_du_site_002", "Photo_du_site_003"):
        assert f'ref="/data/avancement/{nom}"' in xform


def test_sync_envoie_aussi_les_types():
    c, s = client({("GET", "/files/"): [Resp(200, {"results": []})], ("POST", "/files/"): [Resp(201)],
                   ("GET", "/aTEST/"): [Resp(200, {"version_id": "v42"})], ("PATCH", "/deployment/"): [Resp(200)]})
    res = sync_to_kobo(c, ROWS, structures=[{"name": "UTR_BAF", "label": "UTR du Bafing", "regions": "BAF"}],
                       types=[{"name": "BS", "label": "Blocs sanitaires", "structures": "UTR_BAF"}])
    envoyes = [k["files"]["content"][0] for m, u, k in s.calls if m == "POST"]
    assert envoyes == [STRUCTURES_NAME, TYPES_NAME, CSV_NAME] and res["types"] == 1
