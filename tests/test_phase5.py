import sys
from datetime import date, datetime
from pathlib import Path

import docx
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr.baseline import import_baseline
from pcr.config import load_config
from pcr.ingest import ingest
from pcr.referentiel import Referentiel
from pcr.seed import seed
from pcr.store import Base
from pcr.word_check import compare, generated_rows, norm
from pcr.word_gen import generate_all, render

SRC = Path(__file__).resolve().parent.parent / "sources"
SOURCES = [("BS", SRC / "matrice_BS_texte.md"), ("MR", SRC / "matrice_MR_texte.md")]
AU = date(2026, 10, 25)


@pytest.fixture(scope="module")
def full(tmp_path_factory):
    d = tmp_path_factory.mktemp("w")
    ref = Referentiel(Base(d / "b.xlsx"), load_config())
    seed(ref, SOURCES)
    import_baseline(ref, SOURCES, AU)
    return ref, generate_all(ref, AU, d), d


def test_deux_matrices_une_par_categorie(full):
    _, files, _ = full
    assert [f.name for f in files] == ["MATRICE_DE_SUIVI_DES_BLOCS_SANITAIRES_25102026.docx", "MATRICE_DE_SUIVI_DES_MARCHES_RURAUX_25102026.docx"]


def test_structure_tableau_a_trois_colonnes_et_en_tetes(full):
    _, (bs, mr), _ = full
    t = docx.Document(bs).tables[0]
    assert len(t.rows) == 27 and len(t.columns) == 3
    assert [c.text for c in t.rows[0].cells] == ["Désignation Activité", "Avancement physique au 25/10/2026", "Commentaires"]
    assert len(docx.Document(mr).tables[0].rows) == 1 + 16 * 3                 # 3 lignes par marché rural


def test_observations_en_tirets_au_dessus_du_taux(full):
    _, (bs, _), _ = full
    mid = docx.Document(bs).tables[0].rows[1].cells[1]
    textes = [p.text for p in mid.paragraphs]
    assert textes[0].startswith("Visite effectuée") and textes[-3].startswith("Taux d’avancement") and textes[-2].startswith("Délai consommé")
    assert mid.paragraphs[0]._p.pPr.numPr is not None                           # vrai tiret de liste Word
    assert mid.paragraphs[-3]._p.pPr is None or mid.paragraphs[-3]._p.pPr.numPr is None


def test_fidelite_au_word_d_origine(full):
    ref, files, _ = full
    rep = compare(ref, {"BS": SOURCES[0][1], "MR": SOURCES[1][1]}, {"BS": files[0], "MR": files[1]})
    assert "| Avancement | 26 / 26 |" in rep and "| Commentaires | 16 / 16 |" in rep
    ecarts = [l for l in rep.split("## Différences de contenu restantes")[1].split("## Délai")[0].splitlines() if l.startswith("- ")]
    assert len(ecarts) <= 6                                                    # écarts connus et expliqués, aucun autre


def test_une_nouvelle_soumission_met_a_jour_observations_et_taux_separement(full, tmp_path):
    ref, _, d = full
    code = "BSBAF001"
    ingest(ref, [{"_id": 5001, "_submission_time": "2026-10-30T09:00:00", "identification/nom_agent": "Kouadio", "identification/date_actualisation": "2026-10-30",
                  "identification/code_activite": code, "avancement/observations_techniques": "Fondation achevée\nÉlévation en cours",
                  "avancement/taux_execution": "35", "remarques/recommandations": "Accélérer la cadence"}], date(2026, 11, 4))
    p = render(ref, "BS", date(2026, 11, 4), tmp_path / "x.docx")
    mid, right = docx.Document(p).tables[0].rows[1].cells[1], docx.Document(p).tables[0].rows[1].cells[2]
    texts = [x.text for x in mid.paragraphs]
    assert texts[:2] == ["Fondation achevée", "Élévation en cours"] and "Taux d’avancement : 35%" in texts
    assert "Difficultés : RAS" in [x.text for x in right.paragraphs]            # champ vide côté technicien -> « RAS »
    assert "Recommandations : Accélérer la cadence" in [x.text for x in right.paragraphs]
    assert norm("Taux d’avancement : 5 %") == norm("Taux d’avancement : 5%")


def test_suivi_rejete_ou_a_verifier_n_apparait_pas(full, tmp_path):
    ref, _, _ = full
    ingest(ref, [{"_id": 5002, "_submission_time": "2026-10-31T09:00:00", "identification/date_actualisation": "2026-10-31",
                  "identification/code_activite": "BSBAF002", "avancement/observations_techniques": "TEST REJET", "avancement/taux_execution": "999"}], date(2026, 11, 4))
    p = render(ref, "BS", date(2026, 11, 4), tmp_path / "y.docx")
    assert "TEST REJET" not in "\n".join(c.text for r in docx.Document(p).tables[0].rows for c in r.cells)
