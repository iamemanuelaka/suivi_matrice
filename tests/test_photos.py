"""Photos Kobo : téléchargement (comme suivi_activites_test) puis insertion dans la colonne « Photo »."""
import io
import sys
from datetime import date
from pathlib import Path

import docx
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pcr.kobo_photos import download_photos, local_name, photo_attachment


def _jpeg_tourne() -> bytes:
    """Photo de téléphone « couchée » : 400x300 avec EXIF orientation=6 (à tourner de 90°)."""
    img, exif = Image.new("RGB", (400, 300), (200, 50, 50)), Image.Exif()
    exif[0x0112] = 6
    buf = io.BytesIO()
    img.save(buf, "JPEG", exif=exif)
    return buf.getvalue()


SUB = {"_id": 77, "avancement/Photo_du_site": "1763560000123.jpg",
       "_attachments": [{"question_xpath": "avancement/Photo_du_site", "media_file_basename": "1763560000123.jpg",
                         "filename": "user/attachments/abc/def/1763560000123.jpg",
                         "download_url": "https://kf.example/api/v2/assets/X/data/77/attachments/1/",
                         "download_large_url": "https://kf.example/api/v2/assets/X/data/77/attachments/1/large/"}]}


class FakeSession:
    def __init__(self):
        self.urls = []

    def get(self, url, timeout=None):
        self.urls.append(url)
        return type("R", (), {"status_code": 200, "content": _jpeg_tourne()})()


def test_piece_jointe_retrouvee_par_question():
    assert photo_attachment(SUB)[0] == "1763560000123.jpg"
    assert photo_attachment({"_id": 1, "avancement/Photo_du_site": None}) is None
    assert local_name("photo.heic") == "photo.jpg"


def test_telechargement_jpeg_oriente_et_non_retelecharge(tmp_path):
    s = FakeSession()
    r = download_photos([SUB, {"_id": 78}], session=s, dest=tmp_path)
    assert (r["telechargees"], r["sans_photo"], r["erreurs"]) == (1, 1, [])
    assert s.urls == [SUB["_attachments"][0]["download_large_url"]]
    assert Image.open(tmp_path / "1763560000123.jpg").size == (300, 400)          # orientation corrigée
    assert download_photos([SUB], session=s, dest=tmp_path)["deja_presentes"] == 1 and len(s.urls) == 1


# --- 4 photos par suivi ------------------------------------------------------------------------
def _sub4(n_photos=4):
    noms = ("Photo_du_site", "Photo_du_site_001", "Photo_du_site_002", "Photo_du_site_003")[:n_photos]
    sub = {"_id": 88, "_attachments": []}
    for i, nom in enumerate(noms):
        sub[f"avancement/{nom}"] = f"img{i}.jpg"
        sub["_attachments"].append({"question_xpath": f"avancement/{nom}", "media_file_basename": f"img{i}.jpg",
                                    "download_url": f"https://kf.example/a/{i}/"})
    return sub


def test_les_4_photos_sont_telechargees(tmp_path):
    s = FakeSession()
    r = download_photos([_sub4()], session=s, dest=tmp_path)
    assert r["telechargees"] == 4 and r["erreurs"] == [] and sorted(p.name for p in tmp_path.iterdir()) == [f"img{i}.jpg" for i in range(4)]
    assert photo_attachment(_sub4(), "Photo_du_site_002")[0] == "img2.jpg"                  # chaque question retrouve SA pièce jointe


def test_les_4_photos_dans_la_base_avec_rattrapage(tmp_path):
    import shutil
    from pcr.config import load_config
    from pcr.ingest import ingest
    from pcr.referentiel import Referentiel
    from pcr.store import Base
    shutil.copy(Path(__file__).parent / "app_ui" / "base_reference.xlsx", tmp_path / "b.xlsx")
    ref = Referentiel(Base(tmp_path / "b.xlsx"), load_config())
    sub = {"_id": 777, "_submission_time": "2026-11-20T09:00:00", "identification/code_activite": "BSBAF001",
           "identification/date_actualisation": "2026-11-20", "avancement/taux_execution": "70",
           "avancement/observations_techniques": "Dallage", "avancement/Photo_du_site": "a.jpg"}
    ingest(ref, [sub], date(2026, 11, 20))
    ligne = lambda: next(r for r in ref.base.rows("Suivi_Activites") if str(r["ID_SOUMISSION_KOBO"]) == "777")   # noqa: E731
    assert ligne()["Photo_du_site"] == "a.jpg" and ligne()["Photo_du_site_003"] is None
    ingest(ref, [{**sub, "avancement/Photo_du_site_001": "b.jpg", "avancement/Photo_du_site_003": "d.jpg"}], date(2026, 11, 20))
    assert (ligne()["Photo_du_site_001"], ligne()["Photo_du_site_003"], ligne()["Photo_du_site"]) == ("b.jpg", "d.jpg", "a.jpg")


def test_cellule_photo_divisee_en_4_dans_le_word(tmp_path, monkeypatch):
    import shutil
    from pcr.config import load_config
    from pcr.ingest import ingest
    from pcr.referentiel import Referentiel
    from pcr.store import Base
    from pcr.word_gen import render
    monkeypatch.chdir(tmp_path)
    (tmp_path / "data" / "photos").mkdir(parents=True)
    for i in range(4):
        Image.new("RGB", (1200, 1600) if i % 2 else (1600, 1200), (40 * i, 90, 160)).save(tmp_path / "data" / "photos" / f"p{i}.jpg")
    shutil.copy(Path(__file__).parent / "app_ui" / "base_reference.xlsx", tmp_path / "b.xlsx")
    ref = Referentiel(Base(tmp_path / "b.xlsx"), load_config())
    base = {"_submission_time": "2026-11-20T09:00:00", "identification/date_actualisation": "2026-11-20",
            "avancement/taux_execution": "70", "avancement/observations_techniques": "Dallage"}
    quatre = {f"avancement/{n}": f"p{i}.jpg" for i, n in enumerate(("Photo_du_site", "Photo_du_site_001", "Photo_du_site_002", "Photo_du_site_003"))}
    ingest(ref, [{**base, "_id": 1, "identification/code_activite": "BSBAF001", **quatre},
                 {**base, "_id": 2, "identification/code_activite": "BSBAF002", "avancement/Photo_du_site": "p1.jpg"}], date(2026, 11, 20))
    doc = docx.Document(render(ref, "BS", date(2026, 11, 20), tmp_path / "BS.docx"))
    t = doc.tables[0]
    c1, c2 = t.rows[1].cells[3], t.rows[2].cells[3]
    assert len(c1.tables) == 1 and len(c1.tables[0].rows) == 2 and len(c1.tables[0].columns) == 2   # 2 × 2
    images = lambda cell: cell._tc.xml.count("<pic:pic")                                            # noqa: E731
    assert images(c1) == 4 and images(c2) == 1 and not c2.tables                                    # une seule photo : pas de grille
    formes = [doc.inline_shapes[i] for i in range(len(doc.inline_shapes))]
    assert len(formes) == 5 and all(f.height.cm <= 2.61 and f.width.cm <= 3.11 for f in formes[:4])   # chaque photo tient dans sa case
    assert formes[4].height.cm <= 4.81                                                               # photo seule : 4,8 cm au plus
