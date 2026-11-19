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
