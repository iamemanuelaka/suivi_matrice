"""Téléchargement des photos du site (question « Photo_du_site ») vers data/photos/.

Même principe que suivi_activites_test :
  - la pièce jointe est retrouvée dans « _attachments » par sa question (question_xpath)
    ou par son nom (media_file_basename) ;
  - l'image est enregistrée en JPEG réduit, orientation corrigée (EXIF), sous le nom
    « <nom de la photo Kobo sans extension>.jpg » ;
  - le dossier data/photos/ est versionné : les photos arrivent dans le dépôt git.

Auth par KOBO_TOKEN (recommandé) ou KOBO_USER/KOBO_PASSWORD. Les photos déjà présentes ne
sont pas retéléchargées. Une photo en échec n'interrompt jamais le reste.
"""
from __future__ import annotations

import io
import os
import sys
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

PHOTOS_DIR = Path("data/photos")
CHAMP_PHOTO = "Photo_du_site"
TAILLE_MAX = (1600, 1600)            # pixels : suffisant pour le Word, léger pour git


def local_name(kobo_value: str) -> str:
    """Nom du fichier local pour la valeur Kobo du champ photo (toujours en .jpg)."""
    return Path(str(kobo_value)).stem + ".jpg"


def photo_attachment(sub: dict, field: str = CHAMP_PHOTO) -> tuple[str, dict] | None:
    """(valeur du champ photo, pièce jointe correspondante) ou None."""
    value = next((v for k, v in sub.items() if k.split("/")[-1] == field and v), None)
    if not value:
        return None
    value = Path(str(value)).name
    for att in sub.get("_attachments") or []:
        if att.get("is_deleted"):
            continue
        xpath = str(att.get("question_xpath") or "").split("/")[-1]
        base = str(att.get("media_file_basename") or str(att.get("filename") or "").split("/")[-1])
        if xpath == field or base == value:
            return value, att
    return None


def to_jpeg(data: bytes) -> bytes:
    from PIL import Image, ImageOps
    try:                                                    # HEIC/HEIF si pillow-heif est installé
        from pillow_heif import register_heif_opener
        register_heif_opener()
    except Exception:                                       # noqa: BLE001
        pass
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail(TAILLE_MAX)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=80, optimize=True)
    return buf.getvalue()


def _session() -> requests.Session:
    s = requests.Session()
    s.mount("https://", HTTPAdapter(max_retries=Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                                                      allowed_methods=("GET",), respect_retry_after_header=True)))
    token = os.environ.get("KOBO_TOKEN")
    if token:
        s.headers["Authorization"] = f"Token {token}"
    elif os.environ.get("KOBO_USER") and os.environ.get("KOBO_PASSWORD"):
        s.auth = (os.environ["KOBO_USER"], os.environ["KOBO_PASSWORD"])
    else:
        raise RuntimeError("Définissez KOBO_TOKEN ou KOBO_USER/KOBO_PASSWORD.")
    return s


def fetch_submissions(session: requests.Session, base_url: str, asset_uid: str) -> list[dict]:
    url, out = f"{base_url.rstrip('/')}/api/v2/assets/{asset_uid}/data/?format=json&limit=1000", []
    while url:
        r = session.get(url, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Échec récupération soumissions (HTTP {r.status_code})")
        payload = r.json()
        out += payload.get("results", [])
        url = payload.get("next")
    return out


def download_photos(submissions: list[dict] | None = None, session: requests.Session | None = None,
                    base_url: str | None = None, asset_uid: str | None = None, dest: Path | None = None) -> dict:
    """Télécharge la photo du site de chaque soumission. Retourne un résumé."""
    session = session or _session()
    if submissions is None:
        base_url = base_url or os.environ.get("KOBO_URL") or "https://kf.kobotoolbox.org"
        asset_uid = asset_uid or os.environ.get("KOBO_ASSET_UID")
        if not asset_uid:
            raise RuntimeError("KOBO_ASSET_UID non défini.")
        submissions = fetch_submissions(session, base_url, asset_uid)
    dest = Path(dest) if dest else PHOTOS_DIR
    dest.mkdir(parents=True, exist_ok=True)

    res = {"telechargees": 0, "deja_presentes": 0, "sans_photo": 0, "erreurs": [], "dossier": str(dest.resolve())}
    for sub in submissions:
        found = photo_attachment(sub)
        if not found:
            res["sans_photo"] += 1
            continue
        value, att = found
        path = dest / local_name(value)
        if path.is_file():
            res["deja_presentes"] += 1
            continue
        data, last = None, "pas d'URL de téléchargement"
        for key in ("download_large_url", "download_medium_url", "download_url"):
            url = att.get(key)
            if not url:
                continue
            try:
                r = session.get(url, timeout=120)
                if r.status_code == 200 and r.content:
                    data = r.content
                    break
                last = f"HTTP {r.status_code}"
            except requests.RequestException as e:
                last = str(e)
        if data is None:
            res["erreurs"].append(f"{value} (soumission {sub.get('_id')}) : {last}")
            continue
        try:
            path.write_bytes(to_jpeg(data))
            res["telechargees"] += 1
        except Exception as e:                              # noqa: BLE001 : image illisible
            res["erreurs"].append(f"{value} (soumission {sub.get('_id')}) : image illisible ({e})")
    return res


def main(argv=None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description="Télécharge les photos Kobo vers data/photos/.")
    parser.add_argument("--url", help="URL de base Kobo")
    parser.add_argument("--uid", help="UID de l'asset")
    parser.add_argument("--dest", type=Path, help="Dossier de destination")
    args = parser.parse_args(argv)
    try:
        r = download_photos(base_url=args.url, asset_uid=args.uid, dest=args.dest)
    except Exception as e:                                  # noqa: BLE001
        print(f"Erreur : {e}", file=sys.stderr)
        return 1
    print(f"Dossier          : {r['dossier']}")
    print(f"Téléchargées     : {r['telechargees']}")
    print(f"Déjà présentes   : {r['deja_presentes']}")
    print(f"Sans photo       : {r['sans_photo']}")
    for err in r["erreurs"][:20]:
        print(f"  - {err}")
    return 0 if not r["erreurs"] else 2


if __name__ == "__main__":
    sys.exit(main())
