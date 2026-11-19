"""Téléchargement des photos jointes aux soumissions Kobo vers data/photos/.

Utilise l'API Kobo v2. Auth par KOBO_TOKEN (recommandé) ou KOBO_USER/KOBO_PASSWORD.
Les photos déjà présentes ne sont pas retéléchargées. Les erreurs n'interrompent pas le reste.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

PHOTOS_DIR = Path("data/photos")


def _session_with_retry() -> requests.Session:
    s = requests.Session()
    retries = Retry(total=4, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
                    allowed_methods=("GET",), respect_retry_after_header=True)
    s.mount("https://", HTTPAdapter(max_retries=retries))
    return s


def download_photos(base_url: str | None = None, asset_uid: str | None = None,
                    dest: Path | None = None) -> dict:
    """Télécharge les pièces jointes des soumissions vers dest (défaut data/photos/)."""
    base_url = (base_url or os.environ.get("KOBO_URL") or "https://kf.kobotoolbox.org").rstrip("/")
    asset_uid = asset_uid or os.environ.get("KOBO_ASSET_UID")
    if not asset_uid:
        raise RuntimeError("KOBO_ASSET_UID non défini.")

    token = os.environ.get("KOBO_TOKEN")
    auth = None
    headers = {}
    if token:
        headers["Authorization"] = f"Token {token}"
    else:
        user, password = os.environ.get("KOBO_USER"), os.environ.get("KOBO_PASSWORD")
        if user and password:
            auth = (user, password)
        else:
            raise RuntimeError("Définissez KOBO_TOKEN ou KOBO_USER/KOBO_PASSWORD.")

    dest = Path(dest) if dest else PHOTOS_DIR
    dest.mkdir(parents=True, exist_ok=True)

    session = _session_with_retry()
    session.headers.update(headers)

    url = f"{base_url}/api/v2/assets/{asset_uid}/data/?format=json&limit=1000"
    telechargees = deja_presentes = total = 0
    erreurs = []

    while url:
        r = session.get(url, timeout=60, auth=auth)
        if r.status_code != 200:
            raise RuntimeError(f"Échec récupération soumissions (HTTP {r.status_code})")
        payload = r.json()
        for sub in payload.get("results", []):
            for att in (sub.get("_attachments") or []):
                total += 1
                filename = att.get("filename") or (att.get("download_url") or "").rstrip("/").split("/")[-1]
                if not filename:
                    continue
                filename = Path(filename).name
                local_path = dest / filename
                if local_path.is_file():
                    deja_presentes += 1
                    continue
                download_url = att.get("download_url") or att.get("download_large_url")
                if not download_url:
                    erreurs.append(f"{filename} : pas d'URL de téléchargement")
                    continue
                try:
                    resp = session.get(download_url, timeout=120, auth=auth, stream=True)
                    if resp.status_code != 200:
                        erreurs.append(f"{filename} (HTTP {resp.status_code})")
                        continue
                    with open(local_path, "wb") as f:
                        for chunk in resp.iter_content(8192):
                            if chunk:
                                f.write(chunk)
                    telechargees += 1
                except Exception as e:
                    erreurs.append(f"{filename} : {e}")
                    if local_path.exists():
                        local_path.unlink(missing_ok=True)
        url = payload.get("next")

    return {"telechargees": telechargees, "deja_presentes": deja_presentes,
            "erreurs": erreurs, "total_attachments": total, "dossier": str(dest.resolve())}


def main(argv=None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description="Télécharge les photos Kobo vers data/photos/.")
    parser.add_argument("--url", help="URL de base Kobo")
    parser.add_argument("--uid", help="UID de l'asset")
    parser.add_argument("--dest", type=Path, help="Dossier de destination")
    args = parser.parse_args(argv)
    try:
        resume = download_photos(base_url=args.url, asset_uid=args.uid, dest=args.dest)
    except Exception as e:
        print(f"Erreur : {e}", file=sys.stderr)
        return 1
    print(f"Dossier : {resume['dossier']}")
    print(f"Pièces jointes vues : {resume['total_attachments']}")
    print(f"Téléchargées       : {resume['telechargees']}")
    print(f"Déjà présentes     : {resume['deja_presentes']}")
    if resume["erreurs"]:
        print(f"Erreurs ({len(resume['erreurs'])}) :")
        for err in resume["erreurs"][:20]:
            print(f"  - {err}")
    return 0 if not resume["erreurs"] else 2


if __name__ == "__main__":
    sys.exit(main())
