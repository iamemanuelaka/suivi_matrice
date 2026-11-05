"""Chargement de la configuration et résolution des noms (accents/casse ignorés)."""
import re
import unicodedata
from pathlib import Path

import yaml

DEFAULT_CONFIG = Path(__file__).resolve().parent.parent / "config" / "config.yaml"


def normalize(text) -> str:
    """MAJUSCULES, sans accents, sans ponctuation : sert aux comparaisons de noms."""
    s = unicodedata.normalize("NFD", str(text or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^A-Z0-9]+", " ", s.upper()).strip()


class Config:
    def __init__(self, data: dict, root: Path):
        self.data, self.root = data, root
        self.categories = {k: v["nom"] for k, v in data["categories"].items()}
        self.conventions = {k: v.get("convention_delai", "calendaire") for k, v in data["categories"].items()}
        self.regions = dict(data["regions"])
        self._cat_idx = {normalize(k): k for k in self.categories} | {normalize(v): k for k, v in self.categories.items()}
        self._reg_idx = {normalize(k): k for k in self.regions} | {normalize(v): k for k, v in self.regions.items()}

    @property
    def admin(self) -> str:
        return self.data.get("administrateur", "Administrateur")

    @property
    def tolerance_jours(self) -> int:
        return int(self.data.get("controles", {}).get("tolerance_jours_fin_previsionnelle", 5))

    def category(self, cat: str) -> dict:
        return self.data["categories"][cat]

    @property
    def date_etat_initial(self):
        from datetime import date
        return date.fromisoformat(str(self.data.get("controles", {}).get("date_etat_initial", "2026-10-25")))

    def path(self, key: str) -> Path:
        return self.root / self.data["chemins"][key]

    def category_code(self, value) -> str:
        code = self._cat_idx.get(normalize(value))
        if not code:
            raise ValueError(f"Catégorie inconnue : {value!r} (connues : {', '.join(self.categories)})")
        return code

    def region_code(self, value) -> str:
        code = self._reg_idx.get(normalize(value))
        if not code:
            raise ValueError(f"Région inconnue : {value!r} (connues : {', '.join(self.regions)})")
        return code


def load_config(path=None) -> Config:
    p = Path(path) if path else DEFAULT_CONFIG
    return Config(yaml.safe_load(p.read_text(encoding="utf-8")), root=p.resolve().parent.parent)
