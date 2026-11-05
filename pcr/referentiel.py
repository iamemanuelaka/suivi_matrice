"""Référentiel des activités : génération centralisée des codes, informations fixes, traçabilité.

Règles garanties ici (et testées) :
  - un CODE_ACTIVITE est unique, permanent, jamais réutilisé (aucune activité n'est supprimée, seulement archivée) ;
  - la séquence est propre à chaque couple catégorie + région et ne redescend jamais ;
  - une information vide n'est JAMAIS une demande de suppression (pour effacer : action explicite `clear`) ;
  - toute modification d'une information fixe est historisée (ancienne valeur, nouvelle valeur, date, utilisateur).
"""
import re
from datetime import date, datetime

from .calculs import fin_previsionnelle
from .config import Config
from .store import Base

FIXED = ["Departement", "Localite", "Designation", "Montant_Marche", "Entreprise", "Numero_OS", "Date_OS", "Date_Debut", "Delai_Mois", "Bureau_Controle"]
EDITABLE = FIXED + ["Statut"]
DATE_FIELDS = {"Date_OS", "Date_Debut"}
INT_FIELDS = {"Montant_Marche", "Delai_Mois"}
STATUTS = ("Actif", "Archivé")
CODE_RE = re.compile(r"^([A-Z]{2})([A-Z]{3})(\d{3})$")
MAX_SEQ = 999


def to_date(raw) -> date:
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    s = str(raw).strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    try:
        if m:
            return date(int(m[3]), int(m[2]), int(m[1]))
        return date.fromisoformat(s[:10])
    except ValueError:
        raise ValueError(f"date invalide : {raw!r} (format attendu JJ/MM/AAAA)") from None


def to_int(raw) -> int:
    if isinstance(raw, bool):
        raise ValueError(f"nombre invalide : {raw!r}")
    if isinstance(raw, (int, float)):
        if raw != int(raw):
            raise ValueError(f"nombre entier attendu : {raw!r}")
        return int(raw)
    digits = re.sub(r"[\s\u00a0]", "", str(raw))
    digits = re.sub(r"(?i)f?cfa|ht|ttc", "", digits)
    if not re.fullmatch(r"\d+", digits):
        raise ValueError(f"nombre invalide : {raw!r}")
    return int(digits)


def coerce(field: str, raw):
    """Convertit une valeur saisie ; None si vide. Lève ValueError si invalide."""
    if raw is None or (isinstance(raw, str) and not raw.strip()):
        return None
    if field in DATE_FIELDS:
        return to_date(raw)
    if field in INT_FIELDS:
        v = to_int(raw)
        if field == "Delai_Mois" and v < 1:
            raise ValueError("le délai doit être d'au moins 1 mois")
        return v
    s = str(raw).strip()
    if field == "Statut" and s not in STATUTS:
        raise ValueError(f"statut invalide : {s!r} (valeurs : {', '.join(STATUTS)})")
    return s


def fmt(v) -> str:
    """Forme texte stable d'une valeur pour l'historique."""
    if v is None:
        return ""
    return v.isoformat() if isinstance(v, (date, datetime)) else str(v)


class Referentiel:
    def __init__(self, base: Base, cfg: Config):
        self.base, self.cfg = base, cfg

    # --- lecture -----------------------------------------------------------
    def activities(self) -> list[dict]:
        return self.base.rows("Activites")

    def get(self, code: str) -> dict:
        for a in self.activities():
            if a["CODE_ACTIVITE"] == code:
                return a
        raise KeyError(f"Activité introuvable : {code}")

    def history(self, code: str | None = None) -> list[dict]:
        rows = self.base.rows("Historique_Modifications")
        return [r for r in rows if code is None or r["CODE_ACTIVITE"] == code]

    # --- codes -------------------------------------------------------------
    def _next_code(self, cat: str, reg: str) -> str:
        """Code suivant du couple (cat, reg). Source unique : max(compteur, plus grand code existant) + 1."""
        counters = self.base.rows("Compteurs_Codes")
        counter = next((c for c in counters if c["Categorie"] == cat and c["Region"] == reg), None)
        existing = [int(m[3]) for a in self.activities() if (m := CODE_RE.match(a["CODE_ACTIVITE"] or "")) and m[1] == cat and m[2] == reg]
        seq = max([counter["Dernier_Numero"] if counter else 0] + existing) + 1
        if seq > MAX_SEQ:
            raise ValueError(f"Plus de numéro disponible pour {cat}{reg} (maximum {MAX_SEQ})")
        if counter:
            self.base.update("Compteurs_Codes", counter["_r"], {"Dernier_Numero": seq})
        else:
            self.base.append("Compteurs_Codes", {"Categorie": cat, "Region": reg, "Dernier_Numero": seq})
        return f"{cat}{reg}{seq:03d}"

    # --- historique --------------------------------------------------------
    def _log(self, code, champ, old, new, user, when):
        rows = self.base.rows("Historique_Modifications")
        next_id = max([r["ID_MODIFICATION"] or 0 for r in rows] + [0]) + 1
        self.base.append("Historique_Modifications", {
            "ID_MODIFICATION": next_id, "CODE_ACTIVITE": code, "Champ_Modifie": champ,
            "Ancienne_Valeur": fmt(old) or None, "Nouvelle_Valeur": fmt(new) or None, "Date_Modification": when, "Utilisateur": user,
        })

    # --- création / modification -------------------------------------------
    def create_activity(self, categorie, region, user: str | None = None, **fields) -> str:
        """Crée une activité (informations fixes éventuellement incomplètes) et retourne son code."""
        user, now = user or self.cfg.admin, datetime.now().replace(microsecond=0)
        cat, reg = self.cfg.category_code(categorie), self.cfg.region_code(region)
        unknown = set(fields) - set(EDITABLE)
        if unknown:
            raise ValueError(f"Champs non reconnus : {sorted(unknown)}")
        values = {f: coerce(f, fields.get(f)) for f in EDITABLE}
        values["Statut"] = values["Statut"] or "Actif"
        code = self._next_code(cat, reg)
        row = {"CODE_ACTIVITE": code, "Categorie": cat, "Region": reg, **values,
               "Date_Fin_Previsionnelle": fin_previsionnelle(values["Date_Debut"], values["Delai_Mois"], self.cfg.conventions[cat]),
               "Date_Creation": now, "Date_Modification": now}
        self.base.append("Activites", row)
        self._log(code, "(création)", None, code, user, now)
        return code

    def update_activity(self, code: str, user: str | None = None, changes: dict | None = None, clear=()) -> list[tuple]:
        """Modifie des informations fixes. Une valeur vide/None dans `changes` est IGNORÉE (jamais une suppression) ;
        pour effacer un champ il faut le nommer explicitement dans `clear`. Retourne [(champ, ancienne, nouvelle)]."""
        user, now = user or self.cfg.admin, datetime.now().replace(microsecond=0)
        act = self.get(code)
        targets: dict = {}
        for f, raw in (changes or {}).items():
            if f not in EDITABLE:
                raise ValueError(f"Champ non modifiable : {f} (Categorie et Region font partie du code et ne changent jamais)")
            new = coerce(f, raw)
            if new is not None:
                targets[f] = new
        for f in clear:
            if f not in EDITABLE or f == "Statut":
                raise ValueError(f"Champ qu'on ne peut pas effacer : {f}")
            targets[f] = None
        done = []
        for f, new in targets.items():
            old = act[f]
            if fmt(old) != fmt(new):
                done.append((f, old, new))
        if not done:
            return []
        merged = {**{f: act[f] for f in EDITABLE}, **{f: n for f, _, n in done}}
        update = {f: n for f, _, n in done}
        update["Date_Fin_Previsionnelle"] = fin_previsionnelle(merged["Date_Debut"], merged["Delai_Mois"], self.cfg.conventions[act["Categorie"]])
        update["Date_Modification"] = now
        self.base.update("Activites", act["_r"], update)
        for f, old, new in done:
            self._log(code, f, old, new, user, now)
        return done

    def archive(self, code: str, user: str | None = None):
        return self.update_activity(code, user, {"Statut": "Archivé"})

    # --- contrôles ---------------------------------------------------------
    def verify(self) -> list[str]:
        """Contrôles d'intégrité du référentiel. Liste vide = tout est conforme."""
        problems, seen, maxseq = [], {}, {}
        for a in self.activities():
            code = a["CODE_ACTIVITE"] or ""
            m = CODE_RE.match(code)
            if not m:
                problems.append(f"Ligne {a['_r']} : code {code!r} au mauvais format (attendu 2 lettres + 3 lettres + 3 chiffres)")
                continue
            if code in seen:
                problems.append(f"Code en double : {code} (lignes {seen[code]} et {a['_r']})")
            seen[code] = a["_r"]
            if m[1] not in self.cfg.categories or m[2] not in self.cfg.regions:
                problems.append(f"{code} : catégorie ou région absente de la configuration")
            if (m[1], m[2]) != (a["Categorie"], a["Region"]):
                problems.append(f"{code} : ne correspond pas aux colonnes Categorie/Region ({a['Categorie']}/{a['Region']})")
            maxseq[(m[1], m[2])] = max(maxseq.get((m[1], m[2]), 0), int(m[3]))
        for c in self.base.rows("Compteurs_Codes"):
            if (c["Dernier_Numero"] or 0) < maxseq.get((c["Categorie"], c["Region"]), 0):
                problems.append(f"Compteur {c['Categorie']}{c['Region']} en retard sur les codes existants")
        for (cat, reg), top in maxseq.items():
            nums = {int(CODE_RE.match(a["CODE_ACTIVITE"])[3]) for a in self.activities() if a["Categorie"] == cat and a["Region"] == reg}
            holes = sorted(set(range(1, top + 1)) - nums)
            if holes:
                problems.append(f"Séquence {cat}{reg} : numéros manquants {holes} (normal seulement si une ligne a été supprimée à la main)")
        return problems
