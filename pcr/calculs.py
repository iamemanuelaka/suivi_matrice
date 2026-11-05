"""Indicateurs calculés (jamais saisis). Convention : mois calendaires comme EDATE d'Excel."""
import calendar
from datetime import date, datetime, timedelta


def _d(v):
    return v.date() if isinstance(v, datetime) else v


def add_months(d: date, n: int) -> date:
    """Ajoute n mois calendaires ; si le jour n'existe pas dans le mois d'arrivée, on prend le dernier jour (comme EDATE)."""
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


CONVENTIONS = ("calendaire", "jours_30")


def fin_previsionnelle(debut, delai_mois, convention="calendaire"):
    """Date de début + délai en mois (« calendaire », comme EDATE) ou + 30 jours par mois (« jours_30 »).
    None si l'une des deux données manque ou est incohérente."""
    debut = _d(debut)
    if not isinstance(debut, date) or isinstance(delai_mois, bool) or not isinstance(delai_mois, (int, float)):
        return None
    if delai_mois <= 0 or delai_mois != int(delai_mois):
        return None
    if convention == "jours_30":
        return debut + timedelta(days=30 * int(delai_mois))
    if convention != "calendaire":
        raise ValueError(f"Convention de délai inconnue : {convention!r}")
    return add_months(debut, int(delai_mois))


def indicateurs_delai(debut, delai_mois, au: date, convention="calendaire"):
    """Trois indicateurs distincts du taux d'exécution physique. None si non calculable."""
    fin = fin_previsionnelle(debut, delai_mois, convention)
    if fin is None:
        return None
    debut, au = _d(debut), _d(au)
    duree = (fin - debut).days
    ecoule = max(0, (au - debut).days)          # avant le début : 0 jour écoulé
    return {
        "fin_previsionnelle": fin,
        "duree_jours": duree,
        "ecoule_jours": ecoule,
        "consommation_pct": round(ecoule / duree * 100, 2),
        "depassement_jours": max(0, ecoule - duree),
    }
