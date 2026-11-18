"""Mises à jour en masse par l'administrateur (historisées, avec simulation).

Exemple : tous les chantiers « En cours » ont un PGES validé  →  pges_en_masse(ref, "En cours", "Validé").
Le statut pris en compte est celui AFFICHÉ au tableau de bord : le statut saisi par l'administrateur, sinon celui déduit du taux.
"""
from __future__ import annotations

from .excel_gen import statut_chantier
from .referentiel import PGES_VALEURS, STATUTS_CHANTIER, Referentiel, _choice
from .word_gen import latest_valid


def statut_affiche(a: dict, suivi: dict | None) -> str:
    return a["Statut_Chantier"] or statut_chantier(suivi["Taux_Execution_Physique"] if suivi else None)


def pges_en_masse(ref: Referentiel, statut: str, valeur: str, seulement_vides: bool = False, user: str | None = None,
                  dry_run: bool = False) -> dict:
    """Met `valeur` dans PGES pour toutes les activités actives dont le statut affiché est `statut`.

    Retourne {"modifiees": [(code, ancien)], "deja": [codes déjà à cette valeur], "gardees": [(code, valeur existante)]}.
    """
    statut = _choice("statut du chantier", statut, STATUTS_CHANTIER)
    valeur = _choice("PGES", valeur, PGES_VALEURS)
    suivis = latest_valid(ref)
    rap = {"modifiees": [], "deja": [], "gardees": []}
    for a in ref.activities():
        if a["Statut"] != "Actif" or statut_affiche(a, suivis.get(a["CODE_ACTIVITE"])) != statut:
            continue
        code, ancien = a["CODE_ACTIVITE"], a["PGES"]
        if ancien == valeur:
            rap["deja"].append(code)
        elif seulement_vides and ancien:
            rap["gardees"].append((code, ancien))
        else:
            rap["modifiees"].append((code, ancien))
            if not dry_run:
                ref.update_activity(code, user, {"PGES": valeur})
    return rap
