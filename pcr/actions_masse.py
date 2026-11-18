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


def figer_statuts(ref: Referentiel, user: str | None = None, dry_run: bool = False) -> list[tuple[str, str]]:
    """Enregistre le statut actuellement déduit du dernier suivi valide, sans écraser un choix admin.

    En simulation, retourne les changements prévus sans écrire dans la base.
    """
    suivis = latest_valid(ref)
    faits = []
    for a in ref.activities():
        if a.get("Statut_Chantier"):
            continue
        code = a["CODE_ACTIVITE"]
        suivi = suivis.get(code)
        taux = suivi["Taux_Execution_Physique"] if suivi else None
        statut = statut_chantier(taux)
        faits.append((code, statut))
        if not dry_run:
            ref.update_activity(code, user, {"Statut_Chantier": statut})
    return faits


def statuts_a_verifier(ref: Referentiel) -> list[str]:
    """Signale les statuts admin incohérents avec le taux du dernier suivi valide.

    « Suspendu » est une décision de gestion et ne peut pas être déduite du taux.
    """
    suivis = latest_valid(ref)
    ecarts = []
    for a in ref.activities():
        statut = a.get("Statut_Chantier")
        suivi = suivis.get(a["CODE_ACTIVITE"])
        if not statut or statut == "Suspendu" or not suivi:
            continue
        taux = suivi.get("Taux_Execution_Physique")
        if taux is None:
            continue
        suggere = statut_chantier(taux)
        if statut != suggere:
            ecarts.append(
                f"{a['CODE_ACTIVITE']} : statut « {statut} » mais taux d’exécution physique à {taux:g} % "
                f"(statut suggéré : « {suggere} »)."
            )
    return ([f"Statuts à vérifier ({len(ecarts)})"] + ecarts) if ecarts else []


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
