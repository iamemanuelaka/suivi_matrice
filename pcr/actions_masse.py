"""Mises à jour en masse par l'administrateur (historisées, avec simulation).

Exemple : tous les chantiers « En cours » ont un PGES validé  →  pges_en_masse(ref, "En cours", "Validé").
Le statut pris en compte est celui AFFICHÉ au tableau de bord : le statut saisi par l'administrateur, sinon celui déduit du taux.
"""
from __future__ import annotations

from .referentiel import PGES_VALEURS, STATUTS_CHANTIER, Referentiel, _choice, statut_deduit
from .word_gen import latest_valid


def statut_affiche(a: dict, suivi: dict | None) -> str:
    return a["Statut_Chantier"] or statut_deduit(suivi["Taux_Execution_Physique"] if suivi else None)


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


def figer_statuts(ref: Referentiel, user: str | None = None, dry_run: bool = False) -> list[tuple]:
    """Prend le statut SUGGÉRÉ par le taux comme statut de l'administrateur, pour les activités actives qui n'en ont pas encore.
    Point de départ du pilotage par l'administrateur : ensuite le statut ne suit plus automatiquement le taux (voir statuts_a_verifier)."""
    suivis = latest_valid(ref)
    faits = []
    for a in ref.activities():
        if a["Statut"] != "Actif" or a["Statut_Chantier"]:
            continue
        s = suivis.get(a["CODE_ACTIVITE"])
        nouveau = statut_deduit(s["Taux_Execution_Physique"] if s else None)
        faits.append((a["CODE_ACTIVITE"], nouveau))
        if not dry_run:
            ref.update_activity(a["CODE_ACTIVITE"], user, {"Statut_Chantier": nouveau})
    return faits


def statuts_a_verifier(ref: Referentiel) -> list[str]:
    """Écarts entre le statut saisi par l'administrateur et le taux physique remonté du terrain (signalés dans le rapport du lundi)."""
    suivis = latest_valid(ref)
    lignes = []
    for a in ref.activities():
        st = a["Statut_Chantier"]
        s = suivis.get(a["CODE_ACTIVITE"])
        taux = s["Taux_Execution_Physique"] if s else None
        if a["Statut"] != "Actif" or not st or taux is None:
            continue
        if st == "Non démarré" and taux > 0:
            lignes.append(f"{a['CODE_ACTIVITE']} : statut « Non démarré » mais taux d'exécution {taux:g} %")
        elif st == "En cours" and taux >= 100:
            lignes.append(f"{a['CODE_ACTIVITE']} : statut « En cours » mais taux d'exécution {taux:g} % (achevé ?)")
        elif st == "Achevé" and taux < 100:
            lignes.append(f"{a['CODE_ACTIVITE']} : statut « Achevé » mais taux d'exécution {taux:g} %")
    if not lignes:
        return []
    return [f"Statuts à vérifier ({len(lignes)}) : à confirmer ou corriger dans le fichier admin (colonne Statut_Chantier)"] + \
           ["  " + x for x in lignes[:10]] + ([f"  … et {len(lignes) - 10} autre(s)"] if len(lignes) > 10 else [])
