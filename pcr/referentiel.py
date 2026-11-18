def statut_deduit(taux) -> str:
    """Déduit le statut chantier du taux physique.

    - taux absent ou inférieur ou égal à 0 : « Non démarré »
    - taux inférieur à 100 : « En cours »
    - taux supérieur ou égal à 100 : « Achevé »
    """
    if taux is None or taux <= 0:
        return "Non démarré"
    return "Achevé" if taux >= 100 else "En cours"
