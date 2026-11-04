# Suivi des activités PCR-CI — Phase 2 : référentiel des activités

Base de données = `data/base_pcr.xlsx` (ne pas éditer à la main). L'administrateur travaille dans `data/referentiel_admin.xlsx`.

    python -m pcr charger-initial --bs ... --mr ... [--excel ...]   # une seule fois
    python -m pcr appliquer      # applique le fichier d'administration (sauvegarde auto, historique, codes générés)
    python -m pcr exporter       # réécrit le fichier d'administration depuis la base
    python -m pcr historique BSBAF001
    python -m pcr verifier       # unicité des codes, séquences, compteurs
    python -m pytest -q tests

Règles : code = catégorie(2) + région(3) + numéro(3), séquence propre à chaque couple catégorie+région, jamais réutilisé ;
cellule vide = ignorée (pour effacer : `#VIDER`) ; toute modification est historisée ; Categorie/Region ne changent jamais.
Catégories, régions, chemins : `config/config.yaml`.
