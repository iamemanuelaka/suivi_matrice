# Suivi des activités PCR-CI — système complet (référentiel, Kobo, base, Word, Excel, e-mail)

Base de données = `data/base_pcr.xlsx` (ne pas éditer à la main). L'administrateur travaille dans `data/referentiel_admin.xlsx`.

    python -m pcr charger-initial --bs ... --mr ... [--excel ...]   # une seule fois
    python -m pcr appliquer      # applique le fichier d'administration (sauvegarde auto, historique, codes générés)
    python -m pcr exporter       # réécrit le fichier d'administration depuis la base
    python -m pcr historique BSBAF001
    python -m pcr verifier       # unicité des codes, séquences, compteurs
    python -m pcr kobo-formulaire   # génère le questionnaire Kobo (Phase 3)
    python -m pcr kobo-csv          # génère la liste des activités
    python -m pcr kobo-sync         # envoie la liste à Kobo et redéploie (voir docs/kobo_mise_en_place.md)
    python -m pcr recuperer         # récupère les soumissions Kobo (--from-json pour tester hors ligne)
    python -m pcr suivis [--statut A_VERIFIER] ; valider <id> ; rejeter <id> "motif"
    python -m pcr generer-word --au JJ/MM/AAAA ; generer-dashboard ; verifier-word ...
    python -m pcr lancer [--mail]   # cycle complet (c'est ce que fait GitHub Actions chaque lundi)
    python -m pcr demo              # démonstration sur une copie de la base
    python -m pytest -q tests

Règles : code = catégorie(2) + région(3) + numéro(3), séquence propre à chaque couple catégorie+région, jamais réutilisé ;
cellule vide = ignorée (pour effacer : `#VIDER`) ; toute modification est historisée ; Categorie/Region ne changent jamais.
Catégories, régions, chemins : `config/config.yaml`.
