# Phase 3 — Mise en place du questionnaire Kobo (serveur kf.kobotoolbox.org)

## Ce qui est livré
- `kobo/questionnaire_suivi_pcr.xlsx` : le formulaire (XLSForm), validé par pyxform et ODK Validate.
- `data/kobo/activites.csv` : la liste des 42 activités (fichier média du formulaire), générée depuis la base.
- Commandes : `python -m pcr kobo-formulaire`, `kobo-csv`, `kobo-sync`.

## Ce que voit le technicien
1. Nom et prénoms, date de l'actualisation (aujourd'hui par défaut, pas de date future).
2. **Activité suivie** : liste filtrable (taper le code, la région, le département ou l'entreprise).
3. Dès qu'il choisit, une note en lecture seule affiche : désignation, titulaire, montant, délai, début, fin prévue et le dernier taux enregistré.
4. Observations techniques (obligatoire, champ dédié), taux d'exécution physique (entier 0-100, obligatoire).
5. Commentaires, difficultés, recommandations (facultatifs ; « RAS » sera affiché dans le Word si vide).
Aucune information contractuelle n'est ressaisie.

## Première mise en place (une seule fois, manuelle)
1. Kobo → Nouveau projet → *Téléverser un XLSForm* → `questionnaire_suivi_pcr.xlsx`.
2. Dans le projet : *Paramètres → Médias* → ajouter `activites.csv`.
3. **Déployer** le formulaire. Récupérer l'identifiant du projet (UID `a...` dans l'URL) et le jeton (*Compte → Sécurité*).
4. Partager le projet avec les techniciens en **« Soumettre des données »** uniquement (jamais « Modifier le formulaire »).
5. Sur votre poste : copier `.env.example` en `.env`, le remplir (`KOBO_ASSET_UID`, `KOBO_TOKEN`), puis `python -m pcr kobo-sync`.

## Quand une activité est créée ou modifiée
`python -m pcr appliquer` puis `python -m pcr kobo-sync` : la liste est envoyée à Kobo et le formulaire redéployé.
Le formulaire lui-même n'a pas à être réimporté, sauf si les champs de suivi changent (alors : `kobo-formulaire`, puis réimport).
Ajouter une catégorie ou une région ne demande aucune modification du formulaire.

## À TESTER sur Kobo (je n'ai pas pu joindre votre serveur depuis mon environnement)
| # | Test | Attendu |
|---|---|---|
| 1 | Web (Enketo) : choisir une activité | la note de référence s'affiche avec les bonnes valeurs |
| 2 | Web : taper « Koro » dans la liste | la liste se filtre |
| 3 | Web : taux 150, date de demain, observation vide | chaque erreur bloque l'envoi avec son message |
| 4 | KoboCollect (Android), en ligne : mêmes tests 1 à 3 | identique |
| 5 | KoboCollect **hors ligne** : ouvrir, remplir, enregistrer, envoyer plus tard | la liste reste disponible ; l'envoi fonctionne au retour du réseau |
| 6 | `kobo-sync` avec une activité modifiée, puis rouvrir le formulaire sur le téléphone | la nouvelle valeur apparaît (peut demander « Obtenir le formulaire vierge » dans Collect) |
| 7 | Soumettre un suivi test, vérifier les noms de champs dans l'export | `code_activite`, `observations_techniques`, `taux_execution`, `commentaires`, `difficultes`, `recommandations` |
| 8 | `kobo-sync` deux fois de suite | pas de doublon du fichier `activites.csv` dans les médias |

Points techniques à confirmer par ces tests : l'envoi du fichier média par l'API (comportement face à un fichier déjà présent),
et le rafraîchissement de la liste sur les téléphones après redéploiement. Les téléphones doivent re-télécharger le formulaire vierge.
