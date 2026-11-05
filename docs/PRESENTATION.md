# Aperçu du système de suivi PCR-CI — notes de présentation

## Le principe en une phrase
Les techniciens ne saisissent que l'avancement dans Kobo ; tout le reste (matrices Word, tableau de bord Excel, e-mail) est produit automatiquement
depuis une base unique, avec historique complet.

## Parcours
Référentiel des activités (codes BSBAF001…) → Kobo (liste filtrable, infos fixes en lecture seule) → récupération automatique (doublons ignorés,
saisies invalides conservées avec motif) → base (Excel) → 2 matrices Word + tableau de bord → e-mail → journal.

## Ce qui est réalisé et testé (82 tests automatiques)
| Étape | État |
|---|---|
| Référentiel : 42 activités (26 BS, 16 MR), codes uniques et permanents, modifications tracées, informations incomplètes acceptées | fait |
| Questionnaire Kobo (validé ODK Validate) + liste d'activités générée et synchronisable | fait, **à tester sur votre serveur Kobo** |
| Récupération/validation/doublons/historique (VALIDE, À VÉRIFIER, REJETÉ) | fait |
| Matrices Word (BS, MR) générées depuis la base, contrôle de fidélité ligne à ligne | fait |
| Tableau de bord Excel alimenté sans toucher aux formules/graphiques/images | fait |
| Automatisation GitHub Actions (lundi 7 h + lancement manuel), sauvegardes, journal, e-mail | fait, **à activer** (secrets) |

## Démonstration (2 minutes, sans risque pour les vraies données)
    python -m pcr demo
Rejoue 8 soumissions fictives sur une COPIE de la base : 4 valides, 1 doublon ignoré, 2 rejetées (code inconnu, taux 150), 1 « à vérifier »
(taux en baisse). Les 3 documents mis à jour sont écrits dans `demo_sortie/`.

## Points d'honnêteté à connaître avant de présenter
1. **Fidélité des Word** : contenu et structure identiques à vos matrices (voir `fidelite_word.md` : 25/26 et 13/16 désignations, 26/26 et 15/16
   avancements, 25/26 et 16/16 commentaires). Les écarts restants viennent de données incohérentes dans le Word d'origine, que le système refuse
   d'inventer. La mise en page est une reconstruction : vos fichiers fournis étaient des extraits texte, pas de vrais .docx.
2. **Délai consommé** : recalculé à la date de situation. Dans votre Word il avait été calculé à des dates différentes selon les lignes.
3. **Non collecté par le nouveau questionnaire** (hors cahier des charges) : décaissements, PGES, ONG, réception provisoire. Ces colonnes du tableau
   de bord gardent leurs dernières valeurs pour les blocs sanitaires. Le statut du chantier est dérivé du taux (0 = non démarré, 1-99 = en cours, 100 = achevé).
4. **Marchés ruraux** : fin prévue = début + 30 jours par mois (règle de votre Word), blocs sanitaires = mois calendaires. Modifiable dans `config/config.yaml`.
5. Le formulaire Kobo et l'envoi de la liste d'activités n'ont pas pu être testés sur un vrai serveur Kobo (voir `kobo_mise_en_place.md`).
