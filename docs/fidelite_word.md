# Contrôle de fidélité des matrices Word générées

Chaque cellule générée est comparée à la cellule correspondante des matrices fournies (même ordre de lignes). Sont ignorés : étoiles de mise en forme, espaces, virgules, « 5 % » / « 5% », dates mal saisies (17/012/2026), « FCFA HT ». Le fond du texte, les chiffres, les dates et l'ordre des lignes sont comparés.

## Blocs sanitaires
| Colonne | Cellules identiques |
|---|---|
| Désignation | 25 / 26 |
| Avancement | 26 / 26 |
| Commentaires | 25 / 26 |

## Marchés ruraux
| Colonne | Cellules identiques |
|---|---|
| Désignation | 13 / 16 |
| Avancement | 15 / 16 |
| Commentaires | 16 / 16 |

## Différences de contenu restantes

- **BS ligne 7 (BERE) – Commentaires** : source « Difficultés : Entreprise non mobilisée à ce jour Recommandations : Une mise en demeure est transmise » → généré « ∅ »
- **BS ligne 26 (WORODOUGOU) – Désignation** : source « 24/12/2026 » → généré « ∅ »
- **MR ligne 6 (KABADOUGOU) – Désignation** : source « 18/09/2026 » → généré « ∅ » ; source « 24/12/2026 » → généré « ∅ »
- **MR ligne 10 (PORO) – Désignation** : source « 19/02/2027 » → généré « ∅ »
- **MR ligne 10 (PORO) – Avancement** : source « ∅ » → généré « Taux d'avancement : Délai consommé : »
- **MR ligne 11 (TCHOLOGO) – Désignation** : source « 06/01/2027 » → généré « 08/01/2027 »

**Causes possibles** : (1) date « source → ∅ » : début ou fin prévue non enregistré parce que la donnée du Word est incohérente ou absente (voir rapport_chargement_initial.md) : à confirmer dans le fichier d'administration, jamais inventé ; (2) bloc « Difficultés / Recommandations » répété deux fois dans le Word d'origine : une seule occurrence est reprise ; (3) date de fin différente de 2 jours : la fin prévue du Word ne suit pas la convention de calcul du projet ; (4) libellés vides en plus : marché sans suivi, affiché avec les trois lignes habituelles.

## Délai consommé (recalculé)

0 valeur(s) identique(s) sur 38. Le délai consommé du Word n'a pas été calculé à une date unique (ex. Koro : 54,09 % = situation au 22/10 ; Touba : 35,24 % = 29/09). Il est maintenant calculé pour la date de situation du document.

| Ligne | Word | Généré |
|---|---|---|
| BS ligne 1 (BAFING) | 54.09 | 56.56 |
| BS ligne 2 (BAFING) | 35.24 | 56.56 |
| BS ligne 3 (BAGOUE) | 31.14 | 50.82 |
| BS ligne 4 (BAGOUE) | 25.4 | 27.87 |
| BS ligne 5 (BERE) | 48.36 | 50.82 |
| BS ligne 6 (BERE) | 48.36 | 50.82 |
| BS ligne 7 (BERE) | 48.36 | 50.82 |
| BS ligne 8 (BOUNKANI) | 45.9 | 48.36 |
| BS ligne 9 (BOUNKANI) | 22.95 | 25.41 |
| BS ligne 10 (FOLON) | 48.36 | 50.82 |
| BS ligne 11 (FOLON) | 48.36 | 50.82 |
| BS ligne 12 (GONTOUGO) | 36.88 | 45.08 |
| BS ligne 13 (GONTOUGO) | 45.9 | 48.36 |
| BS ligne 14 (HAMBOL) | 45.9 | 56.56 |
| BS ligne 15 (HAMBOL) | 45.9 | 56.56 |
| BS ligne 16 (HAMBOL) | 45.9 | 56.56 |
| BS ligne 17 (KABADOUGOU) | 48.36 | 50.82 |
| BS ligne 18 (KABADOUGOU) | 48.36 | 50.82 |
| BS ligne 19 (KABADOUGOU) | 48.36 | 50.82 |
| BS ligne 20 (PORO) | 40.96 | 48.36 |
| BS ligne 21 (PORO) | 36.08 | 48.36 |
| BS ligne 22 (TCHOLOGO) | 40.96 | 48.36 |
| BS ligne 23 (TCHOLOGO) | 36.08 | 48.36 |
| BS ligne 24 (WORODOUGOU) | 26.22 | 33.61 |
| BS ligne 25 (WORODOUGOU) | 42.62 | 50.82 |
| BS ligne 26 (WORODOUGOU) | 42.62 |  |
| MR ligne 1 (BERE) | 29.5 | 32.5 |
| MR ligne 2 (BOUNKANI) | 31.14 | 39.17 |
| MR ligne 3 (FOLON) | 25.4 | 28.33 |
| MR ligne 4 (GONTOUGO) | 41.8 | 55.83 |
| MR ligne 5 (HAMBOL) | 48.36 | 51.67 |
| MR ligne 6 (KABADOUGOU) | 19.67 |  |
| MR ligne 7 (KABADOUGOU) | 48.36 | 22.5 |
| MR ligne 8 (PORO) | 28.68 | 31.67 |
| MR ligne 9 (PORO) | 19.67 | 48.33 |
| MR ligne 11 (TCHOLOGO) | 34.42 | 37.5 |
| MR ligne 12 (TCHOLOGO) | 31.14 | 34.17 |
| MR ligne 13 (WORODOUGOU) | 19.67 | 21.67 |