# Guide de l'administrateur — informations fixes des activités

Les informations fixes (désignation, montant du marché, entreprise, ordre de service, dates, délai, bureau de contrôle…)
se renseignent **une seule fois** dans `data/referentiel_admin.xlsx`. Les techniciens ne les ressaisissent jamais.

## Procédure (5 étapes)
1. `git pull` — toujours partir de la dernière version (le workflow met la base à jour chaque lundi).
2. `python -m pcr exporter` — réécrit le fichier Excel depuis la base.
3. Ouvrir `data/referentiel_admin.xlsx` (onglet **Activites**), modifier, **enregistrer, fermer Excel**.
4. `python -m pcr appliquer` — **lire le rapport** (créations, modifications ancienne → nouvelle valeur, erreurs ligne par ligne).
5. `git add data` puis `git commit -m "Mise à jour référentiel"` et `git push`.
   *Sans push, le workflow GitHub continue avec l'ancienne base.*

## Que modifier, et où
| Besoin | Colonne |
|---|---|
| Montant du marché (FCFA, nombre sans texte : `120000000`) | `Montant_Marche` |
| Désignation / intitulé | `Designation` |
| Entreprise, groupement ou consortium | `Entreprise` (ex. `ENTREPRISE A / ENTREPRISE B (groupement)`) |
| Ordre de service | `Numero_OS`, `Date_OS` |
| Début des travaux, délai contractuel | `Date_Debut`, `Delai_Mois` → `Date_Fin_Previsionnelle` se recalcule seule |
| Département, localité, bureau de contrôle | `Departement`, `Localite`, `Bureau_Controle` |
| Archiver une activité (sans la supprimer) | `Statut` = `Archivé` |

## Créer une activité
Ajouter une ligne **en bas**, laisser `CODE_ACTIVITE` vide, choisir `Categorie` (BS, MR…) et `Region` dans les listes.
Tous les autres champs peuvent rester vides et être complétés plus tard. Le code (ex. `MRPOR004`) est généré à l'application.
Pour que les techniciens voient la nouvelle activité dans Kobo : `python -m pcr kobo-sync` (en local, avec votre `.env`).

## Règles à retenir
* **Une cellule vide n'efface rien.** Pour effacer volontairement une valeur : écrire `#VIDER` dans la cellule.
* Ne jamais modifier `CODE_ACTIVITE`, `Categorie`, `Region` d'une activité existante ; les colonnes grises sont ignorées.
* Dates `JJ/MM/AAAA` ; délai en mois entiers ; montant en nombre. Une valeur invalide est **refusée** (message « nombre invalide »),
  les autres lignes sont quand même appliquées.
* **Le fichier est réécrit depuis la base après `appliquer`** : une cellule refusée est remplacée par l'ancienne valeur.
  Lisez le rapport, puis ressaisissez la valeur corrigée.
* Aucune ressaisie double : relancer `appliquer` sans nouvelle modification donne « 0 modifiée ».

## Historique et contrôles
* `python -m pcr historique BSBAF001` — toutes les modifications d'une activité (date, auteur, ancienne → nouvelle valeur).
* `python -m pcr verifier` — contrôle de cohérence du référentiel.
* `python -m pcr lister` — liste des activités.
* Une copie de la base est faite avant chaque traitement (`data/sauvegardes/`) : pour revenir en arrière, remplacer `data/base_pcr.xlsx`
  par l'une de ces copies.

## Capacité
Le tableau de bord gère jusqu'à **500 activités**. Au-delà : `python -m pcr etendre-modele --capacite 1000`.
