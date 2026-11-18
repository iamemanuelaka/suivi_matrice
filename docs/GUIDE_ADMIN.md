# Guide de l'administrateur

Les informations des activités qui ne viennent pas des techniciens se gèrent dans **`data/referentiel_admin.xlsx`** :
informations contractuelles (désignation, montant, entreprise, ordre de service, dates, délai, bureau de contrôle) **et** informations de suivi
gérées par l'administrateur (montant décaissé, statut du chantier, PGES, ONG mobilisées, remise de site, réception provisoire).
Les techniciens ne renseignent que l'avancement dans Kobo.

## Procédure (5 étapes)
1. `git pull` — toujours partir de la dernière version (le workflow met la base à jour chaque lundi).
2. `python -m pcr exporter` — réécrit le fichier Excel depuis la base.
3. Ouvrir `data/referentiel_admin.xlsx` (onglet **Activites**), modifier, **enregistrer, fermer Excel**.
4. `python -m pcr appliquer` — **lire le rapport** (créations, modifications ancienne → nouvelle valeur, erreurs ligne par ligne).
5. `git add data` puis `git commit -m "Mise à jour référentiel"` et `git push`.
   *Sans push, le workflow GitHub continue avec l'ancienne base.*

## Que modifier, et où
| Besoin | Colonne | Valeurs |
|---|---|---|
| Montant du marché | `Montant_Marche` | nombre en FCFA, sans texte (`120000000`) |
| **Montant décaissé** | `Montant_Decaisse` | nombre ; ne peut pas dépasser le montant du marché |
| Taux financier | `Taux_Financier` *(grise, calculée)* | décaissé ÷ marché, ignorée à la lecture |
| Désignation / intitulé | `Designation` | texte libre |
| Entreprise, groupement ou consortium | `Entreprise` | ex. `ENTREPRISE A / ENTREPRISE B (groupement)` |
| Ordre de service | `Numero_OS`, `Date_OS` | |
| Début des travaux, délai contractuel | `Date_Debut`, `Delai_Mois` | `Date_Fin_Previsionnelle` se recalcule seule |
| **Statut du chantier** | `Statut_Chantier` | **4 valeurs seulement** : `Non démarré`, `En cours`, `Achevé`, `Suspendu` (liste déroulante). Vide = déduit du taux (0 % non démarré, 100 % achevé, sinon en cours). `Suspendu` ne se déduit pas : il se saisit ici. |
| **PGES** | `PGES` | `Validé`, `En cours de validation`, `Non transmis`, `Sans PGES` (liste) |
| **ONG mobilisées** | `ONG_Mobilisees` | nom de l'ONG, ou `Non mobilisée` |
| Remise de site, réception provisoire | `Date_Remise_Site`, `Date_Reception_Provisoire` | `JJ/MM/AAAA` |
| Département, localité, bureau de contrôle | `Departement`, `Localite`, `Bureau_Controle` | |
| Archiver une activité (sans la supprimer) | `Statut` | `Archivé` *(ne pas confondre avec le statut du chantier)* |

## Créer une activité
Ajouter une ligne **en bas**, laisser `CODE_ACTIVITE` vide, choisir `Categorie` (BS, MR) et `Region` dans les listes.
Tous les autres champs peuvent rester vides et être complétés plus tard. Le code (ex. `MRPOR004`) est généré à l'application.
Pour que les techniciens voient la nouvelle activité dans Kobo : `python -m pcr kobo-sync` (en local, avec votre `.env`).

## Règles à retenir
* **Une cellule vide n'efface rien.** Pour effacer volontairement une valeur : écrire `#VIDER` dans la cellule.
* Ne jamais modifier `CODE_ACTIVITE`, `Categorie`, `Region` d'une activité existante ; les colonnes grises sont ignorées.
* Dates `JJ/MM/AAAA` ; délai en mois entiers ; montants en nombres. Une valeur invalide (`120 M`, un statut hors liste, un décaissé supérieur au marché)
  est **refusée** avec son message ; les autres lignes sont quand même appliquées.
* **Le fichier est réécrit depuis la base après `appliquer`** : une cellule refusée est remplacée par l'ancienne valeur.
  Lisez le rapport, puis ressaisissez la valeur corrigée.
* Relancer `appliquer` sans nouvelle modification donne « 0 modifiée ».
* Les valeurs sont écrites dans le tableau de bord **par code d'activité** : changer l'ordre des activités ne décale plus rien.

## Statuts des chantiers : l'administrateur décide
Les 4 statuts : `Non démarré`, `En cours`, `Achevé`, `Suspendu` (les anciens « Démarrage autorisé », « OS à reprendre » et « Abandonné » sont retirés).
* **Où** : colonne `Statut_Chantier` du fichier admin (liste déroulante). La colonne grise `Statut_Suggere` donne la suggestion d'après le taux
  remonté du terrain (0 % non démarré, 100 % achevé, sinon en cours) ; `Suspendu` ne se suggère jamais.
* **Lien avec le tableau de bord** : le statut saisi est celui affiché partout (cartes, filtre « Statut », graphique). Le filtre « Statut » propose
  **toujours les 4 statuts**, même si aucune activité n'a encore « Achevé » ou « Suspendu ».
  Si `Statut_Chantier` est vide, c'est la suggestion qui s'applique.
* **Prendre la main d'un coup** : `python -m pcr statuts-figer --a-blanc` (simulation) puis `python -m pcr statuts-figer` : chaque activité reçoit
  comme statut de départ la suggestion actuelle ; ensuite vous corrigez les exceptions dans le fichier admin
  (`exporter`, modifier, `appliquer`, puis `git add -A`, `commit`, `push`). **Dès lors le statut ne suit plus automatiquement le taux.**
* **Garde-fou** : chaque lundi, le rapport (et le courriel) liste les « Statuts à vérifier » quand le statut saisi contredit le terrain :
  « Non démarré » avec un taux > 0, « En cours » à 100 %, « Achevé » sous 100 %. Un « Suspendu » n'est jamais signalé.
* Un chantier `Achevé` n'est pas évalué pour les alertes (comme dans le tableau de bord).

## Première mise en service (une seule fois, automatique)
Avant, le montant décaissé, le PGES, les ONG… étaient saisis à la main dans le modèle Excel, ligne par ligne. Au **premier traitement** après la
mise en service, ces valeurs sont **reprises une seule fois dans la base** (rapprochement par montant du marché, puis nom d'entreprise ; rien n'est
écrasé ; chaque reprise est historisée sous l'utilisateur « Reprise du modèle »). Le rapport signale les lignes à vérifier.
Pour la voir à l'avance : `python -m pcr migrer-valeurs-modele --a-blanc` ; pour la lancer à la main : `python -m pcr migrer-valeurs-modele`.

## ONG mobilisées : comment les régler
* Colonne `ONG_Mobilisees` : choisir dans la liste (ONG déjà saisies) ou taper le nom. Deux valeurs sans nom : `Non mobilisée` et `Mobilisée`
  (ONG mobilisée dont le nom n'est pas encore connu → comptée dans « activité(s) avec ONG non nommée »). « oui », « aucune », « néant »… sont ramenés à ces deux valeurs.
* Le tableau de bord compte les **ONG distinctes** (« ONG mobilisées : N »). Une même ONG écrite de deux façons (ex. `ASAPSU` et `ASAPSU ONG`) compte pour deux :
  * `python -m pcr ong-liste` — liste des ONG, avec les noms proches signalés (⚠) ;
  * `python -m pcr renommer-ong "ASAPSU ONG" "ASAPSU"` — remplace dans toutes les activités (historisé), puis `git add data`, `commit`, `push`.
* Le PGES (`PGES`, « 13 Validés »…) se règle de la même façon, activité par activité (colonne `PGES` du fichier admin).
* **Même PGES pour tous les chantiers d'un statut** (ex. tous les chantiers « En cours » ont un PGES validé) :
  `python -m pcr pges-en-masse --a-blanc` (simulation : liste ce qui changerait, n'écrit rien), puis `python -m pcr pges-en-masse`,
  puis `git add -A`, `commit`, `push`. Options : `--statut "En cours"` (défaut), `--valeur Validé` (défaut), `--seulement-vides` (ne touche pas aux PGES déjà renseignés).
  Chaque changement est historisé ; les exceptions se corrigent ensuite dans le fichier admin (colonne `PGES`).

## Carte « Activités sélectionnées / total »
Elle affiche le nombre d'activités de la sélection sur le total : **42 / 42** sans filtre ; avec le filtre statut « En cours », par exemple, **27 / 42** ;
avec « Marchés ruraux », 16 / 42. Toutes les activités comptent (y compris `Suspendu`).

## Historique et contrôles
* `python -m pcr historique BSBAF001` — toutes les modifications d'une activité (date, auteur, ancienne → nouvelle valeur).
* `python -m pcr verifier` — contrôle de cohérence du référentiel · `python -m pcr lister` — liste des activités.
* Une copie de la base est faite avant chaque traitement (`data/sauvegardes/`) : pour revenir en arrière, remplacer `data/base_pcr.xlsx`.

## Soumissions Kobo « à vérifier »
Une soumission dont le taux baisse (ex. 40 % → 0 %) est mise « à vérifier » et **n'est pas utilisée** tant qu'elle n'est pas validée :
`python -m pcr suivis --statut A_VERIFIER`, puis `python -m pcr valider <id>` ou `python -m pcr rejeter <id>`.

## Capacité
Le tableau de bord gère jusqu'à **500 activités**. Au-delà : `python -m pcr etendre-modele --capacite 1000`.
