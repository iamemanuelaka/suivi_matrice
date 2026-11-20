# Application PCR-CI : guide d'utilisation

L'application fait tout ce que faisaient les commandes `python -m pcr …`, avec des boutons. Elle ne demande ni Python ni git sur l'ordinateur.

## 1. Fabriquer puis installer

**Fabriquer (une fois par version, sur l'ordinateur du responsable, qui a Python 3.12 de python.org)** : double-cliquer sur
`FABRIQUER_APPLICATION.bat` à la racine du dépôt. Il installe ce qu'il faut dans un dossier à part (`.venv-fabrication`), vérifie
le code (tests), fabrique l'application et ouvre le dossier `dist` : `PCR-CI-portable-AAAA.MM.JJ.zip` (et
`PCR-CI-Installation-AAAA.MM.JJ.exe` si Inno Setup 6 est installé). Première fois : 10 minutes environ, ensuite 3 à 5.

**Installer (sur l'ordinateur du gestionnaire, qui n'a besoin ni de Python ni de Git)** :
- ZIP portable : le dézipper (par exemple dans Documents), puis double-clic sur `PCR-CI.exe` ; clic droit → Envoyer vers → Bureau
  pour avoir l'icône sur le bureau ;
- ou installateur : le lancer, « Créer une icône sur le bureau », Installer.

> Windows peut afficher « Windows a protégé votre ordinateur » (logiciel non signé) : cliquez sur **Informations complémentaires → Exécuter quand même**.

## 2. Premier démarrage : Paramètres

| Section | À remplir | Où trouver l'information |
|---|---|---|
| Général | Votre nom | Il apparaîtra dans l'historique des modifications |
| GitHub | Dépôt `iamemanuelaka/suivi_matrice`, branche `main`, **jeton** | github.com → photo de profil → *Settings* → *Developer settings* → *Personal access tokens* → *Fine-grained tokens* → *Generate new token* : choisir le dépôt `suivi_matrice`, permissions **Contents : Read and write** et **Actions : Read and write** |
| KoboToolbox | Adresse, UID du formulaire, jeton API | Les mêmes valeurs que les secrets `KOBO_ASSET_UID` / `KOBO_TOKEN` du dépôt. Jeton : kf.kobotoolbox.org → *Account settings* → *API key* |
| E-mail | Serveur SMTP, utilisateur, mot de passe d'application, destinataires | Les mêmes valeurs que les secrets `SMTP_*` / `MAIL_TO` du dépôt |

Cliquez sur **Tester** dans chaque section, puis sur **Synchroniser maintenant** (GitHub) : la base, les photos, les modèles et la configuration sont récupérés.

Les jetons et mots de passe restent dans votre profil Windows (`%APPDATA%\PCR-CI\parametres.json`), jamais dans le dépôt.

## 3. Utilisation

- **Tableau de bord** : les mêmes filtres, cartes, graphique et alertes que le classeur Excel.
- **Lancer le suivi**, deux possibilités ; dans les deux cas, les documents arrivent dans `Documents\PCR-CI\sorties` et s'ouvrent avec le bouton **Ouvrir** :
  - **Lancer sur GitHub** (recommandé) : au choix, **Matrices Word** (fiches Kobo et photos, base, matrices Word) ou **Tableau de bord Excel**, e-mail si la case est cochée ; avec les accès Kobo et e-mail enregistrés sur GitHub. L'écran affiche l'étape en cours, puis télécharge la base mise à jour et les documents produits. Vous pouvez fermer l'application pendant le traitement : le suivi reprend à la réouverture.
  - **Générer les matrices ici** / **Générer le tableau de bord ici** : le même travail, fait sur l'ordinateur (accès Kobo et e-mail à renseigner dans Paramètres).
- **Suivis à valider** : cliquez sur une fiche pour voir les observations et la photo, puis **Valider** ou **Rejeter** (le motif est obligatoire pour rejeter).
- **Activités** : recherche, création (**Nouvelle activité**, avec un code attribué automatiquement), modification, historique. L'onglet **Chantier** de chaque activité sert à saisir ce que les agents ne renseignent pas dans Kobo : statut du chantier, PGES, ONG (les noms déjà saisis sont proposés), montant décaissé (le taux d'exécution financière s'affiche), dates de remise du site et de réception. Après la création d'une activité, cliquez sur **Envoyer la liste des activités à Kobo** pour que les agents la voient dans le formulaire.

## 4. GitHub en coulisse

- Avant chaque modification, l'application récupère la dernière base sur GitHub. Après la modification, elle la renvoie.
- Le traitement automatique du lundi (GitHub Actions) continue de fonctionner, même ordinateur éteint.
- Sans connexion internet, vous pouvez continuer à travailler : la base est enregistrée sur l'ordinateur et sera envoyée à la prochaine synchronisation.
- **Conflit** (la base a changé à la fois ici et sur GitHub) : rien n'est écrasé. Dans Paramètres → GitHub, choisissez **Conflit : garder…** la version GitHub ou celle de cet ordinateur. La version écartée est sauvegardée dans `data\sauvegardes`.

## 5. En cas de problème

- La fenêtre ne s'ouvre pas : utilisez le raccourci **PCR-CI (dans le navigateur)** du menu Démarrer.
- Journal technique : `%APPDATA%\PCR-CI\journal.log`.

## Pour le développeur

```bash
pip install -r requirements.txt -r requirements-app.txt
python -m pcr_app               # fenêtre
python -m pcr_app --navigateur  # dans le navigateur (http://127.0.0.1:8765)
python -m pytest -q tests       # inclut les tests de l'interface (utilisateur simulé)
```

Code : `pcr_app/services.py` (une fonction par bouton, appelle le moteur `pcr`), `pcr_app/github_sync.py` (synchronisation par l'API GitHub), `pcr_app/app.py` (écrans).
Nouvelle version de l'application : relancer `FABRIQUER_APPLICATION.bat` et redonner le ZIP (paramètres et données du gestionnaire conservés).
