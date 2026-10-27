# Software Design Description (SDD)

## 1. Objet et périmètre

Ce document décrit l’application de bureau de la Bibliothèque d’outils métier et son flux de contribution Git. Il complète le mode d’emploi de [README.md](./README.md) et le guide utilisateur affiché dans la rubrique **Parcours Git** ([ROADMAP.md](./ROADMAP.md)).

L’application aide une équipe à explorer des modules VBA et des scripts Python, à modifier certaines sources, puis à enregistrer et publier explicitement des changements. Elle ne remplace ni la revue de code, ni les contrôles de qualité de l’équipe, ni les règles de protection du dépôt distant.

## 2. Objectifs fonctionnels

- Découvrir les scripts `.py` et modules `.bas` présents dans les dossiers métier.
- Afficher une fiche d’outil avec une description indicative, des procédures reconnues et des fichiers associés.
- Ouvrir une source dans un éditeur intégré, puis l’enregistrer.
- Afficher l’état Git du dépôt, ses fichiers modifiés et les fichiers non suivis.
- Permettre le changement immédiat de branche depuis l’interface et recharger le catalogue sans redémarrer l’application.
- Laisser l’utilisateur choisir les fichiers à inclure, le message de commit et une branche locale/distante, ou créer une branche.
- Créer un commit local et tenter de le pousser vers `origin`.
- Permettre l’authentification HTTPS par identifiant et jeton à usage ponctuel, ou réutiliser la configuration SSH/Git existante.

## 3. Technologies et dépendances

- Python 3.10+.
- PySide6 / Qt 6 pour les fenêtres et les signaux.
- Git en ligne de commande pour l’état du dépôt, les branches, le commit et le push.
- Les dépendances tierces des outils métier sont déclarées dans `requirements.txt` et ne sont pas utilisées par l’interface Git.

Git est une dépendance système et ne figure pas comme bibliothèque Python dans `requirements.txt`. La liste des paquets Python et ses annotations sont maintenues dans ce fichier.

## 4. Architecture logique

### Présentation

`app/main_window.py` construit la fenêtre principale avec une bibliothèque, la page **Gestion Git** et le guide affiché depuis `ROADMAP.md`. La page Git regroupe les champs remote/identifiants, le statut, la liste à cocher, la cible de branche et le message de commit.

### Catalogue et édition

`app/catalog.py` découvre et décrit les sources Python et VBA. L’éditeur de `MainWindow` lit une source en conservant son encodage reconnu et ses fins de ligne, puis l’enregistre dans son emplacement courant. L’enregistrement ne crée pas un commit automatiquement ; l’utilisateur choisit ensuite les changements à publier.

### Opérations Git

`GitPushTask`, un `QRunnable`, réalise les commandes Git dans un worker Qt pour ne pas bloquer l’interface. Les arguments sont transmis à `subprocess.run` sous forme de liste, sans interprétation par un shell. Le worker émet un signal de succès ou d’échec consommé par la fenêtre.

`GitBranchSwitchTask` réalise séparément le changement de branche. Il refuse le basculement si l’arbre de travail contient des modifications suivies ou non suivies. Après succès, la fenêtre recharge le catalogue de sources et le statut Git ; il n’est pas nécessaire de redémarrer l’application.

`GitFetchTask` récupère les références distantes (`git fetch --prune origin`) en tâche de fond. Il configure `origin` depuis l’URL saisie, utilise au besoin le même mécanisme temporaire d’authentification HTTPS, puis actualise la liste des branches. Il ne modifie pas les fichiers de travail ni la branche active.

Les commandes principales sont :

1. vérifier l’absence de changements déjà indexés par l’utilisateur ;
2. ajouter ou mettre à jour l’URL du remote `origin` ;
3. sélectionner une branche locale, une branche distante connue de `origin`, `main`, ou créer une branche nouvelle ;
4. indexer les chemins explicitement cochés ;
5. créer un commit si ces chemins produisent un changement indexé ;
6. pousser la branche choisie vers `origin`.

Les erreurs de Git sont affichées dans la fenêtre. Une erreur après le commit ne supprime pas le commit local.

## 5. Flux utilisateur Git

1. Le dépôt de travail est obtenu avec `git clone`. Une archive ZIP ne contient pas les métadonnées `.git` requises.
2. Le contributeur configure `user.name` et `user.email` dans la configuration locale du dépôt s’ils ne sont pas déjà définis. Ces valeurs identifient l’auteur du commit.
3. Il modifie et enregistre un fichier source depuis la bibliothèque.
4. Au besoin, il récupère les dernières branches distantes avec le bouton prévu. Il actualise ensuite l’état, coche les fichiers voulus et choisit une branche locale ou distante, ou en crée une nouvelle, puis saisit un résumé.
5. Il confirme les chemins montrés dans la boîte de confirmation.
6. Le worker crée le commit et le pousse ; l’interface indique le résultat.
7. Pour une nouvelle branche, le contributeur ouvre ensuite une Pull request sur la plateforme d’hébergement.

## 6. Modèle d’authentification

### HTTPS

L’utilisateur peut entrer un nom d’utilisateur et un jeton. Les deux champs doivent être remplis ensemble. Si ces identifiants ponctuels sont présents :

- l’interface masque le jeton ;
- elle efface les champs au démarrage du fetch ou du push ;
- le worker passe les valeurs à un helper `GIT_ASKPASS`, et non aux arguments de commande ou à l’URL du remote ;
- la valeur du secret est conservée dans un fichier temporaire à accès restreint pendant l’opération, puis nettoyée par le gestionnaire de répertoire temporaire ;
- l’appel du push neutralise les `credential.helper` configurés afin que des identifiants précédemment mémorisés ne remplacent pas le compte explicitement saisi ;
- les secrets connus sont retirés du message d’erreur avant son émission vers l’interface.

Les plateformes peuvent demander un jeton personnel (PAT) à la place d’un mot de passe. Pour GitHub, recommander un jeton fine-grained limité au dépôt concerné et à `Contents: Read and write`. Le jeton ne donne pas davantage de droits que le compte qui l’a créé.

### SSH ou authentification existante

Si les champs sont vides, aucune variable `GIT_ASKPASS` temporaire n’est définie et Git utilise son mécanisme d’authentification configuré, comme SSH ou Git Credential Manager. Les champs d’identifiants de l’application sont réservés aux URL HTTPS.

## 7. Modèle de données et persistance

- Les sources éditées sont écrites directement au chemin du fichier choisi.
- Les notes métier sont stockées dans `module_descriptions.json`.
- Les commits et les branches sont stockés dans la base Git du dépôt local.
- Le remote `origin` est stocké dans la configuration Git locale du dépôt.
- Les identifiants HTTPS saisis ne sont pas écrits dans `module_descriptions.json`, dans le remote, ni dans une configuration Git persistante.
- Les autres modifications locales non sélectionnées restent dans l’arbre de travail et ne sont pas ajoutées par l’application au commit.

## 8. Sécurité et limites

- L’utilisateur doit relire les sources et les chemins avant le commit ; une confirmation de publication liste les fichiers sélectionnés.
- Ne jamais inclure un jeton dans une URL, une source, une note libre ou un commit.
- L’authentification n’octroie pas de permission à un utilisateur qui n’en a pas déjà sur le dépôt.
- Les branches distantes proposées correspondent aux références `origin/*` présentes dans le clone. Le bouton **Actualiser** ne fait pas de `fetch`. Une branche locale de suivi est créée depuis cette référence avant le commit si besoin.
- Le bouton **Récupérer les branches distantes** effectue explicitement un `fetch --prune` ; le simple bouton **Actualiser** n’interroge que l’état local. Une branche locale de suivi est créée depuis une référence distante avant le commit ou le changement de branche si besoin.
- Le changement de branche depuis l’interface refuse le changement si l’arbre de travail n’est pas propre.
- Les politiques du serveur peuvent refuser un push direct vers `main` ou vers une autre branche protégée.
- Le flux ne fait pas automatiquement de `pull`, rebase ou résolution de conflits. Le `fetch` n’est déclenché que par le bouton **Récupérer les branches distantes** ; il ne synchronise pas le contenu des fichiers ni la branche de travail.
- Un commit est créé localement avant le push. Si le push échoue, le commit demeure local ; l’utilisateur doit suivre le guide de reprise plutôt que répéter aveuglément la publication ou forcer un push.
- Si des changements sont déjà dans l’index Git, le worker arrête le flux pour éviter d’en mélanger le contenu avec les fichiers cochés.
- Un processus interrompu de force pendant une publication peut empêcher le nettoyage normal de données temporaires par le système d’exploitation ; le jeton doit donc rester à durée limitée et être révoqué s’il a pu être exposé.

## 9. Vérification

La suite automatisée est exécutée depuis la racine :

```bash
python -m unittest discover -s tests
```

Les tests pertinents vérifient la navigation Qt, l’enregistrement de l’éditeur, la publication sur un dépôt local temporaire, le choix explicite des fichiers, la création d’une branche, les erreurs de commit/push et le helper d’identification temporaire.

## 10. Décisions d’évolution

- Garder Git comme outil natif appelé en sous-processus : aucun paquet `GitPython` n’est nécessaire pour le flux actuel.
- Garder les identifiants ponctuels séparés de la configuration persistante du dépôt.
- Favoriser les branches de travail et les Pull requests plutôt que le push direct sur `main`.
- Si un flux de synchronisation (`fetch`/`pull`), une sélection de dépôt multi-projets ou un retry graphique de commit local échoué est ajouté, documenter et tester explicitement les conséquences sur l’arbre de travail et l’historique.
