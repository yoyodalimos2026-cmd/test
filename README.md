# Bibliothèque d’outils métier

Application de bureau pour parcourir, comprendre et modifier les scripts Python et modules VBA de l’équipe. Elle permet aussi de créer un commit et de publier les fichiers sélectionnés vers un dépôt Git distant, sans demander aux utilisateurs de manipuler les commandes Git courantes.

- [Installation et lancement](#installation-et-lancement)
- [Utiliser la bibliothèque](#utiliser-la-bibliothèque)
- [Modifier un fichier et le publier avec Git](#modifier-un-fichier-et-le-publier-avec-git)
- [Configurer un jeton GitHub](#configurer-un-jeton-github)
- [Résoudre les problèmes Git courants](#résoudre-les-problèmes-git-courants)
- [Dépendances et tests](#dépendances-et-tests)
- [Conception de l’application (SDD)](./SDD.md)

Le guide utilisateur intégré à l’application se trouve également dans [ROADMAP.md](./ROADMAP.md).

## Installation et lancement

### Prérequis

- Python 3.10 ou plus récent.
- Git installé et disponible dans le terminal et le `PATH`.
- Une interface graphique compatible avec Qt 6.
- Un compte autorisé sur le dépôt distant pour pousser des modifications.

Une archive ZIP téléchargée depuis GitHub **n’est pas un dépôt Git**. Clonez le dépôt avec Git afin que l’application puisse consulter son état, créer des commits et pousser des branches.

### Linux

Depuis un terminal :

```bash
git clone https://github.com/<organisation>/<depot>.git
cd <depot>
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

Sur une installation Ubuntu/Debian minimale, les bibliothèques système requises par Qt peuvent devoir être installées par l’équipe informatique.

### Windows

Depuis PowerShell :

```powershell
git clone https://github.com/<organisation>/<depot>.git
Set-Location <depot>
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

Si l’activation PowerShell est bloquée, ouvrir `cmd.exe` dans le dépôt et utiliser :

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python main.py
```

## Utiliser la bibliothèque

La page **Bibliothèque** recense les fichiers `.bas` et `.py` des dossiers métier. La recherche accepte un nom de fichier, une procédure ou un dossier ; le filtre permet d’afficher les modules VBA ou les scripts Python.

La fiche d’un fichier présente un résumé indicatif de son utilité, les procédures reconnues et les liens détectés avec des fichiers du même dossier. Le résumé est généré automatiquement à partir du code : il peut être incomplet ou général et ne remplace pas une revue métier ou technique.

**Ajouter une note métier** permet de conserver un complément dans `module_descriptions.json`. Ne mettez pas dans cette note de donnée personnelle, de secret ni d’information confidentielle.

### Lancer un script Python

1. Sélectionnez un fichier Python dans la bibliothèque.
2. Cliquez sur **Lancer le script**.
3. Suivez le statut et les éventuelles fenêtres ou sorties produites par le script.

Le script est lancé avec l’interpréteur Python de l’application et avec son dossier source comme répertoire de travail. Les scripts s’exécutent avec les droits de l’utilisateur et peuvent modifier des fichiers ou demander des dépendances supplémentaires.

## Modifier un fichier et le publier avec Git

### Préparer le dépôt

Clonez le projet avec Git, puis ouvrez l’application à partir de ce dossier. Pour vérifier que Git est installé :

```bash
git --version
```

Git enregistre l’auteur de chaque commit séparément des identifiants d’accès GitHub. Si aucun nom et aucune adresse ne sont configurés pour le dépôt, configurez-les une fois dans un terminal ouvert à sa racine :

```bash
git config --local user.name "Prénom Nom"
git config --local user.email "adresse-associée-au-compte@example.com"
```

Utilisez une adresse de commit appropriée à la politique de votre équipe, éventuellement l’adresse `noreply` fournie par GitHub. Cette identité est inscrite dans les commits ; elle ne permet pas de s’authentifier et n’accorde aucun droit d’écriture.

### Publier avec la page Gestion Git

1. Dans **Bibliothèque**, sélectionnez un fichier `.py` ou `.bas`, puis cliquez sur **Ouvrir le fichier source**.
2. Modifiez le texte et cliquez sur **Enregistrer**. **Annuler** abandonne les modifications non enregistrées.
3. Ouvrez la page **Gestion Git** et vérifiez l’URL HTTPS du dépôt distant. L’URL doit correspondre au dépôt cloné et au compte ou à l’organisation qui le possède. Le remote utilisé est `origin`.
4. Pour une URL HTTPS, saisissez votre nom d’utilisateur et un jeton dans les champs d’identification. Pour une URL SSH, laissez ces champs vides et configurez votre clé SSH dans Git.
5. Cliquez sur **Actualiser**. La liste indique les fichiers modifiés, supprimés ou non suivis et recharge les branches déjà présentes dans le clone.
6. Cochez uniquement les fichiers qui doivent faire partie de ce changement. Vérifiez les chemins avant de continuer.
7. Choisissez **Nouvelle branche** de préférence et indiquez un nom tel que `feature/description-outil` ou `fix/correction-export`, ou sélectionnez une branche existante affichée dans la liste. Les branches distantes connues sont indiquées comme « distante » ; la copie locale de suivi est créée automatiquement à la publication. Choisissez `main` seulement si votre équipe autorise le push direct.
8. Écrivez un résumé clair du changement, par exemple `docs: clarifier les instructions Git`.
9. Cliquez sur **Publier les fichiers cochés** et relisez la confirmation avant de valider.
10. En cas de succès, le commit a été créé localement puis envoyé vers le dépôt distant.

L’application prépare uniquement les fichiers cochés. Les autres changements du dossier de travail restent non inclus. Si des fichiers étaient déjà indexés (`git add`) dans un terminal, l’opération est arrêtée afin de ne pas les inclure par erreur. Le commit est créé sur la branche sélectionnée ; une branche distante sans branche locale correspondante est d’abord créée en suivi de `origin`.

Pour proposer un changement sur une nouvelle branche, ouvrez ensuite GitHub, comparez cette branche à `main` et créez une **Pull request** pour demander une revue. Les organisations peuvent exiger une Pull request et interdire les pushes directs sur `main`.

### Changer de branche sans fermer l’application

Dans **Gestion Git**, choisissez une branche existante dans la liste puis cliquez sur **Basculer vers cette branche**. L’application met à jour la branche active et recharge la bibliothèque sans redémarrage. Les modifications non publiées (y compris les fichiers non suivis) bloquent le basculement pour éviter toute perte : publiez-les ou traitez-les avant de changer.

Les branches distantes affichées sont celles que le clone connaît déjà. Cliquez sur **Récupérer les branches distantes** pour télécharger les dernières références depuis `origin` directement dans l’application ; si le dépôt est privé, saisissez l’identifiant HTTPS et le jeton avant de cliquer. **Actualiser** relit ensuite l’état local sans contacter le serveur.

### Nom d’utilisateur, mot de passe et jeton GitHub

Le **nom d’utilisateur GitHub** est le nom de compte affiché sur le profil ou dans l’URL `https://github.com/nom-du-compte`. Ce n’est ni le nom complet ni l’adresse e-mail.

GitHub n’accepte pas le mot de passe du compte pour une opération Git HTTPS. Créez un **Fine-grained personal access token** (« jeton d’accès personnel à accès limité ») :

1. Connectez-vous avec le compte qui a accès au dépôt et ouvrez [la page de création de jeton GitHub](https://github.com/settings/personal-access-tokens/new).
2. Donnez au jeton un nom reconnaissable et une date d’expiration.
3. Choisissez comme **Resource owner** le compte ou l’organisation qui possède le dépôt.
4. Dans **Repository access**, choisissez **Only select repositories**, puis sélectionnez uniquement le dépôt concerné.
5. Dans **Repository permissions**, ajoutez **Contents** et choisissez **Read and write**. N’accordez pas de permissions supplémentaires sans nécessité.
6. Générez le jeton et copiez-le immédiatement dans un endroit sûr. GitHub ne permet généralement pas d’afficher à nouveau sa valeur complète.
7. Dans **Gestion Git**, entrez le nom de compte et collez le jeton dans le champ masqué **Jeton d’accès personnel**. Ne saisissez pas le mot de passe GitHub.

Le compte doit lui-même disposer du droit d’écriture sur le dépôt. Un jeton ne permet pas de contourner les droits de l’organisation ; les jetons d’une organisation peuvent aussi nécessiter une approbation de son administrateur.

L’application utilise les identifiants saisis uniquement pour le push en cours, ne les inscrit pas dans l’URL du dépôt ou dans la configuration Git, et les efface du formulaire lorsqu’elle démarre la publication. Pour l’authentification HTTPS, Git les lit à partir d’un fichier temporaire protégé, supprimé à la fin du traitement normal. Si les champs sont vides, Git utilise l’authentification déjà configurée, par exemple SSH.

> Ne partagez jamais votre jeton et ne le commitez jamais dans le projet. En cas d’exposition, révoquez-le immédiatement dans les paramètres GitHub et créez-en un nouveau.

## Résoudre les problèmes Git courants

| Message ou situation | Que vérifier / faire |
|---|---|
| `Authentication failed` ou jeton refusé | Vérifier le nom de compte, l’expiration, les permissions `Contents: Read and write`, le dépôt sélectionné et les droits d’écriture du compte. Recréer ou faire approuver le jeton si besoin. |
| `Repository not found` ou accès refusé | Vérifier l’URL HTTPS complète, le propriétaire, le dépôt cloné, les droits du compte et le propriétaire sélectionné dans le jeton. |
| Aucun changement n’apparaît | Enregistrer le fichier source, cliquer sur **Actualiser** et vérifier que l’application a ouvert le dossier du dépôt cloné, pas une archive ZIP ou une autre copie. |
| Échec lors du commit, identité inconnue | Définir `user.name` et `user.email` avec les commandes ci-dessus. Le nom d’utilisateur GitHub des champs HTTPS ne remplace pas cette identité. |
| Push direct vers `main` refusé | La branche est probablement protégée. Publier sur une nouvelle branche et créer une Pull request. |
| Push refusé après création du commit | Le commit reste dans le dépôt local. Vérifier l’accès et la branche ; si l’interface n’affiche plus de fichiers modifiés, demander de l’aide au référent pour transmettre le commit local. Ne pas utiliser `push --force`. |
| Modifications déjà indexées | L’application bloque la publication pour éviter de mélanger des fichiers. Demander au référent de vérifier l’index avant de réessayer. |

Pour les étapes détaillées destinées aux nouveaux utilisateurs, y compris la création du jeton et la Pull request, consulter [le guide Git intégré](./ROADMAP.md).

## Dépendances et tests

`requirements.txt` contient les dépendances Python de l’application, des outils métier existants et des tests. **Git est un programme système, pas une bibliothèque Python** : il doit être installé séparément et accessible depuis `PATH`.

Exécuter la suite de tests depuis la racine du projet :

```bash
python -m unittest discover -s tests
```

Sur Linux sans écran :

```bash
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests
```

## Conception et sécurité

La description de l’architecture, des composants Git, des flux de données, des limites connues et des décisions de sécurité est documentée dans [SDD.md](./SDD.md) (Software Design Description).
