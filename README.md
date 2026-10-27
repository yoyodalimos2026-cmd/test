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

Ce guide accompagne l’installation sur **Windows** et **Linux**, y compris si vous n’avez pas l’habitude d’utiliser un terminal. Il faut installer Python et Git, télécharger le projet avec Git, puis installer les bibliothèques dont l’application a besoin.

> Les commandes ci-dessous sont à copier dans le terminal indiqué. Pour utiliser l’application, vous n’avez pas besoin de connaître les commandes Git : seules quelques commandes de préparation sont nécessaires.

### Étape 1 — Installer Python

L’application nécessite **Python 3.10 ou plus récent**. Si vous devez choisir une version, installez **Python 3.12 en 64 bits**.

#### Windows

1. Ouvrez la [page officielle de téléchargement de Python pour Windows](https://www.python.org/downloads/windows/).
2. Téléchargez l’installateur 64 bits pour Python 3.12 et ouvrez le fichier téléchargé.
3. Dans la première fenêtre, cochez **Add python.exe to PATH** si cette option est proposée, puis cliquez sur **Install Now**.
4. Une fois l’installation terminée, ouvrez **PowerShell** depuis le menu Démarrer.
5. Vérifiez Python en copiant cette commande puis en appuyant sur Entrée :

```powershell
py -3.12 --version
```

Le résultat doit commencer par `Python 3.12`. Si `py` n’est pas reconnu, relancez l’installateur et vérifiez que le lanceur Python est installé, ou demandez de l’aide à votre support informatique.

#### Linux

Ouvrez l’application **Terminal** et vérifiez si Python est déjà installé :

```bash
python3 --version
```

La version affichée doit être `3.10` ou plus récente. Si Python est absent ou plus ancien, installez une version récente depuis le magasin de logiciels de votre distribution ou demandez l’aide de votre support informatique. Il faut aussi installer le composant `venv`, utilisé pour isoler les bibliothèques de ce projet.

**Ubuntu ou Debian :**

```bash
sudo apt update
sudo apt install python3 python3-venv python3-pip
```

Le mot de passe demandé est celui de votre session Linux. Pendant sa saisie dans le terminal, aucun caractère ne s’affiche : saisissez-le puis appuyez sur Entrée.

**Fedora :**

```bash
sudo dnf install python3 python3-pip
```

Sur une autre distribution Linux, utilisez son magasin de logiciels pour installer Python 3.10 ou plus récent et le module `venv`.

### Étape 2 — Installer Git

Git est un programme séparé de Python. Il est nécessaire pour que l’application puisse gérer les branches et publier les modifications.

#### Windows

1. Téléchargez Git depuis le [site officiel Git pour Windows](https://git-scm.com/download/win).
2. Ouvrez l’installateur et conservez les choix proposés par défaut, sauf indication différente de votre équipe informatique.
3. Ouvrez **PowerShell** depuis le menu Démarrer.
4. Vérifiez Git :

```powershell
git --version
```

Le résultat doit ressembler à `git version 2.x.x`. Si PowerShell indique que `git` n’est pas reconnu, fermez puis rouvrez PowerShell. Si le problème persiste, redémarrez Windows ou contactez le support informatique.

#### Linux

Ouvrez **Terminal** et installez Git avec le gestionnaire de paquets de votre distribution.

**Ubuntu ou Debian :**

```bash
sudo apt update
sudo apt install git
```

**Fedora :**

```bash
sudo dnf install git
```

Vérifiez l’installation :

```bash
git --version
```

### Étape 3 — Télécharger le projet avec Git

Sur la page GitHub du projet, cliquez sur le bouton vert **Code**, puis copiez l’adresse affichée sous l’onglet **HTTPS**. Elle ressemble à `https://github.com/organisation/nom-du-projet.git`.

Ouvrez **PowerShell** sous Windows ou **Terminal** sous Linux et copiez la commande suivante en remplaçant l’adresse d’exemple par celle que vous venez de copier :

```text
git clone https://github.com/organisation/nom-du-projet.git
```

Appuyez sur Entrée et attendez la fin du téléchargement. Un nouveau dossier portant le nom du projet est créé. Entrez dans ce dossier avec `cd`, en remplaçant `nom-du-projet` par son nom réel :

```text
cd nom-du-projet
```

Toutes les étapes ci-dessous doivent être réalisées dans ce dossier. Si vous avez téléchargé le projet avec le bouton **Download ZIP**, cette copie ne permet pas de publier des changements : clonez le dépôt avec Git dans un nouveau dossier.

### Étape 4 — Installer les prérequis Python et démarrer

Le fichier `requirements.txt` contient les bibliothèques nécessaires. Les commandes suivantes créent un environnement Python propre au projet et y installent les prérequis. Cette méthode ne change pas les bibliothèques des autres applications Python. Il n’est pas nécessaire d’activer l’environnement : les commandes utilisent directement Python installé dans `.venv`.

#### Windows — PowerShell

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

#### Linux — Terminal

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python main.py
```

La première installation peut prendre quelques minutes et nécessite une connexion Internet. Une fenêtre **Bibliothèque d’outils** doit s’ouvrir après la dernière commande.

Si l’installation de PySide6 échoue sous Linux en indiquant qu’une bibliothèque système Qt ou graphique manque, installez les mises à jour proposées par votre distribution ou contactez le support informatique. Certaines installations Linux minimales ne comprennent pas l’environnement graphique nécessaire.

### Relancer l’application

Après l’installation initiale, il n’est pas nécessaire de réinstaller les bibliothèques. Ouvrez un terminal dans le dossier du projet et exécutez :

**Windows — PowerShell :**

```powershell
.\.venv\Scripts\python.exe main.py
```

**Linux — Terminal :**

```bash
.venv/bin/python main.py
```

### Problèmes d’installation fréquents

| Problème | Solution |
|---|---|
| `git`, `py` ou `python3` n’est pas reconnu | Vérifiez l’installation de l’outil, puis fermez et rouvrez le terminal. Sous Windows, redémarrez l’ordinateur si nécessaire. |
| `No module named venv` ou échec de création de `.venv` sous Linux | Installez `python3-venv` sur Ubuntu/Debian, ou demandez au support informatique d’installer le module `venv` correspondant à Python. |
| `No module named ...` au lancement | Recommencez l’étape 4 depuis le dossier du projet avec les commandes de votre système. |
| Téléchargement ou installation des prérequis impossible | Vérifiez la connexion Internet. Un proxy ou une restriction de l’entreprise peut nécessiter l’aide du support informatique. |
| L’application ne s’ouvre pas | Vérifiez que le terminal se trouve dans le dossier du projet cloné et contactez le support informatique avec le message affiché dans le terminal. |

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
8. Si vous avez coché des fichiers, écrivez un résumé clair du changement, par exemple `docs: clarifier les instructions Git`. Si vous créez une branche sans publier de fichier, aucun résumé n’est nécessaire.
9. Cliquez sur **Publier / créer une branche** et relisez la confirmation avant de valider.
10. En cas de succès avec des fichiers cochés, le commit a été créé localement puis envoyé vers le dépôt distant. Si vous avez choisi **Nouvelle branche** sans cocher de fichier, la branche est publiée à partir du commit actuel sans créer de commit.

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

**Windows (PowerShell) :**

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

**Linux :**

```bash
.venv/bin/python -m unittest discover -s tests
```

Sur Linux sans écran graphique :

```bash
QT_QPA_PLATFORM=offscreen .venv/bin/python -m unittest discover -s tests
```

## Conception et sécurité

La description de l’architecture, des composants Git, des flux de données, des limites connues et des décisions de sécurité est documentée dans [SDD.md](./SDD.md) (Software Design Description).
