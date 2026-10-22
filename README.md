# Bibliothèque d’outils métier

Application de bureau pour parcourir et modifier les modules VBA (`.bas`) et scripts Python du workspace, comprendre leurs traitements, lancer les scripts Python et publier les changements avec Git.

## Prérequis

- Python 3.10 ou plus récent.
- Un environnement graphique pris en charge par Qt 6.
- Git installé et configuré, avec un accès au dépôt distant via Git Credential Manager ou une clé SSH.

## Installation et lancement

Depuis la racine du dépôt, créer un environnement virtuel et installer les dépendances :

### Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main.py
```

### Windows

```powershell
py -3 -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```

Si PowerShell interdit l’activation du script, utiliser `cmd.exe` et lancer `.venv\Scripts\activate.bat`.

## Utilisation

### Bibliothèque

La page **Bibliothèque** recense les fichiers `.bas` et `.py` sous les dossiers métier. Le champ de recherche accepte un nom de fichier, une procédure ou un dossier ; le filtre sépare VBA et Python.

En sélectionnant un élément, l’application affiche les procédures détectées ainsi que les fichiers du même dossier et les appels VBA directs reconnus. Sélectionner un fichier lié dans cette liste ouvre sa fiche.

La zone **Utilité** est un résumé automatique des opérations reconnues dans le code. Elle ne reprend pas les commentaires, docstrings, noms de procédures, noms de feuilles, noms de personnes ou dates ; elle peut donc rester générale si le traitement n’est pas reconnu. Les rapprochements automatiques ne remplacent pas une revue du code par un référent métier.

Le bouton **Ajouter une note métier** permet de saisir un complément validé par l’équipe. Ce texte est conservé séparément dans `module_descriptions.json` et ne modifie pas le résumé déduit du code. Cette note libre n’est pas anonymisée automatiquement : ne pas y saisir de données personnelles ou confidentielles.

### Lancer un script Python

1. Sélectionner un fichier Python dans la bibliothèque.
2. Cliquer sur **Lancer le script**.
3. Lire l’état dans la fiche. Les éventuelles fenêtres et sorties du script apparaissent séparément de la bibliothèque, selon le comportement propre du script.

Le script est lancé avec le même interpréteur Python que l’application ; son répertoire de travail est son dossier source. Certains scripts ouvrent leur propre interface, demandent des fichiers ou créent des classeurs et d’autres résultats. Le bouton lance le script tel quel : il ne transforme pas une ancienne interface Tkinter en interface Qt.

### Modifier une source et publier avec Git

1. Dans la **Bibliothèque**, sélectionner un fichier `.py` ou `.bas`, puis cliquer sur **Ouvrir le fichier source**.
2. Modifier le texte et cliquer sur **Enregistrer**. Le fichier est enregistré avec son encodage et ses fins de ligne.
3. Ouvrir **Gestion Git**, vérifier l’URL du dépôt distant et actualiser l’état.
4. Cocher uniquement les fichiers à publier, choisir une nouvelle branche ou `main`, écrire le message de commit, puis confirmer.

La publication configure le remote `origin`, crée un commit pour les seuls fichiers cochés et pousse vers la branche choisie. Les fichiers non cochés restent hors du commit. Les changements déjà indexés manuellement bloquent l’opération afin d’éviter de les inclure par erreur. La sélection de `main` nécessite que cette branche existe localement.

Pour un dépôt HTTPS, l’utilisateur peut saisir son nom de compte et son jeton dans les champs de la page. Ils ne sont pas sauvegardés et ne sont utilisés que pendant le push ; le champ du jeton est masqué et son fichier temporaire protégé est supprimé à la fin de l’opération. GitHub demande un jeton personnel (PAT) à la place du mot de passe du compte. Pour une URL SSH, laissez les champs vides et configurez la clé SSH dans Git. Sans identifiants saisis, la publication s’appuie sur la méthode d’authentification Git déjà configurée.

## Technologies utilisées

- **Python 3.10+** : application et logique d’inventaire.
- **PySide6 / Qt 6** : fenêtres, navigation, listes et interactions.
- **`ast` et expressions régulières** : découverte des procédures et analyse indicative des opérations Python/VBA.
- **pandas et pdfplumber** : dépendances des outils Python existants de traitement de fichiers PDF et Excel.
- **Git** : commit et publication des sources sélectionnées.
- **`unittest`** : tests automatisés du catalogue, de l’éditeur et des comportements Qt pris en charge.

## Sécurité

- La publication pousse les fichiers explicitement cochés, mais le contenu de ceux-ci doit être relu avant validation. La confirmation affiche les chemins concernés.
- Ne placez jamais de mot de passe ou de jeton dans l’URL du dépôt. L’application refuse les URL avec identifiants intégrés. Les jetons HTTPS saisis dans l’interface ne sont pas conservés après le push ; configurez plutôt une clé SSH ou le gestionnaire d’identifiants Git si vous souhaitez enregistrer votre authentification.
- Les scripts Python sont lancés tels quels avec les droits du compte utilisateur. Ils peuvent eux aussi modifier des fichiers, ouvrir des fenêtres ou nécessiter des dépendances et ressources supplémentaires.

## Tests

Depuis la racine du dépôt :

```bash
python -m unittest discover -s tests
```

Le guide de contribution pour ajouter un outil à un dépôt Git existant se trouve dans [ROADMAP.md](ROADMAP.md).