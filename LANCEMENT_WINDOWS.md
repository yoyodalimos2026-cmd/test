# Lancer l’application sous Windows

Ce guide décrit l’installation, la configuration d’Excel et le premier essai d’une macro VBA. Prévoir un poste Windows avec Microsoft Excel de bureau et l’autorisation de l’équipe informatique pour les réglages de sécurité.

## 1. Installer les prérequis

Installer Python 3.10 ou plus récent (version Windows 64 bits recommandée), Git et Microsoft Excel de bureau. Excel en ligne ou LibreOffice ne remplace pas Excel COM.

Vérifier dans PowerShell :

```powershell
py -3 --version
git --version
```

Si le projet est déjà cloné, passer à l’étape 3. Sinon, récupérer l’URL du dépôt auprès de l’équipe :

```powershell
git clone https://<serveur>/<organisation>/<depot>.git
Set-Location <depot>
```

## 2. Créer l’environnement Python

Depuis la racine du projet :

```powershell
py -3 -m venv .venv
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

La modification de stratégie ci-dessus ne concerne que la fenêtre PowerShell courante. Si la politique de l’entreprise l’interdit, ouvrir **Invite de commandes (`cmd.exe`)** dans le projet et utiliser :

```bat
py -3 -m venv .venv
.venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## 3. Autoriser l’automatisation VBA dans Excel

La macro choisie est importée par l’application dans une copie du classeur. Excel doit autoriser l’accès programmatique au projet VBA :

1. Ouvrir Excel de bureau.
2. Aller à **Fichier → Options → Centre de gestion de la confidentialité → Paramètres du Centre de gestion de la confidentialité…**.
3. Ouvrir **Paramètres des macros**.
4. Activer **Accès approuvé au modèle d’objet du projet VBA**.
5. Valider les fenêtres de réglage et fermer Excel avant de lancer la première macro.

Les noms de menus peuvent changer selon la version ou la langue d’Office. Si le réglage est grisé ou interdit par une stratégie d’entreprise, demander l’aide de l’équipe informatique plutôt que de contourner cette règle. N’autoriser et n’exécuter que du code relu et approuvé.

## 4. Lancer l’application

Dans PowerShell, depuis la racine du projet et avec l’environnement virtuel activé :

```powershell
python main.py
```

La fenêtre **Bibliothèque d’outils** s’ouvre. Garder PowerShell ouvert pendant l’utilisation ; les erreurs de démarrage et la sortie des scripts peuvent y apparaître.

## 5. Charger et prévisualiser un classeur

1. Ouvrir **Classeur Excel** dans le menu à gauche.
2. Cliquer sur **Parcourir…** et choisir un `.xlsx` ou `.xlsm` (les modèles `.xltx`/`.xltm` sont également acceptés pour l’aperçu).
3. Parcourir les onglets affichés.

L’aperçu est limité aux 30 premières lignes et 16 premières colonnes de chaque feuille. Cette limite concerne uniquement la prévisualisation ; Excel travaille sur le classeur complet lors de l’application d’une macro. Le chargement de l’aperçu est effectué en tâche de fond : un indicateur animé et le nom de la feuille en cours s’affichent sans bloquer la fenêtre. Aucun pourcentage exact n’est affiché, car le lecteur Excel ne fournit pas de progression fiable pendant l’ouverture.

## 6. Créer le classeur complet avec la macro

1. Dans **Module à appliquer**, sélectionner le module VBA voulu.
2. Dans la liste voisine, sélectionner la procédure publique à exécuter.
3. Cliquer sur **Appliquer à une copie**.
4. Choisir un dossier de sortie et un nom se terminant par `.xlsm`, puis valider le sélecteur.
5. Lire l’avertissement sur les effets de la macro et confirmer uniquement si le code et les données ont été vérifiés.
6. Suivre l’étape en cours dans la zone d’état : démarrage d’Excel, ouverture du classeur, import des modules, exécution et sauvegarde. Le pourcentage n’est pas disponible pour une macro arbitraire.
7. En cas d’échec, lire le message détaillé affiché dans l’application et vérifier la sortie de PowerShell.
8. Ouvrir le `.xlsm` créé dans Excel et vérifier les feuilles, données et résultats attendus.

Le `.xlsm` est un classeur complet enregistré par Excel, pas un export limité à l’aperçu. L’original n’est pas utilisé comme destination. Pour éviter d’effacer des données, une destination déjà existante est refusée : choisir un autre nom. Pendant l’exécution, la copie de travail temporaire reste près du classeur d’origine afin de respecter les chemins relatifs employés par certaines macros ; elle est supprimée après le traitement.

## 7. Vérifier le comportement métier avant production

Faire un premier essai sur une copie de classeur non sensible. Comparer le résultat à celui attendu par le référent métier. Vérifier aussi les dossiers référencés par la macro et la présence des utilitaires externes éventuels, par exemple XPDF.

L’application ne rend pas les macros inoffensives : une macro peut renommer ou supprimer des fichiers, démarrer un programme externe, accéder à un partage ou modifier d’autres classeurs. La copie protège le classeur original des modifications directes, pas ces autres effets. Éviter les dossiers de production et les fichiers partagés tant que le scénario n’a pas été validé.

## Dépannage

- **Accès au projet VBA refusé** : vérifier le réglage du Centre de gestion de la confidentialité et les règles de l’entreprise.
- **Excel introuvable ou COM indisponible** : installer l’application de bureau Excel, activer `.venv`, puis réinstaller les dépendances avec `python -m pip install -r requirements.txt`.
- **La macro ne trouve pas une feuille, un fichier ou un dossier** : vérifier le classeur, les chemins et les ressources externes. Le code est exécuté dans une copie de travail temporaire placée à côté de l’original.
- **La sortie existe déjà** : choisir un nouveau nom `.xlsm` ; l’application ne l’écrase pas.
- **Erreur Excel pendant le traitement** : vérifier d’abord le message complet, les références VBA et les feuilles attendues, puis reproduire sur un jeu de test non sensible.

## Limite de validation

Les tests automatisés du projet vérifient le choix du chemin, la conservation de l’original et les appels COM à l’aide d’un Excel simulé. Ils ne remplacent pas un essai de recette sur Windows avec la version d’Excel, les classeurs et les dépendances externes réellement utilisés par l’équipe.