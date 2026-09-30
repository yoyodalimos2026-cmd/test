# Bibliothèque d’outils métier

Application de bureau pour parcourir les modules VBA (`.bas`) et scripts Python du workspace, comprendre leurs traitements, lancer les scripts Python et consulter des classeurs Excel.

## Prérequis

- Python 3.10 ou plus récent.
- Un environnement graphique pris en charge par Qt 6. Le projet est prévu et vérifié sous Linux ; l’exécution de macros est réservée à Windows avec Microsoft Excel.
- Pour les macros : Microsoft Excel de bureau, `pywin32` et l’autorisation d’accès au projet VBA décrite dans [Limites Excel](#limites-excel).

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

Si PowerShell interdit l’activation du script, utiliser `cmd.exe` et lancer `.venv\Scripts\activate.bat`. La dépendance `pywin32` est installée par `requirements.txt` uniquement sous Windows.

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

### Consulter un classeur Excel

1. Ouvrir la page **Classeur Excel**.
2. Cliquer sur **Parcourir…** et sélectionner un fichier `.xlsx`, `.xlsm`, `.xltx` ou `.xltm`.
3. Parcourir les onglets pour voir l’aperçu en lecture seule.

La limite de **30 lignes et 16 colonnes** concerne uniquement l’aperçu affiché dans l’application ; elle ne limite pas le classeur utilisé pour une macro. La lecture se fait en tâche de fond : l’interface reste disponible et indique l’ouverture puis la feuille en cours. L’indicateur est animé sans pourcentage, car openpyxl ne fournit pas de progression fiable pendant l’ouverture du gros fichier. L’affichage utilise un modèle Qt léger au lieu de créer un objet par cellule. En cas d’erreur de lecture, le statut et une boîte de dialogue donnent le détail.

L’aperçu montre les valeurs ou formules stockées, mais ne recalcule pas les formules. Il ne reproduit ni la mise en forme, ni les graphiques, ni le comportement des macros. La sélection utilise le dialogue Qt non natif pour éviter les problèmes du sélecteur GTK dans certains environnements Linux/Snap.

### Appliquer une macro VBA

1. Charger un classeur depuis la page **Classeur Excel**.
2. Choisir un module VBA dans **Module à appliquer**.
3. Choisir une procédure publique dans la liste voisine.
4. Cliquer sur **Appliquer à une copie**, choisir le nom et le dossier de sortie du `.xlsm`, puis confirmer l’exécution.

Sous Windows, l’application crée d’abord une copie de travail temporaire à côté de l’original pour que les chemins relatifs utilisés par la macro continuent de pointer vers le dossier source pendant son exécution. Excel enregistre ensuite le **classeur complet** au format `.xlsm` à l’emplacement choisi. Toutes les feuilles et le contenu du classeur sont traités par Excel ; le plafond de l’aperçu ne s’applique pas. Le fichier original n’est pas remplacé. Si la destination existe déjà, l’application demande un autre nom pour éviter tout écrasement. Les modules `.bas` et `.frm` du dossier du module sélectionné sont importés dans la copie, puis la procédure choisie est exécutée dans Excel.

Sous Linux, le bouton indique que Windows et Excel sont nécessaires. Il ne lance pas de macro. LibreOffice n’est pas utilisé comme moteur VBA de remplacement. Consultez le guide [LANCEMENT_WINDOWS.md](LANCEMENT_WINDOWS.md) pour les étapes de configuration et de recette.

L’exécution est faite dans un worker Qt : l’application affiche l’étape en cours (ouverture, préparation de la copie, import des modules, exécution et sauvegarde) et reste réactive. La macro peut toutefois prendre du temps sans permettre de calculer un pourcentage : Excel exécute le code VBA sans fournir une progression générique. Une erreur s’affiche dans le statut et dans une boîte de dialogue détaillée.

## Technologies utilisées

- **Python 3.10+** : application et logique d’inventaire.
- **PySide6 / Qt 6** : fenêtres, navigation, listes, aperçu et interactions.
- **`ast` et expressions régulières** : découverte des procédures et analyse indicative des opérations Python/VBA.
- **openpyxl** : lecture en mode aperçu des classeurs Excel.
- **pandas et pdfplumber** : dépendances des outils Python existants de traitement de fichiers PDF et Excel.
- **pywin32 / COM Excel** : pilotage de Microsoft Excel pour l’import des modules et l’exécution VBA, uniquement sous Windows.
- **`unittest`** : tests automatisés du catalogue, de l’aperçu et des comportements Qt pris en charge.

## Limites Excel et sécurité

- L’exécution VBA exige **Windows et Microsoft Excel de bureau**. Le classeur est ouvert et piloté par Excel via COM ; la compatibilité des macros avec LibreOffice n’est ni fournie ni garantie.
- Dans Excel, autoriser l’accès au modèle d’objet du projet VBA : **Fichier → Options → Centre de gestion de la confidentialité → Paramètres du Centre de gestion de la confidentialité → Paramètres des macros → Accès approuvé au modèle d’objet du projet VBA**. Les intitulés peuvent varier selon la version d’Office et la langue.
- Cette autorisation permet à un programme d’importer du code dans un projet VBA. Ne l’activer que conformément aux règles de sécurité de l’organisation et n’exécuter que du code relu et autorisé.
- Excel enregistre une copie temporaire complète au format `.xlsm` à côté de la source, exécute la macro dans cette copie, puis enregistre le résultat complet à l’emplacement choisi. Cela ne limite pas les lignes aux seules données d’aperçu. Une conversion depuis `.xlsx` vers `.xlsm` change le format du fichier de sortie ; vérifier le résultat dans Excel. Une fois ce fichier déplacé vers un autre dossier, ses éventuelles références relatives peuvent pointer ailleurs.
- La copie protège le classeur original des modifications directes effectuées par la macro ; elle ne constitue **pas un bac à sable**. Une macro peut encore renommer, créer ou supprimer des fichiers, lancer un programme externe, accéder à un partage ou modifier d’autres classeurs. Sauvegarder aussi les fichiers concernés et tester avec des données non sensibles.
- Les modules exportés peuvent dépendre d’autres fichiers, références Excel, feuilles, colonnes, chemins ou programmes externes. L’application importe les `.bas` et `.frm` du dossier sélectionné, mais ne découvre pas ni ne copie toutes les dépendances externes ou tous les appels VBA indirects. Une macro peut échouer si le classeur choisi n’a pas la structure attendue.
- Le catalogue identifie principalement les déclarations `Sub`/`Function` et les appels VBA écrits avec `Call`. Une procédure détectée n’est pas nécessairement sûre ni autonome ; les résumés générés sont indicatifs.
- L’aperçu Excel ne prend pas en charge les anciens formats binaires `.xls` ou `.xlsb`, ne calcule pas les formules et n’évalue pas le contenu des macros. Les classeurs chiffrés ou protégés peuvent ne pas être lisibles ou modifiables.
- Les scripts Python sont lancés tels quels avec les droits du compte utilisateur. Ils peuvent eux aussi modifier des fichiers, ouvrir des fenêtres ou nécessiter des dépendances et ressources supplémentaires.

## Tests

Depuis la racine du dépôt :

```bash
python -m unittest discover -s tests
```

Le guide de contribution pour ajouter un outil à un dépôt Git existant se trouve dans [ROADMAP.md](ROADMAP.md).