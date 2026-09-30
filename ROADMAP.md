# Ajouter un outil au dépôt Git

Ce guide part du principe que le dépôt existe déjà et que la nouvelle recrue y a accès. Remplacer l’URL, l’organisation et `main` par les valeurs communiquées par l’équipe si elles sont différentes.

## 1. Cloner le dépôt

Récupérer l’URL de clonage dans GitHub, GitLab ou auprès du référent, puis ouvrir un terminal :

```bash
git clone https://<serveur>/<organisation>/<depot>.git
cd <depot>
```

Vérifier que Git reconnaît le dépôt et identifier la branche principale :

```bash
git status
git branch --show-current
```

## 2. Créer une branche de travail

Partir de la dernière version de la branche principale, puis créer une branche dédiée :

```bash
git switch main
git pull --ff-only origin main
git switch -c feature/ajout-nom-outil
```

Exemples : `feature/ajout-convertisseur-pdf`, `fix/correction-export-excel`, `docs/description-macro-fmd`. Une tâche distincte mérite sa propre branche.

## 3. Ajouter et vérifier le nouvel outil

- Placer le `.bas` ou le `.py` dans le dossier métier correspondant ; éviter de dupliquer un outil existant.
- Ajouter les dépendances Python nécessaires dans `requirements.txt` et documenter les versions requises.
- Ajouter ou corriger la description de l’outil, son utilité, ses prérequis, ses entrées et ses sorties.
- Pour une macro, noter les feuilles, colonnes, chemins et procédures appelées. Tester uniquement sur un classeur d’essai non sensible.
- Ne jamais ajouter de données personnelles, de classeur de production, de secret, de fichier temporaire ou de résultat généré.
- Lancer les tests et vérifier le contenu du changement :

```bash
python -m unittest discover -s tests
git status --short
git diff
```

## 4. Enregistrer et envoyer le changement

Ajouter explicitement les fichiers utiles, relire la liste préparée, puis créer un commit clair :

```bash
git add chemin/du/module.bas chemin/vers/les-tests.py
git diff --cached
git commit -m "feat: ajouter nom de l'outil"
git push -u origin feature/ajout-nom-outil
```

Ajouter les autres fichiers concernés explicitement à `git add`. Ne pas valider de fichier sensible ou de sortie de test.

## 5. Ouvrir une pull request

Dans GitHub/GitLab, ouvrir une pull request de `feature/ajout-nom-outil` vers `main` et demander une revue au référent métier et à un relecteur technique. Décrire le besoin, les fichiers changés, les tests exécutés, les dépendances et les risques métier.

Après les retours, créer un commit et pousser les corrections sur la même branche :

```bash
git add chemin/des/fichiers-modifies
git commit -m "fix: prendre en compte les retours de revue"
git push
```

Attendre les validations avant la fusion et suivre les règles d’approbation de l’équipe.

## 6. Reprendre le travail après fusion

Une fois la pull request fusionnée, actualiser la branche principale locale et supprimer l’ancienne branche :

```bash
git switch main
git pull --ff-only origin main
git branch -d feature/ajout-nom-outil
```

## Règles à retenir

- Une branche et une pull request par tâche ; ne pas travailler directement sur `main`.
- Ne pas utiliser `git push --force` sur une branche partagée sans accord du référent.
- En cas de conflit ou de doute avant une fusion, demander une revue plutôt que supprimer ou réinitialiser des changements.
- Une macro n’est prête qu’après revue du code et essai sur des données non sensibles.