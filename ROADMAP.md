# Guide d’utilisation de la gestion Git

Ce guide explique comment préparer l’accès à un dépôt et publier des fichiers avec la page **Gestion Git**. Il s’adresse aux personnes qui ne connaissent pas encore Git.

## Avant de commencer

Il faut :

- avoir un compte sur la plateforme qui héberge le dépôt ;
- avoir au minimum le droit de lecture pour cloner le dépôt, et le droit d’écriture pour y pousser des changements ;
- installer Git sur l’ordinateur ;
- ouvrir une copie **clonée** du dépôt dans l’application.

> **Important :** télécharger le dépôt sous forme de fichier ZIP ne suffit pas. L’archive ne contient pas l’historique Git ; la page **Gestion Git** ne pourra pas y créer de commit ni pousser de modifications. Utilisez `git clone`.

## 1. Cloner le dépôt

Sur GitHub, ouvrir la page du dépôt, cliquer sur **Code**, puis copier l’URL HTTPS, par exemple :

```text
https://github.com/mon-organisation/mon-depot.git
```

Ouvrir un terminal dans le dossier où vous souhaitez placer le projet et lancer :

```bash
git clone https://github.com/mon-organisation/mon-depot.git
cd mon-depot
```

Lancer ensuite l’application depuis ce dossier :

```bash
python main.py
```

Sur Windows, si la commande `python` n’est pas disponible, utiliser :

```powershell
py main.py
```

## 2. Définir une fois l’identité des commits

Git inscrit le nom et l’adresse électronique de l’auteur dans chaque commit. Cette identité est différente du nom d’utilisateur et du jeton utilisés pour s’authentifier auprès de GitHub.

Configurer l’identité pour ce dépôt depuis son terminal :

```bash
git config --local user.name "Prénom Nom"
git config --local user.email "adresse-associée-au-compte@example.com"
```

Utiliser de préférence l’adresse liée au compte GitHub, ou l’adresse privée `noreply` fournie par GitHub si vous ne souhaitez pas exposer votre adresse personnelle dans les commits. Cette configuration est enregistrée dans le dépôt local ; elle ne donne aucun droit d’accès au dépôt distant.

## 3. Créer un jeton GitHub pour HTTPS

GitHub n’accepte pas le mot de passe du compte pour pousser par HTTPS. Il faut utiliser un **Personal Access Token (PAT)**. Pour créer un jeton à accès limité :

1. Connectez-vous au compte GitHub qui a accès au dépôt.
2. Ouvrez **Settings → Developer settings → Personal access tokens → Fine-grained tokens**, ou [la page de création](https://github.com/settings/personal-access-tokens/new).
3. Cliquez sur **Generate new token** si nécessaire, puis donnez-lui un nom explicite et une date d’expiration.
4. Dans **Resource owner**, sélectionnez le propriétaire du dépôt : votre compte ou l’organisation concernée.
5. Dans **Repository access**, choisissez **Only select repositories**, puis sélectionnez le dépôt à utiliser.
6. Dans **Repository permissions**, ajoutez **Contents** et réglez-le sur **Read and write**. Gardez les autres permissions au minimum nécessaire.
7. Générez le jeton et copiez-le immédiatement dans un gestionnaire sûr. GitHub ne permet généralement pas de le consulter à nouveau après avoir quitté la page.

Si l’organisation exige une approbation pour les jetons, demandez cette approbation à son administrateur. Le jeton ne donne pas à votre compte plus de droits que ceux qu’il possède déjà sur le dépôt.

> Ne communiquez jamais le jeton, ne le placez pas dans l’URL du dépôt et ne l’ajoutez pas à un fichier du projet.

## 4. Retrouver le nom d’utilisateur GitHub

Le champ **Nom d’utilisateur** attend le nom de compte, et non le nom complet ou l’adresse électronique. Il s’agit du texte après `@` dans votre profil ou de la partie après `github.com/` dans son URL.

Par exemple, pour `https://github.com/marie-exemple`, saisissez `marie-exemple`.

## 5. Modifier une source et publier

1. Dans la page **Bibliothèque**, sélectionnez un fichier Python (`.py`) ou VBA (`.bas`).
2. Cliquez sur **Ouvrir le fichier source**, effectuez la modification, puis cliquez sur **Enregistrer**.
3. Ouvrez **Gestion Git**. Vérifiez que **Dépôt distant** affiche la bonne URL HTTPS du dépôt. L’application utilise le remote `origin` et le met à jour avec l’URL saisie lors de la publication.
4. Si vous utilisez HTTPS, saisissez le nom d’utilisateur GitHub et collez le PAT dans le champ masqué. Pour une URL SSH, laissez les champs d’identifiants vides et utilisez votre clé SSH déjà configurée.
5. Cliquez sur **Actualiser** pour afficher les modifications détectées.
6. Cochez uniquement les fichiers que vous souhaitez envoyer. Vérifiez leurs chemins dans la liste.
7. Choisissez **Nouvelle branche** de préférence et donnez-lui un nom explicite (par exemple `feature/description-outil` ou `fix/correction-export`), ou sélectionnez une branche existante dans la liste. Les branches connues de `origin` sont affichées avec la mention « distante » ; l’application crée leur branche locale de suivi au moment de la publication. Choisissez `main` uniquement si votre équipe autorise les pushes directs.
8. Entrez un résumé clair du changement, par exemple `docs: expliquer la publication avec Git`.
9. Cliquez sur **Publier les fichiers cochés**, relisez la confirmation et confirmez.
10. Attendez le résultat affiché par l’application. En cas de succès, le commit et le push ont été effectués.

L’application crée un commit local avec le nom et l’adresse Git configurés à l’étape 2, puis pousse les fichiers cochés sur la branche sélectionnée. Les autres modifications restent non publiées. Il n’est pas nécessaire d’ajouter manuellement les fichiers dans Git avant cette opération.

## Changer de branche dans l’application

Il n’est pas nécessaire de fermer puis relancer Python pour changer de branche :

1. Dans **Gestion Git**, choisissez une branche existante dans la liste **Publication**.
2. Cliquez sur **Basculer vers cette branche** et confirmez.
3. La bibliothèque recharge immédiatement les fichiers de la branche choisie ; le statut indique la branche active.

Le changement est refusé si Git détecte des modifications locales ou des fichiers non suivis. Enregistrez et publiez-les avant de basculer. L’application ne les supprime pas et ne les met pas automatiquement de côté.

La liste comprend les branches locales et les références distantes `origin/*` déjà connues du clone. Pour voir les dernières branches du serveur, cliquez sur **Récupérer les branches distantes** : cette action télécharge les références dans l’application, sans terminal. Pour un dépôt HTTPS privé, saisissez d’abord votre nom d’utilisateur et votre jeton. **Actualiser** relit seulement l’état local Git.

## 6. Branches : laquelle choisir ?

- **Nouvelle branche** : choix recommandé. Elle part de la branche actuellement active sur votre copie locale, puis reçoit le commit sélectionné. Elle permet de demander une revue avant de fusionner le changement.
- **Branche existante** : permet de publier sur une branche locale ou une branche distante connue du clone. Une branche distante est marquée « distante » ; sa copie locale de suivi est créée automatiquement avant le commit.
- **`main`** : pousse directement sur la branche principale locale. Cette option peut être refusée par les règles de protection du dépôt ; ne l’utilisez que si votre équipe l’autorise.

Après un push sur une nouvelle branche, ouvrez la page du dépôt sur GitHub et créez une **Pull request** pour proposer le changement à la revue et à la fusion.

## 7. Comprendre les problèmes fréquents

### « Authentication failed », « Invalid username or token »

Vérifiez le nom de compte, la date d’expiration et les permissions du PAT. Vérifiez également qu’il couvre le bon dépôt et que le compte possède le droit d’écriture. Si nécessaire, révoquez le jeton compromis ou expiré et générez-en un nouveau.

### « Repository not found » ou accès refusé

Vérifiez l’URL complète, le propriétaire du dépôt, l’accès accordé à votre compte et le dépôt sélectionné dans le PAT. Les dépôts privés ne sont visibles que par les comptes autorisés.

### La branche `main` n’existe pas localement

Choisissez **Nouvelle branche**, ou demandez au référent de vous aider à récupérer la branche principale avec Git.

### La publication sur `main` est refusée

Le dépôt peut protéger cette branche ou interdire les pushes directs. Publiez sur une nouvelle branche, puis ouvrez une Pull request.

### Le commit est créé, mais le push échoue

Le commit reste dans votre copie locale ; il n’est pas perdu et n’a pas été transmis au serveur. Après avoir corrigé l’accès réseau ou l’authentification, vous pouvez renvoyer ce commit depuis un terminal :

```bash
git push -u origin nom-de-votre-branche
```

Remplacez `nom-de-votre-branche` par la branche affichée dans Git. Si le push a échoué car le serveur a reçu des changements plus récents ou s’il y a un conflit, ne forcez pas la publication : demandez de l’aide au référent.

### Aucun fichier n’apparaît

Cliquez sur **Actualiser**. Vérifiez ensuite que le fichier a été enregistré et qu’il se trouve bien dans le dépôt cloné ouvert par l’application.

### Git refuse de publier à cause de changements déjà indexés

La page s’arrête volontairement si des fichiers sont déjà préparés avec `git add` dans un terminal. Cela évite d’inclure des changements imprévus. Demandez au référent de vérifier l’index Git avant de recommencer.

## Bonnes pratiques

- Préférez une branche dédiée et une Pull request à un push direct sur `main`.
- Relisez chaque fichier coché avant la confirmation.
- Ne publiez jamais de jeton, mot de passe, donnée personnelle ou fichier métier confidentiel.
- Donnez au jeton une date d’expiration, limitez-le à un seul dépôt et révoquez-le s’il a été exposé.
- N’utilisez pas `git push --force` pour contourner un refus.
- Si un message d’erreur n’est pas clair, conservez-le et demandez l’aide du référent Git du projet.
