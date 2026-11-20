# Base de connaissances IRESEN

Le portail fonctionne dans un navigateur, indépendamment de ChatGPT. Les collègues se connectent avec leur compte professionnel IRESEN. Le portail lit le classeur maître et les dossiers dans SharePoint avec les autorisations du compte connecté. Les fiches restent consultables sur le site tant que ce compte peut lire le classeur. Les liens vers les pièces sources ouvrent SharePoint, qui contrôle l’accès à chaque document.

GitHub Pages sert l’interface de connexion et son programme. L’écran d’entrée est accessible à l’adresse du site. Les données et les documents restent dans SharePoint. La fenêtre d’identification est celle de Microsoft et les règles de sécurité définies par IRESEN s’appliquent à l’application.

## 1. Préparer le dépôt

Le dépôt public `sebbahis/iresen-connaissance` héberge le programme et les logos. Le classeur et les documents sont lus dans SharePoint après identification Microsoft. Choisir **GitHub Actions** comme source dans **Settings → Pages**.  
[Disponibilité et visibilité de GitHub Pages](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site)

## 2. Autoriser l’application dans Microsoft Entra

Pour l’administration Microsoft 365 d’IRESEN.

1. Enregistrer l’application **IRESEN – Base de connaissances**. Choisir les comptes de cet annuaire d’organisation uniquement.
2. Dans **Authentification**, ajouter la plateforme **Application monopage**. Pour le dépôt indiqué ci-dessus, enregistrer l’URI exacte `https://sebbahis.github.io/iresen-connaissance/auth-retour.html`. Utiliser l’URL publique réelle si le dépôt ou le domaine change.
3. Ajouter la permission **déléguée** Microsoft Graph `Sites.Selected` et accorder le consentement administrateur. Accorder à l’application le rôle `read` sur le site **Dir. ParMI** via la gestion des autorisations de site. Le consentement Entra et l’autorisation du site sont tous deux nécessaires.
4. Dans **Applications d’entreprise**, exiger l’affectation. Affecter les comptes de collaborateurs IRESEN concernés ou le groupe interne existant qui les réunit. SharePoint continue de vérifier les droits de chaque compte.
5. Communiquer l’**ID de l’application** à la personne qui configure le dépôt. Communiquer également l’ID de l’annuaire Microsoft Entra. L’application monopage utilise le flux d’autorisation PKCE et fonctionne sans secret client.

[Configuration MSAL et page de retour](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/login-user) · [Autorisations sélectionnées dans Microsoft Graph](https://learn.microsoft.com/en-us/graph/permissions-selected-overview) · [Affectation des utilisateurs](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/assign-user-or-group-access-portal)

Pour attribuer la lecture au portail, l’administration utilise un client de gestion déjà autorisé pour la création de permissions SharePoint.

```http
POST https://graph.microsoft.com/v1.0/sites/<ID du site SharePoint>/permissions
Content-Type: application/json

{
  "roles": ["read"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "<ID de l’application IRESEN – Base de connaissances>",
        "displayName": "IRESEN – Base de connaissances"
      }
    }
  ]
}
```

[Création d’une permission de site et rôle requis](https://learn.microsoft.com/en-us/graph/api/site-post-permissions?view=graph-rest-1.0)

## 3. Configurer et publier

Dans **Settings → Secrets and variables → Actions → Variables**, créer `IRESEN_CLIENT_ID` avec l’ID de l’application transmis par l’administration Microsoft 365. Créer également les variables `IRESEN_TENANT_ID`, `IRESEN_DRIVE_ID`, `IRESEN_WORKBOOK_ID` et `IRESEN_FOLDER_ID` avec les références internes du portail. Ces valeurs sont à renseigner dans les réglages du dépôt.

Ouvrir **Actions → Publier l’interface sur GitHub Pages → Run workflow**. Le workflow compile et teste le site, puis publie uniquement le dossier `dist/`. Les modifications apportées au classeur dans SharePoint apparaissent à la prochaine ouverture du portail ou avec **Actualiser les données**.

## 4. Vérifier le service

- Ouvrir le portail sans session Microsoft et vérifier que l’écran de connexion apparaît.
- Utiliser un compte affecté qui a accès au classeur. Vérifier le décompte de la version de référence, soit 70 projets et 35 laboratoires, puis ouvrir une fiche GEP et un document.
- Utiliser un compte affecté sans accès au classeur. Vérifier que la lecture est refusée et que les fiches restent masquées.
- Ouvrir un document soumis à une autorisation particulière. Vérifier que SharePoint applique les droits de ce compte. Toute personne autorisée à lire le classeur peut consulter les fiches qu’il contient.


## Vérification locale

Avec Node.js 24, exécuter les commandes ci-dessous depuis le dossier du portail.

```sh
npm ci
npm run build
npm test
```

Servir le dossier `dist/` avec un serveur HTTP local pour vérifier l’affichage. Une compilation sans identifiants affiche l’écran de configuration. Dans le workflow GitHub Pages, la compilation s’arrête si l’ID de l’application Entra manque.
