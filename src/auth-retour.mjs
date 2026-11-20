import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';

broadcastResponseToMainFrame().catch(() => {
  document.querySelector('#auth-retour-message').textContent =
    'Le retour de la connexion Microsoft a expiré. Rouvrez le portail pour vous reconnecter.';
  document.querySelector('#auth-retour-link').hidden = false;
});
