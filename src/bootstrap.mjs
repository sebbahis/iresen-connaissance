import { InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
import { SharePointError, SharePointReader } from './sharepoint.mjs';
import { readWorkbook } from './workbook.mjs';
import { mountPortal } from '../web/app.js';

const config = __PORTAL_SETTINGS__;
const screen = document.querySelector('#auth-screen');
const portal = document.querySelector('#portal');
const message = document.querySelector('#auth-message');
const login = document.querySelector('#microsoft-login');
const retry = document.querySelector('#retry-read');
const signOut = document.querySelector('#sign-out');
const scopes = ['https://graph.microsoft.com/Sites.Selected'];
const page = new URL('./', document.baseURI);
let microsoft;
let account;
let stage = 'initialisation';

function showGate(text, { signIn = false, reread = false } = {}) {
  portal.hidden = true;
  screen.hidden = false;
  message.textContent = text;
  login.hidden = !signIn;
  login.disabled = false;
  retry.hidden = !reread;
  retry.disabled = false;
}

function tenantAccount(candidate) {
  return !!candidate && candidate.tenantId.toLowerCase() === config.tenantId.toLowerCase();
}

async function token() {
  stage = 'autorisation_microsoft';
  const result = await microsoft.acquireTokenSilent({ account, scopes });
  if (!tenantAccount(result.account) || !result.accessToken) throw new Error('Compte Microsoft non autorisé.');
  stage = 'lecture_sharepoint';
  return result.accessToken;
}

function loginError(error) {
  if (/AADSTS50105/.test(error.message)) {
    showGate('Votre compte doit être affecté au portail dans l’annuaire IRESEN.', { signIn: true });
  } else if (/AADSTS65001|AADSTS90094|consent_required/.test(error.message)) {
    showGate('L’application doit être autorisée par l’administration Microsoft 365 d’IRESEN.', { signIn: true });
  } else if (error instanceof InteractionRequiredAuthError || error.status === 401) {
    showGate('Reconnectez-vous avec votre compte professionnel IRESEN pour renouveler la session.', { signIn: true });
  } else if (error instanceof SharePointError) {
    showGate(error.message, { signIn: true, reread: true });
  } else {
    showGate('L’accès au portail n’a pas abouti. Vérifiez votre connexion ou réessayez avec votre compte IRESEN.', { signIn: true, reread: !!account });
  }
  const codes = [error?.errorCode, error?.subError, error?.name]
    .filter(value => typeof value === 'string' && /^[a-zA-Z0-9_]{1,90}$/.test(value));
  const microsoftCode = /AADSTS\d{4,9}/.exec(error?.message ?? '')?.[0];
  if (microsoftCode) codes.unshift(microsoftCode);
  const typeDetail = /^(Cannot (?:read|set) properties of (?:undefined|null) \((?:reading|setting) '[a-zA-Z0-9_]{1,60}'\)|[a-zA-Z0-9_.]{1,80} is not a function|Failed to fetch)$/.exec(error?.message ?? '')?.[0];
  if (typeDetail) codes.push(typeDetail);
  message.textContent += ' Référence de diagnostic ' + stage + ' / ' + (codes.join(' / ') || 'erreur_inconnue') + '.';
}

async function openPortal() {
  showGate('Vérification de vos droits et lecture du classeur SharePoint…');
  try {
    const reader = new SharePointReader({ ...config, getToken: token });
    stage = 'lecture_sharepoint';
    const file = await reader.workbook(AbortSignal.timeout(35_000));
    let snapshot;
    try {
      stage = 'lecture_classeur';
      snapshot = await readWorkbook(file.bytes, { ...file, auditDate: config.auditDate });
    } catch {
      showGate('Le classeur est accessible, mais sa structure ne correspond plus aux tableaux du portail.', { reread: true });
      return;
    }
    stage = 'affichage_portail';
    mountPortal({
      snapshot,
      spreadsheetUrl: file.url,
      accountKey: account.localAccountId,
      readFolder: path => reader.listFolder(path, AbortSignal.timeout(30_000))
    });
    screen.hidden = true;
    portal.hidden = false;
  } catch (error) {
    loginError(error);
  }
}

async function start() {
  if (!config.tenantId || !config.clientId) {
    showGate('Le portail est prêt à être relié à l’annuaire Microsoft d’IRESEN. L’administrateur doit d’abord enregistrer l’application.');
    return;
  }
  try {
    microsoft = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: 'https://login.microsoftonline.com/' + config.tenantId,
        redirectUri: new URL('auth-retour.html', page).href,
        postLogoutRedirectUri: page.href
      },
      cache: { cacheLocation: 'sessionStorage' }
    });
    await microsoft.initialize();
    const response = await microsoft.handleRedirectPromise();
    const active = response?.account ?? microsoft.getActiveAccount();
    const eligible = microsoft.getAllAccounts().filter(tenantAccount);
    account = tenantAccount(active) ? active : eligible.length === 1 ? eligible[0] : undefined;
    if (!account) {
      showGate('Identifiez-vous avec votre compte professionnel @iresen.ma.', { signIn: true });
      return;
    }
    microsoft.setActiveAccount(account);
    stage = 'session_connectee';
    await openPortal();
  } catch (error) {
    loginError(error);
  }
}

login.addEventListener('click', async () => {
  login.disabled = true;
  message.textContent = 'Ouverture de la connexion Microsoft…';
  try {
    stage = 'connexion_microsoft';
    await microsoft.loginRedirect({ scopes, prompt: 'select_account', redirectStartPage: location.href });
  } catch (error) {
    loginError(error);
  }
});

retry.addEventListener('click', openPortal);
signOut.addEventListener('click', async () => {
  portal.hidden = true;
  if (document.querySelector('#detail').open) document.querySelector('#detail').close();
  showGate('Fermeture de la session du portail…');
  await microsoft.clearCache({ account });
  location.replace(page.href);
});

window.addEventListener('pageshow', event => {
  if (event.persisted && account) location.reload();
});
void start();
