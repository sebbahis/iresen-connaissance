const GRAPH = 'https://graph.microsoft.com';
const FILE_FIELDS = 'id,name,webUrl,size,lastModifiedDateTime,parentReference,folder,file';
const MAX_WORKBOOK_BYTES = 15 * 1024 * 1024;
const GRAPH_ID = /^[a-zA-Z0-9!._-]{1,256}$/;

export class SharePointError extends Error {
  constructor(status, message) {
    super(message);
    this.name = 'SharePointError';
    this.status = status;
  }
}

function checkResponse(response, subject) {
  if (response.ok) return;
  if (response.status === 401) throw new SharePointError(401, 'La session Microsoft doit être renouvelée.');
  if (response.status === 403 || response.status === 404) {
    throw new SharePointError(response.status, 'Vérifiez que votre compte Microsoft peut ouvrir ' + subject + ' dans SharePoint.');
  }
  if (response.status === 429) throw new SharePointError(429, 'SharePoint traite trop de demandes. Réessayez dans un instant.');
  throw new SharePointError(response.status, 'La lecture de ' + subject + ' a échoué. Réessayez dans un instant.');
}

export function sharePointUrl(value) {
  try {
    const url = new URL(value);
    if (url.protocol === 'https:' && !url.port && !url.username && !url.password &&
        /^(rabatcommerce|rabatcommerce-my)\.sharepoint\.com$/i.test(url.hostname)) return url;
  } catch {
    return null;
  }
  return null;
}

function encodedPath(path) {
  const parts = path.split('/');
  if (path.length > 1500 || parts.some(part => !part || ['.', '..'].includes(part) || /[\\\x00-\x1f]/.test(part))) {
    throw new SharePointError(400, 'Le chemin du dossier est invalide.');
  }
  return parts.map(part => encodeURIComponent(part)).join('/');
}

async function limitedBuffer(response) {
  const announced = Number(response.headers.get('content-length'));
  if (announced > MAX_WORKBOOK_BYTES) throw new Error('Le classeur dépasse la taille prise en charge.');
  if (!response.body) throw new Error('SharePoint a renvoyé un fichier vide.');
  const reader = response.body.getReader();
  const chunks = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > MAX_WORKBOOK_BYTES) throw new Error('Le classeur dépasse la taille prise en charge.');
      chunks.push(value);
    }
  } finally {
    await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

export class SharePointReader {
  constructor({ driveId, workbookId, folderId, getToken, request = globalThis.fetch.bind(globalThis) }) {
    if (![driveId, workbookId, folderId].every(id => typeof id === 'string' && GRAPH_ID.test(id))) {
      throw new Error('Les références SharePoint du portail sont incomplètes.');
    }
    this.base = '/v1.0/drives/' + encodeURIComponent(driveId) + '/items/';
    this.workbookPath = this.base + encodeURIComponent(workbookId);
    this.folderPath = this.base + encodeURIComponent(folderId);
    this.getToken = getToken;
    this.request = request;
  }

  async readJson(link, subject, signal) {
    const url = new URL(link, GRAPH);
    if (url.origin !== GRAPH || !url.pathname.startsWith(this.base)) throw new Error('Réponse Graph inattendue.');
    const token = await this.getToken();
    const response = await this.request(url.href, {
      method: 'GET',
      headers: { Authorization: 'Bearer ' + token },
      credentials: 'omit',
      cache: 'no-store',
      redirect: 'error',
      referrerPolicy: 'no-referrer',
      signal
    });
    checkResponse(response, subject);
    return await response.json();
  }

  async workbook(signal) {
    const metadata = await this.readJson(this.workbookPath + '?$select=' + FILE_FIELDS + ',@microsoft.graph.downloadUrl', 'le classeur maître', signal);
    const web = sharePointUrl(metadata.webUrl);
    const download = sharePointUrl(metadata['@microsoft.graph.downloadUrl']);
    if (!web || !download || !metadata.file || !/\.xlsx$/i.test(metadata.name)) {
      throw new Error('SharePoint ne fournit pas un classeur Excel téléchargeable pour ce compte.');
    }
    if (metadata.size > MAX_WORKBOOK_BYTES) throw new Error('Le classeur dépasse la taille prise en charge.');
    // Graph renvoie une URL d’accès temporaire. Le jeton Graph reste sur graph.microsoft.com.
    const response = await this.request(download.href, {
      credentials: 'omit',
      cache: 'no-store',
      redirect: 'error',
      referrerPolicy: 'no-referrer',
      signal
    });
    checkResponse(response, 'le classeur maître');
    const bytes = await limitedBuffer(response);
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('Le contenu renvoyé par SharePoint n’est pas un classeur Excel.');
    return { bytes, name: metadata.name, modified: metadata.lastModifiedDateTime, url: web.href };
  }

  async listFolder(path, signal) {
    const itemPath = path ? this.folderPath + ':/' + encodedPath(path) + ':' : this.folderPath;
    const metadata = await this.readJson(itemPath + '?$select=id,webUrl,folder', 'ce dossier', signal);
    const web = sharePointUrl(metadata.webUrl);
    if (!metadata.folder || !web || !GRAPH_ID.test(metadata.id)) throw new Error('Ce dossier SharePoint est indisponible.');
    let next = this.base + encodeURIComponent(metadata.id) + '/children?$select=' + FILE_FIELDS + '&$top=200';
    const items = [];
    // Limiter le travail d’une lecture tout en suivant la pagination Graph.
    for (let page = 0; next && page < 10; page++) {
      const response = await this.readJson(next, 'ce dossier', signal);
      if (!Array.isArray(response.value)) throw new Error('Inventaire SharePoint invalide.');
      for (const item of response.value) {
        const url = sharePointUrl(item.webUrl);
        if (url && typeof item.name === 'string') {
          items.push({
            id: item.id,
            name: item.name,
            url: url.href,
            folder: !!item.folder,
            size: item.size,
            modified: item.lastModifiedDateTime
          });
        }
      }
      next = response['@odata.nextLink'];
    }
    return { url: web.href, items, partial: !!next, readAt: new Date().toISOString() };
  }
}
