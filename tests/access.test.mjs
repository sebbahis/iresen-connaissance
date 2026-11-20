import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import { SharePointReader } from '../src/sharepoint.mjs';
import { readWorkbook } from '../src/workbook.mjs';
import * as Core from '../web/core.js';

const transport = request => new SharePointReader({
  driveId: 'b!example',
  workbookId: 'workbook',
  folderId: 'folder',
  getToken: async () => 'test-token',
  request
});

test('le transport natif conserve le contexte du navigateur', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = function () {
    assert.equal(this, globalThis);
    return Promise.resolve(new Response(null, { status: 403 }));
  };
  try {
    const reader = new SharePointReader({
      driveId: 'b!example', workbookId: 'workbook', folderId: 'folder',
      getToken: async () => 'test-token'
    });
    await assert.rejects(reader.workbook(), { status: 403 });
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('un refus de lecture SharePoint arrête la requête avant tout téléchargement', async () => {
  let requests = 0;
  const reader = transport(async (url, init) => {
    requests++;
    assert.equal(new URL(url).origin, 'https://graph.microsoft.com');
    assert.equal(init.method, 'GET');
    assert.equal(init.headers.Authorization, 'Bearer test-token');
    return new Response(null, { status: 403 });
  });
  await assert.rejects(reader.workbook(), { status: 403 });
  assert.equal(requests, 1);
});

test('le fichier est lu depuis SharePoint sans transfert du jeton Graph', async () => {
  const workbookUrl = 'https://rabatcommerce.sharepoint.com/sites/Test/Doc.aspx';
  const workbook = new Uint8Array([0x50, 0x4b, 1, 2]);
  let requests = 0;
  const reader = transport(async (url, init) => {
    requests++;
    if (requests === 1) {
      assert.match(new URL(url).pathname, /\/items\/workbook$/);
      assert.equal(init.headers.Authorization, 'Bearer test-token');
      return Response.json({
        name: 'exemple.xlsx', file: {}, size: workbook.length,
        webUrl: workbookUrl, lastModifiedDateTime: '2026-11-20T12:00:00Z',
        '@microsoft.graph.downloadUrl': 'https://rabatcommerce.sharepoint.com/lecture-temporaire'
      });
    }
    assert.equal(new URL(url).hostname, 'rabatcommerce.sharepoint.com');
    assert.equal(init.headers, undefined);
    assert.equal(init.cache, 'no-store');
    assert.equal(init.credentials, 'omit');
    return new Response(workbook);
  });
  const file = await reader.workbook();
  assert.deepEqual(file.bytes, workbook);
  assert.equal(file.url, workbookUrl);
  assert.equal(Object.hasOwn(file, 'downloadUrl'), false);
});

test('les dossiers sont lus avec les mêmes droits et les pages suivantes restent sur Graph', async () => {
  const calls = [];
  const reader = transport(async (url, init) => {
    assert.equal(init.headers.Authorization, 'Bearer test-token');
    calls.push(url);
    if (calls.length === 1) return Response.json({
      id: 'sous-dossier', folder: {}, webUrl: 'https://rabatcommerce.sharepoint.com/site/dossier'
    });
    const doc = index => ({id: 'file-' + index, name: index + '.pdf', webUrl: 'https://rabatcommerce.sharepoint.com/site/' + index + '.pdf'});
    return Response.json(calls.length === 2
      ? {value: [doc(1)], '@odata.nextLink': url + '&$skiptoken=2'}
      : {value: [doc(2)]});
  });
  const result = await reader.listFolder('Données R&D/Archives');
  assert.equal(result.items.length, 2);
  assert.equal(result.partial, false);
  assert.ok(calls[0].includes('Donn%C3%A9es%20R%26D/Archives'));
  assert.ok(calls[2].endsWith('$skiptoken=2'));
});

test('la cible de téléchargement et le chemin de navigation sont vérifiés', async () => {
  let requests = 0;
  const reader = transport(async () => {
    requests++;
    return Response.json({
      file: {}, name: 'exemple.xlsx', webUrl: 'https://rabatcommerce.sharepoint.com/document',
      '@microsoft.graph.downloadUrl': 'https://outside.example/document'
    });
  });
  await assert.rejects(reader.workbook(), /classeur Excel téléchargeable/);
  await assert.rejects(reader.listFolder('../Autre bibliothèque'), {status: 400});
  assert.equal(requests, 1);
});

test('le paquet hébergé contient uniquement le programme et ses ressources', async () => {
  const entries = await readdir(new URL('../dist/', import.meta.url), {recursive: true});
  assert.ok(entries.includes('index.html') && entries.includes('auth-retour.html') && entries.includes('portail.js'));
  assert.ok(entries.every(path => !/\.xlsx$|\.zip$|(^|\/)data(\.js|\/)|base\.json|files\.json/.test(path)));
  const html = await readFile(new URL('../dist/index.html', import.meta.url), 'utf8');
  const program = await readFile(new URL('../dist/portail.js', import.meta.url), 'utf8');
  assert.match(html, /id="portal" hidden/);
  assert.match(html, /Content-Security-Policy/);
  assert.doesNotMatch(program, /IRESEN_DATA|PUB_550EDC282569|PRJ-0060|Périmètre des appels/);
});

test('lecture du classeur maître et reprise des fiches et liens', {
  skip: !(process.env.IRESEN_TEST_WORKBOOK && process.env.IRESEN_TEST_REFERENCE)
}, async () => {
  const book = await readFile(process.env.IRESEN_TEST_WORKBOOK);
  const known = JSON.parse(await readFile(process.env.IRESEN_TEST_REFERENCE, 'utf8'));
  const current = Core.validate(await readWorkbook(book, {
    name: 'Base de connaissances.xlsx', modified: known.meta.sourceModified, auditDate: known.meta.auditDate
  }));
  for (const [slug, table] of Object.entries(current.tables)) {
    if (slug === 'contributions') continue;
    assert.deepEqual(JSON.parse(JSON.stringify(table)), known.tables[slug]);
  }
  const stats = Core.stats(current);
  assert.equal(stats.projects, 70);
  assert.equal(stats.labs, 35);
  assert.equal(stats.agencyPubs, 408);
  assert.equal(stats.patents, 59);
});
