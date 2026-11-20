import { build } from 'esbuild';
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const output = join(root, 'dist');
const deploy = process.argv.includes('--deployment');
const value = name => (process.env[name] ?? '').trim();
const guid = /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i;
const graphId = /^[a-zA-Z0-9!._-]{1,256}$/;
const settings = {
  tenantId: value('IRESEN_TENANT_ID'),
  clientId: value('IRESEN_CLIENT_ID'),
  driveId: value('IRESEN_DRIVE_ID'),
  workbookId: value('IRESEN_WORKBOOK_ID'),
  folderId: value('IRESEN_FOLDER_ID'),
  auditDate: value('IRESEN_AUDIT_DATE') || '2026-11-17'
};
const { tenantId, clientId, driveId, workbookId, folderId, auditDate } = settings;
const identifiers = [tenantId, clientId, driveId, workbookId, folderId];
if (deploy || identifiers.some(Boolean)) {
  if (!guid.test(tenantId) || !guid.test(clientId) || ![driveId, workbookId, folderId].every(id => graphId.test(id))) {
    throw new Error('Renseigner les deux identifiants Entra et les trois identifiants SharePoint dans la configuration du déploiement.');
  }
}
if (!/^\d{4}-\d{2}-\d{2}$/.test(auditDate) || Number.isNaN(Date.parse(auditDate))) {
  throw new Error('La date d’audit doit suivre le format AAAA-MM-JJ.');
}
await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
for (const name of ['index.html', 'auth-retour.html', 'styles.css', 'assets']) {
  await cp(join(root, 'web', name), join(output, name), { recursive: true });
}
await build({
  entryPoints: {
    portail: join(root, 'src/bootstrap.mjs'),
    'auth-retour': join(root, 'src/auth-retour.mjs')
  },
  outdir: output,
  bundle: true,
  platform: 'browser',
  format: 'esm',
  target: ['es2022'],
  minify: true,
  legalComments: 'linked',
  define: { __PORTAL_SETTINGS__: JSON.stringify(settings) }
});
await writeFile(join(output, '.nojekyll'), '');
const html = await readFile(join(output, 'index.html'), 'utf8');
if (/src=["'](?:\/|https?:)|data\.js/.test(html)) throw new Error('Les ressources de la page doivent être locales et relatives.');
console.log(deploy ? 'Site assemblé pour GitHub Pages.' : 'Site assemblé dans dist/.');
