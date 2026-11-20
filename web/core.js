const text = v => v == null ? '' : String(v);
const normalize = v => text(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
const key = v => normalize(v).replace(/[^a-z0-9]/g, '');
const escape = v => text(v).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const safeUrl = v => { try { const u = new URL(text(v)); return ['https:', 'http:'].includes(u.protocol) ? u.href : null; } catch (_) { return null; } };
const value = (r, k) => r.values[k];
const isGep = (table, r) => {
  const fields = {publications:['Programme ou plateforme'], projets:['Portefeuille'], indicateurs:['Périmètre'], brevets:['Titulaires indiqués']}[table] || [];
  return fields.some(f => /\bGEP\b|Green Energy Park/i.test(text(value(r, f))));
};
const projectNames = (table, r) => {
  const fields = {projets:['Projet'], publications:['Projet lié'], brevets:['Projet lié'], labos:['Projet rapproché ou proposé'], organismes:['Projet']}[table] || [];
  return fields.flatMap(f => text(value(r, f)).split(/[;\n]+/)).map(key).filter(Boolean);
};
const related = (data, table, record, target) => {
  const names = projectNames(table, record);
  if (!names.length) return [];
  return (data.tables[target]?.rows || []).filter(r => !(target === table && r.id === record.id) && projectNames(target, r).some(n => names.includes(n)));
};
const searchRows = (rows, query) => {
  const terms = normalize(query).trim().split(/\s+/).filter(Boolean);
  return rows.filter(r => { const haystack = normalize(Object.values(r.values).join(' ')); return terms.every(t => haystack.includes(t)); });
};
function filterRows(data, table, opts = {}) {
  let rows = [...(data.tables[table]?.rows || [])];
  if (table === 'projets' && opts.scope !== 'all') rows = rows.filter(r => text(value(r, 'Dans la synthèse AAP')).startsWith('Oui'));
  if (table === 'publications') {
    const scope = opts.scope || 'agency';
    if (scope === 'agency') rows = rows.filter(r => value(r, 'Agence de moyens — audit') === 'Retenue Agence');
    if (scope === 'retained') rows = rows.filter(r => text(value(r, 'Prise en compte')).startsWith('Retenue'));
    if (scope === 'confirmed') rows = rows.filter(r => value(r, 'Prise en compte') === 'Retenue');
    if (scope === 'gep') rows = rows.filter(r => value(r, 'Prise en compte') === 'Retenue' && isGep(table, r));
  }
  if (table === 'brevets' && opts.scope !== 'all') rows = rows.filter(r => value(r, 'Prise en compte') === 'Utilisable');
  if (table === 'indicateurs' && opts.scope === 'gep') rows = rows.filter(r => isGep(table, r));
  if (opts.project) rows = rows.filter(r => projectNames(table, r).includes(key(opts.project)));
  if (opts.year) rows = rows.filter(r => text(value(r, 'Année')) === text(opts.year));
  if (opts.field && opts.value) rows = rows.filter(r => text(value(r, opts.field)) === opts.value);
  if (opts.favorites) rows = rows.filter(r => opts.favorites.includes(table + '/' + r.id));
  return searchRows(rows, opts.query || '');
}
function stats(data) {
  return {projects:filterRows(data,'projets').length, labs:data.meta.institutionalLabs,
    inventories:data.tables.labos.rows.length, agencyPubs:filterRows(data,'publications').length,
    gepPubs:filterRows(data,'publications',{scope:'gep'}).length,
    retainedPubs:filterRows(data,'publications',{scope:'retained'}).length,
    patents:filterRows(data,'brevets').length, editions:data.tables.editions.rows.length,
    sources:data.tables.sources.rows.length, pending:data.tables.verification.rows.length};
}
function years(rows) {
  const buckets = new Map();
  rows.forEach(r => { const n = Number(value(r, 'Année')); if (n > 1900 && n < 2200) buckets.set(n, (buckets.get(n) || 0) + 1); });
  return [...buckets].sort((a,b) => a[0]-b[0]);
}
function validate(data) {
  if (!data || data.schemaVersion !== 1 || !data.meta || !data.tables) throw new Error('Format de données non reconnu.');
  for (const table of ['projets','publications','brevets','labos','sources','indicateurs','appels','editions','organismes','verification','contributions']) {
    const t = data.tables[table];
    if (!t || !Array.isArray(t.rows) || !Array.isArray(t.columns)) throw new Error('Rubrique absente : ' + table);
    const ids = new Set();
    for (const row of t.rows) {
      if (!row.id || ids.has(row.id) || !row.values || !row.links) throw new Error('Référence invalide : ' + table);
      if (Object.values(row.values).some(v => v !== null && !['string','number','boolean'].includes(typeof v))) throw new Error('Valeur inattendue : ' + table);
      for (const url of Object.values(row.links)) if (!safeUrl(url) && !text(url).startsWith('#')) throw new Error('Lien non autorisé : ' + table);
      ids.add(row.id);
    }
  }
  if (data.files && (!data.files.tree || typeof data.files.tree !== 'object')) throw new Error('Index documentaire invalide.');
  return data;
}
const csv = (rows, columns) => '\uFEFF' + [columns, ...rows.map(r => columns.map(k => r.values[k]))].map(row => row.map(v => {
  let s = text(v); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return '"' + s.replace(/"/g, '""') + '"';
}).join(';')).join('\r\n');
export {text,normalize,key,escape,safeUrl,value,isGep,projectNames,related,searchRows,filterRows,stats,years,validate,csv};
