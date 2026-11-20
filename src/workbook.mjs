import ExcelJS from 'exceljs';

const sheetNames = new Map([
  ['Projets', 'projets'],
  ['Appels et budgets', 'appels'],
  ['Labos et équipements', 'labos'],
  ['Publications', 'publications'],
  ['Brevets', 'brevets'],
  ['Chiffres et périmètres', 'indicateurs'],
  ['À vérifier', 'verification'],
  ['Contributions', 'contributions'],
  ['Sources', 'sources'],
  ['Éditions R&D', 'editions'],
  ['Organismes et sites', 'organismes']
]);

const hyperlinkFormula = /^HYPERLINK\(\s*"((?:[^"]|"")*)"\s*[,;]\s*"((?:[^"]|"")*)"\s*\)$/i;
const internalCell = /^#(?:'((?:[^']|'')+)'|([^'!]+))!\$?[A-Z]+\$?(\d+)$/i;

function isWebLink(link) {
  try {
    const url = new URL(link);
    return ['http:', 'https:'].includes(url.protocol) && !url.username && !url.password;
  } catch {
    return false;
  }
}

function scalar(value) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (value == null || ['string', 'number', 'boolean'].includes(typeof value)) return value ?? null;
  if (Array.isArray(value.richText)) return value.richText.map(part => part.text).join('');
  return null;
}

function readCell(cell) {
  const value = cell.value;
  if (value && typeof value === 'object') {
    const formula = cell.formula?.replace(/^=/, '');
    const match = formula?.match(hyperlinkFormula);
    if (match) {
      const target = match[1].replaceAll('""', '"');
      const label = match[2].replaceAll('""', '"');
      return { value: scalar(value.result) ?? label, target };
    }
    if (typeof value.hyperlink === 'string') return { value: scalar(value.text), target: value.hyperlink };
    if ('result' in value) return { value: scalar(value.result) };
  }
  return { value: scalar(value) };
}

function columnNumber(letters) {
  return [...letters].reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
}

function tableRange(table) {
  const range = table.table.tableRef?.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!range) throw new Error('Un tableau du classeur présente une structure inattendue.');
  const [, start, row, end, last] = range;
  const firstCol = columnNumber(start);
  const colCount = columnNumber(end) - firstCol + 1;
  const headerRow = Number(row);
  const lastRow = Number(last) - Number(table.totalsRow);
  if (colCount < 1 || colCount > 200 || lastRow < headerRow || lastRow - headerRow > 30_000) {
    throw new Error('Un tableau du classeur dépasse la taille prévue.');
  }
  return { firstCol, colCount, headerRow, lastRow };
}

function readTable(sheet, slug) {
  const columns = [];
  const rows = [];
  const seen = new Set();
  const blocks = sheet.getTables().map(tableRange).sort((a, b) => a.headerRow - b.headerRow);
  if (!blocks.length && slug !== 'contributions') throw new Error('Tableau structuré manquant : ' + sheet.name);
  for (const block of blocks) {
    const headers = Array.from({length: block.colCount}, (_, i) =>
      String(readCell(sheet.getCell(block.headerRow, block.firstCol + i)).value ?? '').trim()
    );
    if (!columns.length) columns.push(...headers);
    if (headers.some((header, i) => !header || header !== columns[i]) ||
        headers.length !== columns.length || new Set(headers).size !== headers.length) {
      throw new Error('Les en-têtes du tableau ont changé : ' + sheet.name);
    }
    // Les propositions de correction se consultent dans Excel, avec les droits SharePoint.
    if (slug === 'contributions') continue;
    for (let excelRow = block.headerRow + 1; excelRow <= block.lastRow; excelRow++) {
      const values = Object.create(null);
      const links = Object.create(null);
      headers.forEach((header, i) => {
        const { value, target } = readCell(sheet.getCell(excelRow, block.firstCol + i));
        values[header] = value;
        if (target && (isWebLink(target) || internalCell.test(target))) links[header] = target;
      });
      if (!Object.values(values).some(value => value != null && value !== '')) continue;
      const id = String(values['Référence de suivi'] || values['ID relation'] || values['Références de suivi'] || slug + '-' + excelRow);
      if (seen.has(id)) throw new Error('Une référence apparaît plusieurs fois dans le tableau : ' + sheet.name);
      seen.add(id);
      rows.push({ id, excelRow, values, links });
    }
  }
  return {
    sheet: sheet.name,
    title: String(readCell(sheet.getCell('A1')).value || sheet.name),
    description: String(readCell(sheet.getCell('A2')).value || ''),
    columns,
    rows
  };
}

function resolveInternalLinks(tables) {
  const byRow = new Map(Object.entries(tables).map(([slug, table]) => [
    slug, new Map(table.rows.map(row => [row.excelRow, row.id]))
  ]));
  for (const table of Object.values(tables)) {
    for (const record of table.rows) {
      for (const [column, target] of Object.entries(record.links)) {
        const match = target.match(internalCell);
        if (!match) continue;
        const slug = sheetNames.get((match[1] || match[2]).replaceAll("''", "'"));
        const id = byRow.get(slug)?.get(Number(match[3]));
        if (id) record.links[column] = '#/' + slug + '/' + encodeURIComponent(id);
        else delete record.links[column];
      }
    }
  }
}

export async function readWorkbook(bytes, meta) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(bytes);
  const tables = {};
  for (const [name, slug] of sheetNames) {
    const sheet = book.getWorksheet(name);
    if (!sheet) throw new Error('Onglet manquant : ' + name);
    tables[slug] = readTable(sheet, slug);
  }
  resolveInternalLinks(tables);
  return {
    schemaVersion: 1,
    meta: {
      auditDate: meta.auditDate,
      sourceModified: meta.modified,
      readAt: new Date().toISOString(),
      sourceName: meta.name,
      institutionalLabs: 35,
      institutionalLabsNote: 'Repère institutionnel de 35 labos soutenus. Le tableau des équipements recense des projets et ne permet pas de compter les laboratoires physiques distincts.'
    },
    tables
  };
}
