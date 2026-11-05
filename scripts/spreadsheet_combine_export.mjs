// Excel output uses the bundled Artifact Tool; inputs are always literal data.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';

const [directory, modules] = process.argv.slice(2);
const require = createRequire(path.join(modules, '..', 'combine-runtime.cjs'));
const {Workbook, SpreadsheetFile} = await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const payload = JSON.parse(await fs.readFile(path.join(directory, 'payload.json'), 'utf8'));
const workbook = Workbook.create();

function column(index) {
  let result = '';
  for (index++; index > 0; index = Math.floor((index - 1) / 26)) result = String.fromCharCode(65 + (index - 1) % 26) + result;
  return result;
}

function table(name, headers, rows, formats) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const end = column(headers.length - 1);
  // Artifact Tool accepts Excel's apostrophe escape for literal text. Escape
  // both formula prefixes and existing apostrophes to preserve the source.
  const literal = value => typeof value === 'string' && /^[=']/.test(value) ? "'" + value : value;
  sheet.getRange(`A1:${end}${rows.length + 1}`).values = [headers, ...rows].map(row => row.map(literal));
  const range = sheet.getRange(`A1:${end}${rows.length + 1}`);
  range.format.font = {name: 'Arial', size: 10, color: '#202020'};
  range.format.verticalAlignment = 'center';
  range.format.rowHeight = 20;
  sheet.getRange(`A1:${end}1`).format = {
    fill: '#263D4A', font: {name: 'Arial', size: 10, bold: true, color: '#FFFFFF'},
    wrapText: true, rowHeight: 46, horizontalAlignment: 'center', verticalAlignment: 'center',
  };
  sheet.freezePanes.freezeRows(1);
  sheet.tables.add(`A1:${end}${rows.length + 1}`, true, `${name}Table`);
  for (let c = 0; c < headers.length; c++) {
    const sampleLength = Math.max(0, ...rows.slice(0, 100).map(row => typeof row[c] === 'string' ? row[c].length : 0));
    let width = Math.max(14, Math.min(52, Math.max(headers[c].length * .8, sampleLength * .9)));
    if (/source file|^file$/i.test(headers[c])) width = 46;
    if (/sha256/i.test(headers[c])) width = 72;
    sheet.getRange(`${column(c)}1:${column(c)}${rows.length + 1}`).format.columnWidth = width;
  }
  if (formats && rows.length) {
    sheet.getRange(`A2:${end}${rows.length + 1}`).setNumberFormat(formats);
  }
  return sheet;
}

const combined = table('Combined', payload.columns,
  payload.rows.map(row => row.map(cell => cell.value)),
  payload.rows.map(row => row.map(cell => cell.format)));
combined.tabColor = '#263D4A';
const headers = ['File', 'Sheet', 'Data Rows', 'Source Columns', 'Header Row', 'End Row', 'Blank Rows Skipped', 'Repeated Headers Skipped', 'Verified Totals Skipped', 'Columns Absent in Source', 'Recognized Layout', 'SHA256'];
table('Sources', headers, payload.sources.map(s => [s.file, s.sheet, s.rows, s.columns, s.header_row, s.end_row, s.blank_rows, s.repeated_headers, s.footer_rows, (s.missing_columns || []).join(', '), s.profile || '', s.sha256]));
workbook.recalculate();
console.log((await workbook.inspect({kind: 'table', range: 'Combined!A1:H4', include: 'values,formulas', tableMaxRows: 4, tableMaxCols: 8, maxChars: 1800})).ndjson);
// Save before optional visual QA, which uses platform-native graphics.
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(path.join(directory, payload.output_name));
if (process.env.CHINGLUH_COMBINE_RENDER === '1') {
  for (const [sheetName, range] of [['Combined', 'A1:I10'], ['Sources', 'A1:G11']]) {
    const png = await workbook.render({sheetName, range, scale: 1.5, format: 'png'});
    await fs.writeFile(path.join(directory, `${sheetName}.png`), new Uint8Array(await png.arrayBuffer()));
  }
}
