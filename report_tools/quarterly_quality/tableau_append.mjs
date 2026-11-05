import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';

const [inputPath, outputPath, moduleRoot] = process.argv.slice(2);
const require = createRequire(import.meta.url);
const entry = require.resolve('@oai/artifact-tool', {paths: [moduleRoot]});
const { Workbook, SpreadsheetFile } = await import(pathToFileURL(entry).href);
const rows = JSON.parse(await fs.readFile(inputPath, 'utf8'));
const workbook = Workbook.create();
const sheet = workbook.worksheets.add('VH FTT');
sheet.getRange(`A1:I${rows.length}`).values = rows;
await (await SpreadsheetFile.exportXlsx(workbook)).save(outputPath);
// Release the native spreadsheet runtime before Windows process teardown.
process.exit(0);
