// Small report-independent workbooks for exercising the full public flow.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
const [folder, modules] = process.argv.slice(2);
const require = createRequire(path.join(modules, '..', 'fixtures.cjs'));
const {Workbook, SpreadsheetFile} = await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const fixtures = [
  ['Sales_2026-09.xlsx', 'Sales', [['Sales report'], [], ['', 'Item\nCode', 'Amount', 'Completion %'], ['', '001', 100, '25.0%'], ['', '002', 250, '50.0%'], ['', 'Total', 350]]],
  ['Sales_2026-10.xlsx', 'Orders', [['Sales report'], ['', 'Amount', 'Item Code', 'Region'], ['', 300, '003', 'North'], ['', 400, '004', 'South'], ['', 700, 'Total']]],
  ['Stock_20261002.xlsx', 'Inventory', [['SKU', 'Stock'], ['001', 10], ['002', 0]]],
  ['Stock_20261008.xlsx', 'Inventory', [['Warehouse', 'Stock', 'SKU'], ['A', 20, '003']]],
  ['Payroll_Hanoi.xlsx', 'Employees', [['Employee', 'Salary'], ['001', 1000]]],
  ['Payroll_HCM.xlsx', 'Employees', [['Salary', 'Employee'], [2000, '002']]],
];
await fs.mkdir(folder, {recursive:true});
for (const [filename, name, rows] of fixtures) {
  const book = Workbook.create();
  const sheet = book.worksheets.add(name);
  const width = Math.max(...rows.map(row => row.length));
  sheet.getRangeByIndexes(0, 0, rows.length, width).values = rows.map(row => [...row, ...Array(width - row.length).fill(null)]);
  book.recalculate();
  await (await SpreadsheetFile.exportXlsx(book)).save(path.join(folder, filename));
}
