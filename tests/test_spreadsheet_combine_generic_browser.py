"""Real workbooks of several report types through the public combine UI."""
from pathlib import Path
import io
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import openpyxl
import spreadsheet_combine as combine
import sfc_dashboard_server as server
from werkzeug.serving import make_server


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for headless Chrome')
class GenericCombineBrowserTests(unittest.TestCase):
    def test_sales_inventory_and_payroll_with_real_xlsx_files(self):
        from playwright.sync_api import sync_playwright, expect
        root = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as temporary, patch.object(combine, 'ROOT', Path(temporary) / 'jobs'):
            fixture_dir = Path(temporary) / 'fixtures'
            node, modules = combine.runtime()
            subprocess.run([str(node), str(root / 'tests/fixtures/combine_generic_fixtures.mjs'), str(fixture_dir), str(modules)],
                           check=True, capture_output=True, timeout=60)
            http = make_server('127.0.0.1', 0, server.app, threaded=True)
            thread = threading.Thread(target=http.serve_forever, daemon=True)
            thread.start()
            url = f'http://127.0.0.1:{http.server_port}'
            try:
                with sync_playwright() as p:
                    browser = p.chromium.launch(channel='chrome', headless=True)
                    page = browser.new_page(viewport=dict(width=1440, height=1000))
                    errors = []
                    page.on('pageerror', lambda error: errors.append(str(error)))
                    page.goto(url + '/combine')
                    page.locator('#combine-upload').set_input_files([str(fixture_dir / f'Sales_2026-{month}.xlsx') for month in ['09','10']])
                    expect(page.locator('#combine-file-count')).to_have_text('2 file')
                    headers = page.locator('[data-field=header]')
                    expect(headers.nth(0)).to_have_value('3')
                    expect(headers.nth(1)).to_have_value('2')
                    page.locator('[data-field=end]').nth(0).fill('5')
                    page.locator('[data-field=end]').nth(1).fill('4')
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-stats')).to_have_text('2 file · 4 dòng · 7 cột')
                    expect(page.locator('#combine-output-name')).to_have_value('Sales_2026-09_2026-10_Combined.xlsx')
                    expect(page.locator('#combine-preview-table tbody tr').first.locator('td').nth(0)).to_have_text('001')
                    expect(page.locator('#combine-preview-table tbody tr').first.locator('td').nth(2)).to_have_text('25.0%')
                    page.locator('.combine-options summary').click()
                    page.locator('#combine-schema-mode').select_option('strict')
                    expect(page.locator('#combine-result')).to_be_hidden()
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-notice')).to_contain_text('bộ cột khác')
                    expect(page.locator('#combine-result')).to_be_hidden()
                    page.locator('#combine-schema-mode').select_option('union')
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-result')).to_be_visible()
                    page.locator('#combine-run').click()
                    expect(page.locator('#combine-download')).to_be_visible(timeout=60000)
                    response = page.request.get(url + page.locator('#combine-download').get_attribute('href'))
                    self.assertIn('Sales_2026-09_2026-10_Combined.xlsx', response.headers['content-disposition'])
                    book = openpyxl.load_workbook(io.BytesIO(response.body()))
                    self.assertEqual(book['Combined'].max_row, 5)
                    self.assertEqual(book['Combined']['C2'].value, .25)
                    self.assertEqual(book['Combined']['D4'].value, 'North')
                    self.assertEqual(book['Sources']['E2'].value, 3)
                    self.assertEqual(book['Sources']['F2'].value, 5)
                    book.close()
                    page.set_viewport_size(dict(width=390, height=844))
                    self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
                    page.locator('#combine-result').screenshot(path=str(root/'scratch/combine-generic-sales-mobile.png'))
                    page.set_viewport_size(dict(width=1440, height=1000))
                    for files, expected_name, rows, columns in [
                        (['Stock_20261002.xlsx', 'Stock_20261008.xlsx'], 'Stock_2026-10-02_2026-10-08_Combined.xlsx', 3, 6),
                        (['Payroll_Hanoi.xlsx', 'Payroll_HCM.xlsx'], 'Payroll_2files_Combined.xlsx', 2, 5),
                    ]:
                        page.locator('#combine-upload').set_input_files([str(fixture_dir / name) for name in files])
                        expect(page.locator('#combine-notice')).to_contain_text('Đã đọc 2 file')
                        page.locator('#combine-preview').click()
                        expect(page.locator('#combine-stats')).to_have_text(f'2 file · {rows} dòng · {columns} cột')
                        expect(page.locator('#combine-output-name')).to_have_value(expected_name)
                        page.locator('#combine-run').click()
                        expect(page.locator('#combine-download')).to_be_visible(timeout=60000)
                        response = page.request.get(url + page.locator('#combine-download').get_attribute('href'))
                        self.assertEqual(response.status, 200)
                        self.assertIn(expected_name, response.headers['content-disposition'])
                    # A previous output is a valid input; preserve its original
                    # provenance while adding the current combine's columns.
                    previous = fixture_dir / 'PayrollMerged.xlsx'
                    previous.write_bytes(response.body())
                    page.locator('#combine-upload').set_input_files(str(previous))
                    expect(page.locator('#combine-notice')).to_contain_text('Đã đọc 1 file')
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-stats')).to_have_text('1 file · 2 dòng · 8 cột')
                    expect(page.locator('#combine-method')).to_contain_text('Source File (combine 1)')
                    page.locator('#combine-run').click()
                    expect(page.locator('#combine-download')).to_be_visible(timeout=60000)
                    self.assertEqual(errors, [])
                    browser.close()
            finally:
                http.shutdown()
                thread.join(timeout=5)


if __name__ == '__main__':
    unittest.main()
