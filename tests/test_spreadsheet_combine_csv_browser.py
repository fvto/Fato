"""Use the real CSV controls and mixed-file combine path in Chrome."""
from pathlib import Path
import io
import os
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
from test_spreadsheet_combine_csv import csv_bytes, workbook_bytes


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for headless Chrome')
class CsvCombineBrowserTests(unittest.TestCase):
    def test_mixed_files_validation_encoding_override_and_150_file_controls(self):
        from playwright.sync_api import sync_playwright, expect
        with tempfile.TemporaryDirectory() as temporary, patch.object(combine, 'ROOT', Path(temporary) / 'jobs'):
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
                    uploader = page.locator('#combine-upload')
                    expect(uploader).to_have_attribute('data-max-files', '150')
                    expect(uploader).to_have_attribute('accept', '.xls,.xlsx,.xlsm,.csv')
                    # Both picker and drop share upload(), including early
                    # validation that avoids sending unsupported/oversize files.
                    uploader.set_input_files([dict(name='bad.txt', mimeType='text/plain', buffer=b'data')])
                    expect(page.locator('#combine-file-errors')).to_contain_text('bad.txt')
                    uploader.set_input_files([dict(name=f'{i}.csv', mimeType='text/csv', buffer=b'ID\n001') for i in range(151)])
                    expect(page.locator('#combine-notice')).to_contain_text('Tối đa 150 file')
                    uploader.set_input_files([dict(name='broken.csv', mimeType='text/csv', buffer=b'ID\n"broken'),
                                              dict(name='empty.csv', mimeType='text/csv', buffer=b'')])
                    expect(page.locator('#combine-file-errors li')).to_have_count(2)
                    expect(page.locator('#combine-file-errors')).to_contain_text('broken.csv')
                    expect(page.locator('#combine-preview')).to_be_disabled()
                    # Explicit legacy encoding is persisted through preview and
                    # export, while XLSX in the same batch stays unaffected.
                    page.locator('.combine-csv-options summary').click()
                    page.locator('#combine-csv-encoding').select_option('cp1252')
                    page.locator('#combine-csv-delimiter').select_option(';')
                    uploader.set_input_files([
                        dict(name='Sales_Hanoi.csv', mimeType='text/csv', buffer=csv_bytes([
                            ['ID','Name'], ['001','André; Paris']], delimiter=';', encoding='cp1252')),
                        dict(name='Sales_HCM.xlsx', mimeType='application/octet-stream', buffer=workbook_bytes([
                            ['Name','ID'], ['Hà Nội','002']])),
                    ])
                    expect(page.locator('#combine-file-count')).to_have_text('2 file')
                    expect(page.locator('.combine-file').first).to_contain_text('cp1252')
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-stats')).to_have_text('2 file · 2 dòng · 5 cột')
                    expect(page.locator('#combine-preview-table')).to_contain_text('André; Paris')
                    expect(page.locator('#combine-preview-table tbody tr').first.locator('td').first).to_have_text('001')
                    if os.environ.get('CHINGLUH_COMBINE_SCREENSHOT'):
                        page.screenshot(path=os.environ['CHINGLUH_COMBINE_SCREENSHOT'], full_page=True)
                    page.locator('#combine-run').click()
                    expect(page.locator('#combine-download')).to_be_visible(timeout=60000)
                    result = page.request.get(url + page.locator('#combine-download').get_attribute('href'))
                    book = openpyxl.load_workbook(io.BytesIO(result.body()), read_only=True)
                    self.assertEqual(list(book['Combined'].values)[1][:2], ('001', 'André; Paris'))
                    book.close()
                    # Exactly 150 files exercise DOM creation, selection and
                    # preview. Distinct content avoids the duplicate-file guard.
                    page.locator('#combine-csv-encoding').select_option('auto')
                    page.locator('#combine-csv-delimiter').select_option('auto')
                    # A real >8 MiB request uses the background CSV scan,
                    # preview and export endpoints, including browser polling.
                    large_csv=('ID,Text\n'+''.join(f'{i:05},'+('text'*225)+'\n' for i in range(10000))).encode()
                    with page.expect_response(lambda response: response.url.endswith('/api/combine/uploads') and response.request.method=='POST') as upload_response:
                        uploader.set_input_files([dict(name='Async.csv',mimeType='text/csv',buffer=large_csv)])
                    self.assertEqual(upload_response.value.status,202)
                    expect(page.locator('#combine-file-count')).to_have_text('1 file',timeout=20000)
                    expect(page.locator('.combine-file')).to_contain_text('CSV ·')
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-stats')).to_have_text('1 file · 10000 dòng · 5 cột',timeout=20000)
                    page.locator('#combine-run').click()
                    expect(page.locator('#combine-download')).to_be_visible(timeout=30000)
                    uploader.set_input_files([dict(name=f'{i}.csv', mimeType='text/csv', buffer=f'ID,Count\n{i:03},{i}\n'.encode()) for i in range(150)])
                    expect(page.locator('#combine-file-count')).to_have_text('150 file', timeout=15000)
                    expect(page.locator('.combine-sheet')).to_have_count(150)
                    page.locator('#combine-preview').click()
                    expect(page.locator('#combine-stats')).to_have_text('150 file · 150 dòng · 5 cột')
                    # Drop receives the same client validation as the picker.
                    page.evaluate("""() => {
                        const transfer = new DataTransfer();
                        transfer.items.add(new File(['data'], 'dropped.txt'));
                        document.getElementById('combine-drop').dispatchEvent(new DragEvent('drop', {dataTransfer:transfer, bubbles:true}));
                    }""")
                    expect(page.locator('#combine-file-errors')).to_contain_text('dropped.txt')
                    page.set_viewport_size(dict(width=390, height=844))
                    self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
                    self.assertEqual(errors, [])
                    browser.close()
            finally:
                http.shutdown()
                thread.join(timeout=5)


if __name__ == '__main__':
    unittest.main()
