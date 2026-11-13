"""Exercise the PDF library with a separate headless browser and isolated SFC fixtures."""
from io import BytesIO
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from test_workspace_appearance import isolated_server, capture
import sfc_dashboard_server as server
import sfc_qms as qms


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for headless Chrome')
class QmsBrowserTests(unittest.TestCase):
    def test_search_language_pagination_download_error_and_mobile(self):
        from playwright.sync_api import sync_playwright, expect
        documents = [dict(id=f'{i:032x}', source='document', plant='VH', code=f'S{i:04}',
                          name=f'Shoe {i}', kind='QA manual', language='VN' if i == 0 else 'EN',
                          date='2026-11-03', project='HFPA', shoe='', process='',
                          url=f'/api/qms/documents/VH/{i:032x}/pdf') for i in range(25)]
        documents.append(dict(documents[0], id='f' * 32, source='manual', name='<img src=x onerror=alert(1)> Guide',
                              kind='HFPA Guide', code='', language=''))
        catalog = dict(plant='VH', refreshed_at='2026-11-10T10:00:00+07:00', documents=documents,
                       sources=dict(document=dict(rows=25), manual=dict(rows=0)))
        with tempfile.TemporaryDirectory() as zip_folder, isolated_server() as url, \
             patch.object(qms, 'document_catalog', side_effect=lambda plant: catalog if plant == 'VH' else dict(plant=plant, documents=[], refreshed_at=None)), \
             patch.object(qms, 'download_pdf', side_effect=lambda *args: (BytesIO(b'%PDF-1.7\nTest\n%%EOF'), 'S0000_VN.pdf')), \
             patch.object(server.sfc_qms_bulk, 'archive_file', return_value=Path(zip_folder) / 'QA_PDFs.zip'), \
             sync_playwright() as playwright:
            with ZipFile(Path(zip_folder) / 'QA_PDFs.zip', 'w') as archive:
                archive.writestr('test.pdf', b'%PDF-1.7\nTest\n%%EOF')
            browser = playwright.chromium.launch(channel='chrome', headless=True)
            try:
                for width in (1440, 390):
                    page = browser.new_page(viewport=dict(width=width, height=1000))
                    page.route('**/api/status*', lambda route: route.fulfill(json=dict(connected=False)))
                    page.route('**/api/page/inspect', lambda route: route.fulfill(json=dict(connected=False)))
                    errors = []
                    page.on('pageerror', lambda error: errors.append(str(error)))
                    page.goto(url)
                    expect(page.locator('.qms-pdf-shortcut')).to_have_count(2)
                    page.locator('.qms-menu-group li').filter(has=page.get_by_role('link', name='QA Document', exact=True)).get_by_role('button').click()
                    expect(page.locator('#qa-pdf-source')).to_have_value('document')
                    expect(page.locator('#qa-pdf-source option')).to_have_count(2)
                    expect(page.locator('#qa-pdf-download-all')).to_contain_text('QA Document (25 PDF)')
                    expect(page.locator('.qa-pdf-row')).to_have_count(20)
                    page.locator('#qa-pdf-next').click()
                    expect(page.locator('.qa-pdf-row')).to_have_count(5)
                    page.locator('#qa-pdf-search').fill('S0000')
                    expect(page.locator('.qa-pdf-row')).to_have_count(1)
                    expect(page.locator('.qa-pdf-row')).to_contain_text('VN')
                    page.locator('#qa-pdf-language').select_option('EN')
                    expect(page.locator('.qa-pdf-row')).to_have_count(0)
                    page.locator('#qa-pdf-language').select_option('VN')
                    with page.expect_download() as saved:
                        page.locator('.qa-pdf-download').click()
                    self.assertEqual(saved.value.suggested_filename, 'S0000_VN.pdf')
                    page.route('**/api/qms/documents/VH/*/pdf', lambda route: route.fulfill(status=409, json=dict(error='Đăng nhập lại SFC.')))
                    page.locator('.qa-pdf-download').click()
                    expect(page.locator('#qa-pdf-status')).to_contain_text('Đăng nhập lại SFC.')
                    page.locator('#qa-pdf-source').select_option('manual')
                    page.locator('#qa-pdf-language').select_option('')
                    page.locator('#qa-pdf-search').fill('Guide')
                    expect(page.locator('.qa-pdf-row')).to_have_count(0)
                    page.locator('#qa-pdf-collection').select_option('hfpa')
                    expect(page.locator('.qa-pdf-row')).to_have_count(1)
                    self.assertEqual(page.locator('#qa-pdf-list img').count(), 0)
                    page.locator('#qa-pdf-source').select_option('document')
                    page.locator('#qa-pdf-search').fill('')
                    page.route('**/api/qms/refresh', lambda route: route.fulfill(json=dict(modules=57, pdfs=26)))
                    with page.expect_request('**/api/qms/refresh') as refresh:
                        page.locator('#qa-pdf-refresh').click()
                    self.assertEqual(refresh.value.post_data_json, {'plant': 'VH'})
                    expect(page.locator('#qa-pdf-refresh')).to_be_enabled()
                    expect(page.locator('#qa-pdf-status')).to_contain_text('25 / 25 liên kết PDF')
                    expect(page.locator('#qa-pdf-reconciliation')).to_contain_text('25 dòng SFC')
                    job_id = 'b' * 32
                    job = dict(id=job_id, plant='VH', source='document', status='running', completed=0, total=25, current='S0000')
                    page.route('**/api/qms/bulk', lambda route: route.fulfill(status=202, json=job))
                    page.route('**/api/qms/bulk/' + job_id, lambda route: route.fulfill(json=dict(job, status='done', completed=25,
                               url='/api/qms/bulk/' + job_id + '/zip')))
                    with page.expect_request('**/api/qms/bulk') as bulk_request:
                        page.locator('#qa-pdf-download-all').click()
                    self.assertEqual(bulk_request.value.post_data_json['ids'], [item['id'] for item in documents if item['source'] == 'document'])
                    self.assertEqual(bulk_request.value.post_data_json['source'], 'document')
                    expect(page.locator('#qa-pdf-bulk-status')).to_contain_text('Đã lấy đủ 25/25')
                    with page.expect_download() as all_saved:
                        page.locator('#qa-pdf-bulk-link').click()
                    self.assertEqual(all_saved.value.suggested_filename, 'QA_PDFs.zip')
                    retry_state = {'done': False}
                    partial = dict(job, status='partial', completed=24, missing_count=1, resumable=True,
                                   url='/api/qms/bulk/' + job_id + '/zip',
                                   failed=[dict(code='S0001', message='SFC trả HTTP 404: file PDF gốc không tồn tại.')])
                    page.route('**/api/qms/bulk/' + job_id, lambda route: route.fulfill(json=(
                        dict(job, status='done', completed=25, url='/api/qms/bulk/' + job_id + '/zip') if retry_state['done'] else partial)))
                    page.locator('#qa-pdf-source').select_option('manual')
                    page.locator('#qa-pdf-source').select_option('document')
                    expect(page.locator('#qa-pdf-bulk-status')).to_contain_text('SFC trả HTTP 404')
                    expect(page.locator('#qa-pdf-bulk-link')).to_contain_text('CHƯA ĐỦ (24/25 PDF)')
                    expect(page.locator('#qa-pdf-bulk-resume')).to_be_visible()
                    def retry(route):
                        retry_state['done'] = True
                        route.fulfill(status=202, json=dict(job, status='queued', completed=24))
                    page.route('**/api/qms/bulk/' + job_id + '/resume', retry)
                    with page.expect_request('**/api/qms/bulk/' + job_id + '/resume') as resumed:
                        page.locator('#qa-pdf-bulk-resume').click()
                    self.assertEqual(resumed.value.method, 'POST')
                    expect(page.locator('#qa-pdf-bulk-link')).to_contain_text('QA Document (25 PDF)')
                    expect(page.locator('#qa-pdf-bulk-resume')).to_be_hidden()
                    page.locator('#qa-pdf-source').select_option('manual')
                    expect(page.locator('#qa-pdf-bulk-link')).to_be_hidden()
                    page.locator('#qa-pdf-collection').select_option('hfpa')
                    expect(page.locator('#qa-pdf-download-all')).to_contain_text('QA Manual (1 PDF)')
                    manual_job = dict(job, id='c' * 32, source='manual', total=1, completed=1, status='done',
                                      url='/api/qms/bulk/' + 'c' * 32 + '/zip')
                    page.route('**/api/qms/bulk', lambda route: route.fulfill(status=202, json=manual_job))
                    page.route('**/api/qms/bulk/' + 'c' * 32, lambda route: route.fulfill(json=manual_job))
                    with page.expect_request('**/api/qms/bulk') as manual_request:
                        page.locator('#qa-pdf-download-all').click()
                    self.assertEqual(manual_request.value.post_data_json['source'], 'manual')
                    self.assertEqual(manual_request.value.post_data_json['ids'], ['f' * 32])
                    expect(page.locator('#qa-pdf-bulk-link')).to_contain_text('QA Manual (1 PDF)')
                    page.locator('#qa-pdf-source').select_option('document')
                    expect(page.locator('#qa-pdf-bulk-link')).to_contain_text('QA Document (25 PDF)')
                    page.locator('#qa-pdf-library').scroll_into_view_if_needed()
                    capture(page, f'qms-pdf-{width}.png')
                    self.assertLessEqual(page.evaluate('document.documentElement.scrollWidth'), width)
                    panel = page.locator('#qa-pdf-library').bounding_box()
                    self.assertGreaterEqual(panel['x'], 0)
                    self.assertLessEqual(panel['x'] + panel['width'], width)
                    page.locator('#qa-pdf-plant').select_option('VH2')
                    expect(page.locator('#qa-pdf-list')).to_contain_text('Chưa có danh mục')
                    expect(page.locator('.qa-pdf-row')).to_have_count(0)
                    self.assertEqual(errors, [])
                    page.close()
            finally:
                browser.close()


if __name__ == '__main__':
    unittest.main()
