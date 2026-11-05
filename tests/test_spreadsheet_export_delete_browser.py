"""Exercise saved-export deletion in Chrome using disposable files only."""
import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet_compare as compare
import sfc_dashboard_server as server
from werkzeug.serving import make_server


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for Chrome')
class ExportDeleteBrowserTests(unittest.TestCase):
    def test_offline_comparison_works_without_internet_or_ai_requests(self):
        from playwright.sync_api import sync_playwright, expect
        import openpyxl

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            paths = []
            for name, rows in [
                ('A.xlsx', [['Factory', 'Line', 'Score'], ['VH', 'L1', 10], ['VH', 'L2', 20], ['JV', 'L1', 30]]),
                ('B.xlsx', [['Factory', 'Line', 'Score'], ['VH', 'L2', 21], ['VH', 'L1', 10], ['JVB', 'L1', 8]]),
            ]:
                book = openpyxl.Workbook()
                book.active.title = 'Data'
                for row in rows:
                    book.active.append(row)
                path = root/name
                book.save(path)
                book.close()
                paths.append(str(path))
            with patch.object(compare, 'ROOT', root/'comparisons'), patch.object(compare, 'EXPORTS', root/'exports'):
                http = make_server('127.0.0.1', 0, server.app, threaded=True)
                thread = threading.Thread(target=http.serve_forever, daemon=True)
                thread.start()
                try:
                    with sync_playwright() as playwright:
                        browser = playwright.chromium.launch(channel='chrome', headless=True)
                        context = browser.new_context(viewport=dict(width=1280, height=1000))
                        # Old chat settings/history must not trigger any AI requests.
                        context.add_init_script("localStorage.setItem('compare-chat-mode','codex');localStorage.setItem('compare-chat-id','aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa');")
                        blocked = []
                        def local_only(route):
                            url = route.request.url
                            if not url.startswith(f'http://127.0.0.1:{http.server_port}/') or '/api/compare/chat/' in url:
                                blocked.append(url)
                                route.abort()
                            else:
                                route.continue_()
                        context.route('**/*', local_only)
                        page = context.new_page()
                        errors = []
                        page.on('pageerror', lambda error: errors.append(str(error)))
                        page.goto(f'http://127.0.0.1:{http.server_port}/compare')
                        expect(page.locator('html')).to_have_attribute('lang', 'vi')
                        expect(page.locator('#chat-mode')).to_have_value('offline')
                        expect(page.locator('#chat-form')).to_be_hidden()
                        expect(page.locator('#offline-guide')).to_be_visible()
                        page.locator('#compare-upload').set_input_files(paths)
                        expect(page.locator('#workbook-count')).to_have_text('2 file')
                        page.locator('.manual-setup-guide > summary').click()
                        page.locator('#offline-add-comparison').click()
                        for field, value in [('left_range', 'A2:C4'), ('right_range', 'A2:C4'),
                                             ('left_keys', 'A,B'), ('right_keys', 'A,B'),
                                             ('left_columns', 'C'), ('right_columns', 'C')]:
                            page.locator(f'[data-field="{field}"]').fill(value)
                        page.locator('#run-compare').click()
                        expect(page.locator('#compare-notice')).to_contain_text('Đã kiểm tra 4 giá trị. 3 mục cần xem lại.')
                        expect(page.locator('#compare-result-rows .result-row')).to_have_count(3)
                        difference = page.locator('#compare-result-rows .result-row').filter(has=page.locator('.status-tag', has_text='Chênh lệch'))
                        expect(difference.locator('td').nth(2)).to_have_text('20')
                        expect(difference.locator('td').nth(3)).to_have_text('21')
                        expect(difference.locator('td').last).to_contain_text('Các giá trị chênh lệch')
                        expect(page.locator('.results-table thead')).to_contain_text('Lý do và bằng chứng')
                        page.locator('#compare-status').select_option('all')
                        expect(page.locator('#compare-result-rows .result-row')).to_have_count(4)
                        expect(page.locator('#compare-result-rows').get_by_text('Khớp', exact=True)).to_have_count(1)
                        page.locator('#workflow-rounding').select_option('1')
                        page.locator('#workflow-preview').click()
                        expect(page.locator('#workflow-plan')).to_be_visible()
                        expect(page.locator('#workflow-plan-content')).to_contain_text('factory')
                        expect(page.locator('#workflow-run')).to_be_disabled()
                        page.locator('#workflow-ack').check()
                        page.locator('#workflow-run').click()
                        expect(page.locator('#compare-notice')).to_contain_text('Hoàn tất quy trình')
                        expect(page.locator('#compare-result-rows .result-row')).to_have_count(3)
                        with page.expect_download() as download:
                            page.locator('#compare-csv').click()
                        csv = Path(download.value.path()).read_text(encoding='utf-8-sig')
                        self.assertIn('difference', csv)
                        self.assertIn('unmatched', csv)
                        with page.expect_download() as html_download:
                            page.locator('#compare-html').click()
                        report_path = root/'report.html'
                        html_download.value.save_as(report_path)
                        document = report_path.read_text(encoding='utf-8')
                        self.assertEqual(document.count('<tr id="finding-'), 3)
                        self.assertEqual(document.count('src="data:image/png;base64,'), 2)
                        # The recipient opens the exported file with all networking disabled.
                        recipient = browser.new_context(offline=True)
                        report_page = recipient.new_page()
                        report_page.goto(report_path.as_uri())
                        expect(report_page.locator('[id^="finding-"]')).to_have_count(3)
                        self.assertEqual(report_page.locator('details[open]').count(),0)
                        report_page.locator('.finding-group > summary').first.click()
                        first_image_link=report_page.locator('.finding-group[open] a[href^="#image-"]').first
                        first_image_link.click()
                        self.assertTrue(report_page.locator('#evidence').evaluate('(node)=>node.open'))
                        expect(report_page.locator('figure img')).to_have_count(2)
                        self.assertTrue(report_page.locator('figure img').evaluate_all('(images)=>images.every(img=>img.complete&&img.naturalWidth>0)'))
                        report_page.screenshot(path=str(root/'report-preview.png'), full_page=True)
                        recipient.close()
                        page.locator('#compare-export-format').select_option('pptx')
                        with page.expect_download(timeout=120000) as pptx_download:
                            page.locator('#compare-html').click()
                        import zipfile
                        with zipfile.ZipFile(pptx_download.value.path()) as archive:
                            self.assertTrue(any(name.startswith('ppt/slides/slide') and name.endswith('.xml') for name in archive.namelist()))
                            self.assertTrue(any(name.startswith('ppt/media/') and name.endswith('.png') for name in archive.namelist()))
                        self.assertEqual(blocked, [])
                        self.assertEqual(errors, [])
                        context.close()
                        browser.close()
                finally:
                    http.shutdown()
                    http.server_close()
                    thread.join(timeout=5)
                    compare.load_workbook.cache_clear()

    def test_selection_cancel_failure_delete_and_refresh_on_desktop_and_mobile(self):
        from playwright.sync_api import sync_playwright, expect

        with tempfile.TemporaryDirectory() as temporary, patch.object(compare, 'EXPORTS', Path(temporary)):
            http = make_server('127.0.0.1', 0, server.app, threaded=True)
            thread = threading.Thread(target=http.serve_forever, daemon=True)
            thread.start()
            try:
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch(channel='chrome', headless=True)
                    for width in (1280, 390):
                        folder = Path(temporary)/str(width)
                        folder.mkdir()
                        saved = folder/'Delete me.xlsx'
                        saved.write_bytes(b'disposable browser fixture')
                        second = folder/'Second.xlsm'
                        second.write_bytes(b'disposable browser fixture')
                        keep = folder/'Keep me.txt'
                        keep.write_bytes(b'keep')
                        selection = next(item for item in compare.catalog() if item['relative'] == f'{width}/Delete me.xlsx')
                        context = browser.new_context(viewport=dict(width=width, height=1000))
                        page = context.new_page()
                        errors = []
                        page.on('pageerror', lambda error: errors.append(str(error)))
                        page.goto(f'http://127.0.0.1:{http.server_port}/compare')
                        catalog = page.locator('#compare-catalog')
                        delete = page.get_by_role('button', name='Xóa tất cả file', exact=True)
                        expect(catalog.locator(f'option[value="{selection["id"]}"]')).to_have_count(1)
                        expect(catalog).to_have_value('')
                        expect(delete).to_be_enabled()
                        page.once('dialog', lambda dialog: dialog.dismiss())
                        delete.click()
                        expect(delete).to_be_enabled()
                        self.assertTrue(saved.exists())
                        expect(catalog).to_have_value('')

                        url = '**/api/compare/catalog/delete-all'
                        page.route(url, lambda route: route.fulfill(status=400, json={'error': 'File đang mở trong Excel / WPS.'}))
                        page.once('dialog', lambda dialog: dialog.accept())
                        delete.click()
                        expect(page.locator('#compare-notice')).to_contain_text('Excel / WPS')
                        expect(delete).to_be_enabled()
                        self.assertTrue(saved.exists())
                        page.unroute(url)

                        def accept(dialog):
                            self.assertIn('TẤT CẢ 2 file', dialog.message)
                            dialog.accept()
                        page.once('dialog', accept)
                        delete.click()
                        expect(page.locator('#compare-notice')).to_contain_text('Đã xóa 2 file đã xuất.')
                        expect(delete).to_be_disabled()
                        expect(catalog).to_have_value('')
                        expect(catalog.locator(f'option[value="{selection["id"]}"]')).to_have_count(0)
                        self.assertFalse(saved.exists())
                        self.assertFalse(second.exists())
                        self.assertTrue(keep.exists())
                        page.reload()
                        expect(catalog.locator(f'option[value="{selection["id"]}"]')).to_have_count(0)
                        expect(delete).to_be_disabled()
                        self.assertEqual(page.evaluate('document.documentElement.scrollWidth > innerWidth'), False)
                        self.assertEqual(errors, [])
                        context.close()
                    browser.close()
            finally:
                http.shutdown()
                http.server_close()
                thread.join(timeout=5)


if __name__ == '__main__':
    unittest.main()
