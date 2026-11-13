"""Beginner offline flow, manual fallback and mobile layout in Chrome."""
import os
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

import openpyxl
from werkzeug.serving import make_server

import spreadsheet_compare as compare
import sfc_dashboard_server as server


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for Chrome')
class OfflineUiTests(unittest.TestCase):
    def test_beginner_flow_and_manual_fallback_on_desktop_and_mobile(self):
        from playwright.sync_api import sync_playwright, expect
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            paths = []
            for name, rows in [
                ('Original.xlsx', [['ID', 'Item', 'Value'], ['A', 'Alpha', 10], ['B', 'Beta', 20]]),
                ('Updated.xlsx', [['Value', 'ID', 'Item'], [25, 'B', 'Beta'], [10, 'A', 'Alpha']]),
                ('Ambiguous.xlsx', [['Group', 'Value'], ['VH', 10], ['VH', 20]]),
                ('Percent A.xlsx', [['ID', 'Rate', 'Count'], ['A', '95%', 10], ['B', '94.40001%', 20]]),
                ('Percent B.xlsx', [['ID', 'Rate', 'Count'], ['Renamed', '95%', 999], ['B', '94.40002%', 40], ['C', '90%', 50]]),
            ]:
                book = openpyxl.Workbook()
                book.active.title = 'Data'
                for row in rows:
                    book.active.append(row)
                path = root/name
                book.save(path)
                book.close()
                paths.append(str(path))
            with patch.object(compare, 'ROOT', root/'compare'), patch.object(compare, 'EXPORTS', root/'exports'):
                http = make_server('127.0.0.1', 0, server.app, threaded=True)
                thread = threading.Thread(target=http.serve_forever, daemon=True)
                thread.start()
                try:
                    with sync_playwright() as playwright:
                        browser = playwright.chromium.launch(channel='chrome', headless=True)
                        for width in (1280, 390):
                            context = browser.new_context(viewport={'width': width, 'height': 1000})
                            blocked, errors, runs, quick_runs = [], [], [], []
                            def local_only(route):
                                url = route.request.url
                                if not url.startswith(f'http://127.0.0.1:{http.server_port}/') or '/api/compare/chat/' in url:
                                    blocked.append(url)
                                    route.abort()
                                else:
                                    if '/workflows/' in url and url.endswith('/run'):
                                        runs.append(url)
                                    if url.endswith('/api/compare/quick'):
                                        quick_runs.append(url)
                                    route.continue_()
                            context.route('**/*', local_only)
                            page = context.new_page()
                            page.on('pageerror', lambda error: errors.append(str(error)))
                            page.goto(f'http://127.0.0.1:{http.server_port}/compare')
                            expect(page.locator('#advanced-checks')).not_to_have_attribute('open', '')
                            expect(page.locator('#workflow-options')).not_to_have_attribute('open', '')
                            expect(page.locator('#compare-catalog')).to_be_hidden()
                            expect(page.locator('#workflow-preview')).to_be_disabled()
                            expect(page.locator('#workflow-files')).to_be_hidden()
                            page.locator('#compare-upload').set_input_files(paths[0])
                            expect(page.locator('#offline-next-step')).to_contain_text('Thêm 1 file')
                            page.locator('#compare-upload').set_input_files(paths[1])
                            expect(page.locator('#workflow-files')).to_be_visible()
                            expect(page.locator('#workflow-preview')).to_be_enabled()
                            page.locator('#workflow-preview').click()
                            expect(page.locator('#workflow-plan')).to_be_visible()
                            expect(page.locator('#workflow-plan-content')).to_contain_text('ghép dòng theo')
                            expect(page.locator('#workflow-run')).to_be_enabled()
                            expect(page.locator('#workflow-preview')).to_be_hidden()
                            screenshots = os.environ.get('CHINGLUH_UI_PREVIEW_DIR')
                            if screenshots:
                                target = Path(screenshots)
                                target.mkdir(parents=True, exist_ok=True)
                                page.screenshot(path=str(target/f'compare-{width}.png'), full_page=True)
                            self.assertEqual(runs, [])
                            page.locator('#workflow-run').click()
                            expect(page.locator('#compare-result-rows .result-row')).to_have_count(1)
                            expect(page.locator('#offline-step-result')).to_have_attribute('aria-current', 'step')
                            expect(page.locator('#compare-result-rows')).to_contain_text('25')
                            self.assertEqual(len(runs), 1)
                            self.assertFalse(page.evaluate('document.documentElement.scrollWidth > innerWidth'))
                            page.reload()
                            expect(page.locator('#compare-result-rows .result-row')).to_have_count(1)
                            self.assertEqual(len(runs), 1)
                            # Changing a setting clears the suggested run so stale settings cannot run.
                            page.locator('#workflow-preview').click()
                            expect(page.locator('#workflow-plan')).to_be_visible()
                            page.locator('#workflow-options > summary').click()
                            page.locator('#workflow-rounding').select_option('')
                            expect(page.locator('#workflow-plan')).to_be_hidden()
                            expect(page.locator('#workflow-method-note')).to_contain_text('không làm tròn')
                            # A new file with ambiguous records gets a visible manual fallback.
                            page.locator('#compare-upload').set_input_files(paths[2])
                            expect(page.locator('#workbook-count')).to_have_text('3 file')
                            page.locator('#workflow-report').select_option(label='Ambiguous.xlsx')
                            page.locator('#workflow-preview').click()
                            expect(page.locator('#workflow-error')).to_be_visible()
                            expect(page.locator('#workflow-run')).to_be_hidden()
                            page.locator('#offline-add-comparison').click()
                            expect(page.locator('#compare-rules .check-rule')).to_have_count(1)
                            expect(page.locator('#advanced-checks')).to_have_attribute('open', '')
                            expect(page.locator('[data-field="right.workbook"]')).to_have_value(page.locator('#workflow-report').input_value())
                            # The new quick check is separate; the original primary action stays available.
                            page.locator('#compare-upload').set_input_files(paths[3:])
                            expect(page.locator('#workbook-count')).to_have_text('5 file')
                            page.locator('#workflow-source').select_option(label='Percent A.xlsx')
                            page.locator('#workflow-report').select_option(label='Percent B.xlsx')
                            expect(page.locator('#workflow-preview')).to_have_class('primary')
                            expect(page.locator('#workflow-preview')).to_have_text('Xem gợi ý so sánh →')
                            expect(page.locator('#quick-compare')).to_be_enabled()
                            page.locator('#quick-compare').click()
                            expect(page.locator('#quick-verdict')).to_be_visible()
                            expect(page.locator('#quick-verdict')).to_contain_text('Chênh lệch 1 dòng')
                            expect(page.locator('#quick-verdict')).to_contain_text('File B nhiều hơn 1 dòng')
                            expect(page.locator('#quick-row-counts')).to_be_visible()
                            expect(page.locator('#quick-row-counts tbody tr td').nth(1)).to_have_text('3')
                            expect(page.locator('#quick-row-counts tbody tr td').nth(2)).to_have_text('4')
                            expect(page.locator('#compare-result-rows .result-row')).to_have_count(3)
                            expect(page.locator('#compare-result-rows')).to_contain_text('94.40001%')
                            expect(page.locator('#compare-result-rows')).to_contain_text('94.40002%')
                            self.assertEqual(len(runs), 1)
                            self.assertEqual(len(quick_runs), 1)
                            if screenshots:
                                page.screenshot(path=str(target/f'quick-{width}.png'), full_page=True)
                            page.reload()
                            expect(page.locator('#quick-verdict')).to_be_visible()
                            expect(page.locator('#quick-verdict')).to_contain_text('Chênh lệch 1 dòng')
                            expect(page.locator('#compare-result-rows .result-row')).to_have_count(3)
                            self.assertEqual(len(quick_runs), 1)
                            self.assertFalse(page.evaluate('document.documentElement.scrollWidth > innerWidth'))
                            self.assertEqual(blocked, [])
                            self.assertEqual(errors, [])
                            context.close()
                        browser.close()
                finally:
                    http.shutdown()
                    http.server_close()
                    thread.join(timeout=5)
                    compare.load_workbook.cache_clear()


if __name__ == '__main__':
    unittest.main()
