"""Provider switching and saved Antigravity results in the compare UI."""
import os
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

import openpyxl
from werkzeug.serving import make_server

import spreadsheet_compare as compare
import spreadsheet_chat as chat
import sfc_dashboard_server as server


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS') == '1', 'Set CHINGLUH_UI_TESTS=1 for Chrome')
class ProviderBrowserTests(unittest.TestCase):
    def test_switch_provider_send_and_reload(self):
        from playwright.sync_api import sync_playwright, expect
        ready = {'codex': True, 'antigravity': False}
        submitted = []

        def status(provider='codex'):
            return {'ready': ready[provider], 'message': provider + (' ready' if ready[provider] else ' sign in')}

        def execute(job_id):
            try:
                with chat.LOCK:
                    job = chat.read_job(job_id)
                    submitted.append(job['provider'])
                    final = {'answer': 'Đã kiểm tra mẫu.', 'answer_language': 'vi',
                             'comparisons': [], 'needs_clarification': False, 'checked_scope': ['Data!A1:B2']}
                    job.update(status='completed', stage='Complete', **final)
                    compare.write_json(chat.folder(job_id)/'job.json', job)
                    conversation = chat.read_chat(job['chat_id'])
                    conversation['messages'].append({'role': 'assistant', 'content': final['answer'],
                        'provider': job['provider'], 'job_id': job_id, 'comparisons': [], 'checked_scope': final['checked_scope']})
                    conversation['active_job'] = None
                    chat.save_chat(conversation)
            finally:
                chat.ACTIVE.discard(job_id)
                chat.CAPACITY.release()

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            workbook = openpyxl.Workbook()
            workbook.active.title = 'Data'
            workbook.active.append(['Before', 'After'])
            workbook.active.append([10, 15])
            path = root/'Sample.xlsx'
            workbook.save(path)
            workbook.close()
            with patch.object(compare, 'ROOT', root/'compare'), patch.object(compare, 'EXPORTS', root/'exports'), \
                 patch.object(chat, 'status', side_effect=status), patch.object(chat, 'execute', side_effect=execute):
                http = make_server('127.0.0.1', 0, server.app, threaded=True)
                threading.Thread(target=http.serve_forever, daemon=True).start()
                try:
                    with sync_playwright() as playwright:
                        browser = playwright.chromium.launch(channel='chrome', headless=True)
                        page = browser.new_page(viewport={'width': 1280, 'height': 1000})
                        errors = []
                        page.on('pageerror', lambda error: errors.append(str(error)))
                        page.goto(f'http://127.0.0.1:{http.server_port}/compare')
                        page.locator('#compare-upload').set_input_files(str(path))
                        expect(page.locator('#workbook-count')).to_have_text('1 file')
                        page.locator('#chat-mode').select_option('codex')
                        expect(page.locator('#chat-connection')).to_have_text('codex ready')
                        page.locator('#chat-request').fill('So sánh hai cột')
                        expect(page.locator('#chat-send')).to_be_enabled()
                        page.locator('#chat-mode').select_option('antigravity')
                        expect(page.locator('#chat-connection')).to_have_text('antigravity sign in')
                        expect(page.locator('#chat-send')).to_be_disabled()
                        ready['antigravity'] = True
                        page.locator('#chat-mode').select_option('offline')
                        page.locator('#chat-mode').select_option('antigravity')
                        expect(page.locator('#chat-connection')).to_have_text('antigravity ready')
                        expect(page.locator('#chat-disclosure')).to_contain_text('tài khoản Antigravity')
                        page.locator('#chat-send').click()
                        expect(page.locator('#chat-messages')).to_contain_text('Đã kiểm tra mẫu.')
                        expect(page.locator('#chat-messages')).to_contain_text('ANTIGRAVITY')
                        page.reload()
                        expect(page.locator('#chat-mode')).to_have_value('antigravity')
                        expect(page.locator('#chat-messages')).to_contain_text('Đã kiểm tra mẫu.')
                        self.assertEqual(submitted, ['antigravity'])
                        self.assertEqual(errors, [])
                        browser.close()
                finally:
                    http.shutdown()
                    http.server_close()
                    compare.load_workbook.cache_clear()


if __name__ == '__main__':
    unittest.main()
