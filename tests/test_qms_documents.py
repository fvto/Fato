from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_qms as qms


LINK = 'DownloadFile.aspx?EinB64=ABC&amp;N=DEF'
PDF = b'%PDF-1.7\nfixture\n%%EOF\n'


def document_page(page=1, total=2, name='S0001', vn=True):
    links = f'<a id="hplFile_EN_0" href="{LINK}"><img></a>'
    vietnamese = f'<a id="hplFile_VN_0" href="{LINK}VN"><img></a>' if vn else '<a id="hplFile_VN_0"><img></a>'
    cells = ['', links, vietnamese, '2026-11-03', name, 'QA manual', 'HFPA', '', name + '; Shoe name', '', 'Auditor', '']
    return ('<html><form><input name="__VIEWSTATE" value="state' + str(page) + '"></form><table>'
            '<tr id="rpnlMain_ContentPlaceHolder1_gvData_DXDataRow0" class="dxgvFocusedRow">' +
            ''.join('<td>' + value + '</td>' for value in cells) + '</tr></table>' +
            f'<b class="dxp-summary">Page {page} of 2 ({total} items)</b></html>')


class CatalogTests(unittest.TestCase):
    def test_pdf_icons_and_focused_row_become_distinct_language_downloads(self):
        entries, current, pages, total, rows = qms.parse_pdf_page(document_page(), 'document', 'VH')
        self.assertEqual((current, pages, total, rows), (1, 2, 2, 1))
        self.assertEqual([x['language'] for x in entries], ['EN', 'VN'])
        self.assertEqual(entries[0]['name'], 'S0001; Shoe name')
        self.assertNotEqual(entries[0]['id'], entries[1]['id'])
        self.assertEqual(len(qms.parse_pdf_page(document_page(vn=False), 'document', 'VH')[0]), 1)

    def test_all_pages_use_fresh_form_state_and_reconcile_rows(self):
        class Session:
            plant = 'VH'
            def __init__(self): self.calls = []
            def fetch(self, url, fields=None):
                self.calls.append(fields)
                return (document_page(1) if fields is None else document_page(2, name='S0002')).encode(), {}
        session = Session()
        entries, stats = qms.crawl_documents(session, 'document')
        self.assertEqual(len(entries), 2)  # Same fixture file URLs are deduplicated across rows.
        self.assertEqual(stats['rows'], 2)
        self.assertEqual(session.calls[1]['__VIEWSTATE'], 'state1')
        self.assertEqual(session.calls[1]['__EVENTARGUMENT'], '12|PAGERONCLICK3|PN1')
        with patch.object(qms, 'parse_pdf_page', return_value=([], 1, 1, 2, 1)):
            with self.assertRaisesRegex(RuntimeError, 'chưa đủ dòng'):
                qms.crawl_documents(Session(), 'document')

    def test_navigation_excludes_account_links_and_preserves_new_codes(self):
        markup = '''<table><tr><td id="ASPxSiteMapControl1_0_C0">
        <div class="dxsmLevel0_Aqua">QA Document</div>
        <div class="dxsmLevel1_Aqua"><a title="QMS750" href="QADocument.aspx?id=QMS750">QA Document</a></div>
        </td></tr></table><a href="../Logout.aspx">Logout</a>'''
        modules = qms.parse_navigation(markup)
        self.assertEqual(modules[0]['code'], 'QMS750')
        self.assertEqual(modules[0]['category'], 'QA Document')
        self.assertFalse(any(m['name'] == 'Logout' for m in modules))

    def test_failed_crawl_leaves_both_saved_catalogs_intact(self):
        with tempfile.TemporaryDirectory() as folder:
            memory, catalog = Path(folder) / 'memory.json', Path(folder) / 'catalog.json'
            memory.write_text('{"old": true}')
            catalog.write_text('{"VH": {"old": true}}')
            with patch.object(qms, 'MEMORY', memory), patch.object(qms, 'CATALOG', catalog), \
                 patch.object(qms, 'DocumentSession'), patch.object(qms, 'parse_navigation', return_value=[{}]), \
                 patch.object(qms, 'crawl_documents', side_effect=RuntimeError('page failed')):
                with self.assertRaises(RuntimeError): qms.refresh('VH')
            self.assertEqual(json.loads(memory.read_text()), {'old': True})
            self.assertEqual(json.loads(catalog.read_text()), {'VH': {'old': True}})
            self.assertFalse(qms.REFRESH_LOCK.locked())


class PdfTests(unittest.TestCase):
    def setUp(self):
        self.entry = dict(id='a' * 32, source='document', name='Shoe', code='S0001', language='VN',
                          source_url=qms.SOURCES['document'][1].replace('QADocument', 'DownloadFile'))
        self.catalog = patch.object(qms, 'read_json', return_value={'VH': {'documents': [self.entry]}})
        self.catalog.start()
        self.addCleanup(self.catalog.stop)

    def test_viewer_resolves_original_pdf_with_encoded_spaces(self):
        with patch.object(qms, 'DocumentSession') as session:
            session.return_value.fetch.side_effect = [
                (b'<object type="application/pdf" data="/vhmes/QA/UploadFile/Test VN.pdf?id=20261110 090000"></object>', {}),
                (PDF, {})]
            stream, name = qms.download_pdf('VH', 'a' * 32)
            self.assertEqual(stream.read(), PDF)
            self.assertEqual(name, 'Test VN.pdf')
            self.assertIn('Test%20VN.pdf?id=20261110%20090000', session.return_value.fetch.call_args.args[0])

    def test_missing_or_partial_pdf_and_untrusted_embeds_are_rejected(self):
        for body in (b'<html>Login required</html>', b'%PDF-1.7\ntruncated',
                     b'<object type="application/pdf" data="https://evil.test/a.pdf"></object>',
                     b'<object type="application/pdf" data="/vhmes/QA/UploadFile/../a.pdf"></object>'):
            with self.subTest(body=body), patch.object(qms, 'DocumentSession') as session:
                session.return_value.fetch.return_value = (body, {})
                with self.assertRaises(RuntimeError): qms.download_pdf('VH', 'a' * 32)
                self.assertEqual(session.return_value.fetch.call_count, 1)

    def test_unknown_ids_never_request_sfc_and_catalog_hides_source_urls(self):
        with patch.object(qms, 'DocumentSession') as session:
            with self.assertRaises(FileNotFoundError): qms.download_pdf('VH', 'b' * 32)
            session.assert_not_called()
        public = qms.document_catalog('VH')['documents'][0]
        self.assertNotIn('source_url', public)
        self.assertEqual(public['url'], '/api/qms/documents/VH/' + 'a' * 32 + '/pdf')

    def test_api_attachment_errors_and_origin_protection(self):
        import sfc_dashboard_server as server
        client = server.app.test_client()
        with patch.object(qms, 'download_pdf', return_value=(BytesIO(PDF), 'S0001_VN.pdf')):
            response = client.get('/api/qms/documents/VH/' + 'a' * 32 + '/pdf')
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.mimetype, 'application/pdf')
            self.assertIn('attachment', response.headers['Content-Disposition'])
            self.assertEqual(response.headers['Cache-Control'], 'no-store')
        with patch.object(qms, 'download_pdf', side_effect=RuntimeError('expired')):
            self.assertEqual(client.get('/api/qms/documents/VH/' + 'a' * 32 + '/pdf').status_code, 409)
        self.assertEqual(client.get('/api/qms/documents/invalid/a/pdf').status_code, 404)
        self.assertEqual(client.get('/api/qms/documents?plant=invalid').status_code, 400)
        with patch.object(qms, 'refresh') as refresh:
            response = client.post('/api/qms/refresh', json={'plant': 'VH'}, headers={'Origin': 'https://evil.test'})
            self.assertEqual(response.status_code, 403)
            refresh.assert_not_called()


if __name__ == '__main__':
    unittest.main()
