import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet_compare as compare
import spreadsheet_chat as chat
from spreadsheet_chat_tools import AnalysisTools, artifact_csv
import sfc_dashboard_server as server


class ChatAnalysisTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.patch = patch.object(compare, 'ROOT', self.root)
        self.patch.start()
        self.ids = []

    def tearDown(self):
        compare.load_workbook.cache_clear()
        self.patch.stop()
        self.temp.cleanup()

    def attach(self, rows, sheet='Unusual heading'):
        book = openpyxl.Workbook()
        book.active.title = sheet
        for row in rows:
            book.active.append(row)
        stream = io.BytesIO()
        book.save(stream)
        book.close()
        stream.seek(0)
        result = compare.import_file(stream, 'Arbitrary.xlsx')
        self.ids.append(result['id'])
        return result['id']

    def tools(self):
        manifest = self.root / 'job' / 'manifest.json'
        compare.write_json(manifest, {'files': [{'id': book_id} for book_id in self.ids]})
        return AnalysisTools(manifest)

    def test_dynamic_sql_keeps_zeros_missing_values_and_original_rows(self):
        book = self.attach([['arbitrary report'], [], ['Key','First','Second','Claim'],
                            ['X',0,80,80], ['Y','N/A',90,90]])
        result = self.tools().query_data([{'workbook':book,'sheet':'Unusual heading','range':'A4:D5'}],
                                        'SELECT _row,A,D,MEAN(B,C) AS expected FROM t1 WHERE NUM(D) != MEAN(B,C)')
        self.assertEqual(result['rows'], [[4,'X',80,40.0]])
        stored = json.loads((self.root/'job/artifacts'/(result['artifact_id']+'.json')).read_text(encoding='utf-8'))
        self.assertEqual(stored['sources'][0]['range'], 'A4:D5')
        self.assertIn('MEAN', stored['sql'])

    def test_query_joins_reordered_columns_and_rows(self):
        a = self.attach([['Bay-9',10],['Bay-2',20]], 'Before')
        b = self.attach([[25,'Bay-2'],[10,'Bay-9']], 'After')
        result = self.tools().query_data([{'workbook':a,'sheet':'Before','range':'A1:B2'},
                                         {'workbook':b,'sheet':'After','range':'A1:B2'}],
            'SELECT t1.A,t1._row,t2._row,t1.B,t2.A FROM t1 JOIN t2 ON t1.A=t2.B WHERE t1.B != t2.A')
        self.assertEqual(result['rows'], [['Bay-2',2,1,20,25]])

    def test_queries_cannot_write_attach_or_read_unattached_books(self):
        book = self.attach([[2,3]])
        tools = self.tools()
        sources = [{'workbook':book,'sheet':'Unusual heading','range':'A1:B1'}]
        for sql in ['DELETE FROM t1', "ATTACH DATABASE ':memory:' AS other", 'PRAGMA database_list', "SELECT load_extension('x')"]:
            with self.subTest(sql=sql), self.assertRaises(ValueError):
                tools.query_data(sources,sql)
        with self.assertRaises(ValueError):
            tools.query_data([{'workbook':'0'*32,'sheet':'Unusual heading','range':'A1'}], 'SELECT * FROM t1')
        self.assertFalse((self.root/'job/evidence.jsonl').exists())

    def test_uncached_formulas_and_invalid_text_do_not_become_zero(self):
        book = self.attach([['=2+3', 'bad']])
        tools = self.tools()
        with self.assertRaisesRegex(ValueError,'no saved formula result'):
            tools.query_data([{'workbook':book,'sheet':'Unusual heading','range':'A1:B1'}], 'SELECT NUM(A) FROM t1')
        with self.assertRaises(ValueError):
            tools.query_data([{'workbook':book,'sheet':'Unusual heading','range':'B1'}], 'SELECT NUM(B) FROM t1')

    def test_calculation_evidence_uses_attached_workbook_allowlist(self):
        book = self.attach([['Key','One','Two','Mean'],['X',0,80,80]])
        tools = self.tools()
        rule = {'kind':'calculation','name':'Dynamic mean','left':{'workbook':book,'sheet':'Unusual heading'},
                'operation':'average','direction':'rows','source_range':'B2:C2','target_range':'D2','number_mode':'number','tolerance':'0'}
        result = tools.verify_checks([rule])
        self.assertEqual(result['issue_count'],1)
        self.assertEqual(result['issues'][0]['expected'],40)
        rule['left']['workbook'] = '0'*32
        with self.assertRaises(ValueError):
            tools.verify_checks([rule])

    def test_query_truncation_is_explicit_and_full_rows_are_saved(self):
        book = self.attach([[i] for i in range(250)])
        result = self.tools().query_data([{'workbook':book,'sheet':'Unusual heading','range':'A1:A250'}], 'SELECT _row,A FROM t1')
        self.assertEqual(len(result['rows']),200)
        self.assertEqual(result['total_rows'],250)
        self.assertTrue(result['truncated'])
        artifact = json.loads((self.root/'job/artifacts'/(result['artifact_id']+'.json')).read_text(encoding='utf-8'))
        self.assertEqual(len(artifact['rows']),250)

    def test_cell_details_preserve_number_formats_and_formula_text(self):
        book = self.attach([['=1/3', '33.3%']])
        result = self.tools().cell_details(book,'Unusual heading','A1:B1')
        self.assertEqual(result['cells'][0]['formula'],'=1/3')
        self.assertIsNone(result['cells'][0]['value'])
        self.assertEqual(result['cells'][0]['format'],'General')
        self.assertEqual(result['cells'][1]['value'],'33.3%')

    def test_csv_neutralizes_formula_strings(self):
        result = artifact_csv({'columns':['=HEADER','Value'],'rows':[['@SUM(A1)',5],['normal',-4]]})
        self.assertIn("'=HEADER",result)
        self.assertIn("'@SUM(A1)",result)
        self.assertIn('normal,-4',result)

    def test_inference_command_confines_tools_and_does_not_change_global_config(self):
        with patch.object(chat,'codex_path',return_value='codex'):
            args = chat.command('a'*32)
        self.assertIn('--ignore-user-config',args)
        self.assertIn('--ephemeral',args)
        self.assertIn('read-only',args)
        self.assertIn('shell_tool',args)
        self.assertIn('web_search="disabled"',args)
        self.assertTrue(any('--allow-dir' in arg and '--read-only' in arg for arg in args))
        self.assertNotIn('--dangerously-bypass-approvals-and-sandbox',args)

    def test_comparison_cards_survive_execution_reload_and_followup_context(self):
        book = self.attach([['Section', 'Metric', 'Value'], ['Storage', 'Score', 90]])
        finding = {'section': 'Storage', 'item': 'Score', 'status': 'difference',
                   'left': {'label': 'Reported', 'file': 'Arbitrary.xlsx', 'section': 'Storage',
                            'value': '90', 'location': "'Unusual heading'!C2"},
                   'right': {'label': 'Recalculated', 'file': 'Arbitrary.xlsx', 'section': 'Storage',
                             'value': '95', 'location': "'Unusual heading'!C2"},
                   'explanation': 'Test fixture for presentation persistence.'}
        final = {'answer': 'One section differs.', 'answer_language': 'en',
                 'comparisons': [finding], 'checked_scope': ['Storage'], 'needs_clarification': False}
        job_id, chat_id = 'c'*32, 'd'*32
        chat.save_chat({'id': chat_id, 'active_job': job_id, 'messages': [
            {'role': 'assistant', 'content': 'Prior check', 'comparisons': [finding]},
            {'role': 'user', 'content': 'Explain this section', 'job_id': job_id, 'files': [book]}]})
        compare.write_json(chat.folder(job_id)/'job.json', {
            'id': job_id, 'chat_id': chat_id, 'status': 'queued', 'files': [book]})
        process = MagicMock(returncode=0)
        process.poll.return_value = 0
        process.communicate.return_value = (json.dumps({'type': 'item.completed', 'item': {
            'type': 'agent_message', 'text': json.dumps(final)}}), '')
        self.assertTrue(chat.CAPACITY.acquire(blocking=False))
        with patch.object(chat, 'command', return_value=['test']), patch.object(chat.subprocess, 'Popen', return_value=process):
            chat.execute(job_id)
        restored = server.app.test_client().get('/api/compare/chat/conversations/'+chat_id).get_json()
        self.assertEqual(restored['messages'][-1]['comparisons'], [finding])
        self.assertEqual(restored['messages'][-1]['answer_language'], 'en')
        self.assertEqual(restored['jobs'][job_id]['comparisons'], [finding])
        prompt = process.communicate.call_args.args[0]
        self.assertIn('"comparisons": [{"section": "Storage"', prompt)
        self.assertIn('comparisons', chat.SCHEMA['required'])

    def test_chat_submission_is_idempotent_even_without_received_chat_id(self):
        book = self.attach([[1,2]])
        payload = {'workbooks':[book], 'message':'Check these values', 'request_id':'b'*32}
        job_id = None
        try:
            with patch.object(chat,'status',return_value={'ready':True}), patch.object(chat.threading,'Thread') as thread:
                first = server.app.test_client().post('/api/compare/chat/messages',json=payload)
                self.assertEqual(first.status_code,202)
                job_id = first.get_json()['id']
                second = server.app.test_client().post('/api/compare/chat/messages',json=payload)
                self.assertEqual(second.get_json()['id'],job_id)
                self.assertEqual(thread.call_count,1)
                restored = server.app.test_client().get('/api/compare/chat/conversations/'+first.get_json()['chat_id']).get_json()
                self.assertEqual(len(restored['messages']),1)
                self.assertEqual(thread.call_count,1)
        finally:
            if job_id:
                chat.ACTIVE.discard(job_id)
                chat.CAPACITY.release()

    def test_chat_routes_validate_local_origin_and_artifact_paths(self):
        client = server.app.test_client()
        response = client.post('/api/compare/chat/messages',json={},headers={'Origin':'https://other.example'})
        self.assertEqual(response.status_code,403)
        self.assertEqual(client.get('/api/compare/chat/jobs/not-an-id').status_code,400)
        self.assertEqual(client.get('/api/compare/chat/jobs/'+'a'*32+'/artifacts/not-an-id.csv').status_code,400)

    def test_manual_chatgpt_import_saves_cards_without_inference(self):
        book = self.attach([['Factory', 'Score'], ['VH', 90]])
        final = {'answer': 'Imported summary', 'answer_language': 'vi', 'comparisons': [{
            'section': 'Factory', 'item': 'Score', 'status': 'review', 'explanation': 'Check method',
            'left': {'label': 'Summary', 'file': 'Arbitrary.xlsx', 'section': 'VH', 'value': '90', 'location': 'Data!B2'},
            'right': {'label': 'Detail', 'file': 'Arbitrary.xlsx', 'section': 'VH', 'value': '95', 'location': 'Detail!C3'}}],
            'checked_scope': ['Reported by ChatGPT'], 'needs_clarification': True}
        client = server.app.test_client()
        with patch.object(chat, 'status', side_effect=AssertionError('Must not check Codex')), patch.object(chat, 'execute') as execute:
            response = client.post('/api/compare/chat/import', json={'workbooks': [book], 'answer': final})
            self.assertEqual(response.status_code, 201, response.get_json())
            result = response.get_json()
            self.assertEqual(result['source'], 'chatgpt_web')
            execute.assert_not_called()
        restored = client.get('/api/compare/chat/conversations/'+result['chat_id']).get_json()
        self.assertEqual(restored['messages'][-1]['comparisons'], final['comparisons'])
        self.assertEqual(restored['messages'][-1]['source'], 'chatgpt_web')
        self.assertEqual(restored['messages'][0]['files'], [book])
        self.assertEqual(result['evidence'], [])
        for changes in [{'answer_language': []}, {'comparisons': [{'status': []}]},
                        {'comparisons': [{'status': 'match'}]}, {'needs_clarification': 'false'}]:
            rejected = client.post('/api/compare/chat/import', json={'workbooks': [book], 'answer': {**final, **changes}})
            self.assertEqual(rejected.status_code, 400, rejected.get_json())
        self.assertEqual(client.post('/api/compare/chat/import', json=[]).status_code, 400)
        self.assertEqual(client.post('/api/compare/chat/import', json={'workbooks': [book], 'answer': final},
                                     headers={'Origin': 'https://other.example'}).status_code, 403)

    def test_original_workbook_download_preserves_bytes(self):
        book = self.attach([['Score', '=1+2']])
        original = (self.root/'inputs'/book/'source.xlsx').read_bytes()
        client = server.app.test_client()
        response = client.get('/api/compare/workbooks/'+book+'/download')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, original)
        self.assertIn('attachment', response.headers['Content-Disposition'])
        response.close()
        self.assertEqual(client.get('/api/compare/workbooks/not-an-id/download').status_code, 400)


if __name__ == '__main__':
    unittest.main()
