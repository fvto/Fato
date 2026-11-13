"""Large-source CSV conversion, streaming parity, Excel splits and async work."""
from datetime import datetime
from pathlib import Path
import io
import json
import sys
import tempfile
import time
import unittest
from unittest.mock import patch

import openpyxl
from werkzeug.datastructures import FileStorage

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import spreadsheet_combine as combine
import spreadsheet_combine_large as large
import spreadsheet_combine_tasks as tasks
import spreadsheet_combine_xlsx as xlsx
import sfc_dashboard_server as server
from test_spreadsheet_combine_csv import csv_bytes,workbook_bytes,selections,storage


class LargeCombineTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        patcher=patch.object(combine,'ROOT',Path(self.temp.name)/'jobs')
        patcher.start();self.addCleanup(patcher.stop)

    def mark_large(self,job):
        path=combine.folder_for(job['id'])/'inputs.json'
        items=json.loads(path.read_text(encoding='utf-8'))
        for item in items:item['large']=True
        combine.json_write(path,items)

    def test_ascii_prefix_with_late_legacy_byte_is_fully_validated(self):
        path=Path(self.temp.name)/'late.csv'
        path.write_bytes(b'ID,Name\n'+b'001,ASCII\n'*15000+b'002,A\xa0B\n')
        config=combine.csv_input.configuration(path)
        self.assertEqual(config['encoding'],'cp1252')
        self.assertTrue(config['warnings'])
        self.assertEqual(list(combine.csv_input.iter_records(path,config))[-1],['002','A\u00a0B'])
        with self.assertRaises(ValueError):
            combine.csv_input.configuration(path,{'encoding':'utf-8-sig'})

    def test_csv_is_read_directly_and_literals_are_preserved(self):
        job=combine.upload([storage('Input.csv',csv_bytes([
            ['ID','Text','Rate %'],['001','=2+3','25.0%'],['002',"'prefix",'0%'],['003','first\r\nsecond\rthird','100%']]))])
        folder=combine.folder_for(job['id'])
        item=json.loads((folder/'inputs.json').read_text(encoding='utf-8'))[0]
        self.assertTrue(item['stored'].endswith('.csv'))
        self.assertFalse(list(folder.glob('*_converted.xlsx')))
        result=combine.prepare(job['id'],selections(job))
        self.assertEqual(result['rows'][0][2]['value'],.25)
        self.assertEqual([row[1]['value'] for row in result['rows']], ['=2+3', "'prefix", 'first\r\nsecond\rthird'])
        self.assertEqual(result['rows'][0][0]['value'],'001')

    def test_first_large_preview_reuses_upload_scan_without_changing_cleaning(self):
        content=csv_bytes([['ID','Text','Rate %'],['001','first\nsecond','25.0%'],[],
                           ['ID','Text','Rate %'],['002','A\u00a0B','0%']])
        job=combine.upload([storage('Scan.csv',content)])
        small=combine.prepare(job['id'],selections(job))
        self.mark_large(job)
        with patch.object(large,'records',side_effect=AssertionError('Do not rescan CSV for first preview')):
            big=combine.prepare(job['id'],selections(job))
        self.assertEqual(big['sources'],small['sources'])
        self.assertEqual(big['rows'][:12],small['rows'])
        self.assertEqual(list(big['rows']),small['rows'])
        path=combine.folder_for(job['id'])/'input_0.csv'
        with path.open('ab') as stream:stream.write(b'changed')
        with self.assertRaisesRegex(ValueError,'đã thay đổi'):
            combine.prepare(job['id'],selections(job))

    def test_upload_scan_respects_changed_header_range_and_footer_option(self):
        headers=['Factory Code','SO#','Audit Date',"Defective Q'ty(pair)","Total Defect Q'ty"]
        content=csv_bytes([['Report'],headers,['JV','001','2026-10-02',2,3],headers,
                           ['JV','002','2026-10-03',3,4],['','','2','5','']])
        job=combine.upload([storage('Audit.csv',content)])
        choices=[dict(file_id=0,sheet='CSV',header_row=2)]
        small=combine.prepare(job['id'],choices)
        self.mark_large(job)
        with patch.object(large,'records',side_effect=AssertionError('Default full range must use scan')):
            big=combine.prepare(job['id'],choices)
            keep=combine.prepare(job['id'],choices,{'footer_mode':'keep'})
        self.assertEqual(big['sources'],small['sources'])
        self.assertEqual(list(big['rows']),small['rows'])
        self.assertEqual((keep['row_count'],keep['footer_rows']),(3,0))
        with patch.object(large,'records',wraps=large.records) as reader:
            ranged=combine.prepare(job['id'],[dict(choices[0],end_row=3)])
            self.assertEqual(ranged['row_count'],1)
            self.assertTrue(reader.called)
        with patch.object(large,'records',wraps=large.records) as reader:
            changed=combine.prepare(job['id'],[dict(choices[0],header_row=4,end_row=5)])
            self.assertEqual(changed['row_count'],1)
            self.assertTrue(reader.called)

    def test_invalid_csv_after_sample_still_rejected_during_cached_scan(self):
        for tail in [b'bad,"unterminated',b'bad,\x01\n',b'bad,'+b'x'*32768]:
            with self.subTest(tail=tail[:20]),self.assertRaises(combine.UploadValidationError):
                combine.upload([storage('Broken.csv',b'ID,Name\n'+b'001,valid\n'*60+tail)])
        # Invalid layout can be corrected after upload, but cannot hide a later
        # malformed CSV record while completing the mandatory input checks.
        with self.assertRaises(combine.UploadValidationError):
            combine.upload([storage('Broken.csv',b'ID,Name\n001,value,extra\n'+b'001,valid\n'*60+b'bad,"unterminated')])

    def test_cached_scan_does_not_silently_accept_invalid_audit_footer(self):
        content=csv_bytes([['Factory Code','SO#','Audit Date',"Defective Q'ty(pair)","Total Defect Q'ty"],
                           ['JV','001','2026-10-02',2,3],['','','1','99','']])
        job=combine.upload([storage('Footer.csv',content)])
        self.mark_large(job)
        with self.assertRaisesRegex(ValueError,'không khớp'):
            combine.prepare(job['id'],selections(job))
        self.assertEqual(combine.prepare(job['id'],selections(job),{'footer_mode':'keep'})['row_count'],2)

    def test_more_than_500000_cells_are_streamed_and_preview_is_bounded(self):
        headers=['ID']+[f'Column{i}' for i in range(19)]
        data=csv_bytes([headers]+[[f'{i:06}']+['v']*19 for i in range(26000)])
        job=combine.upload([storage('Big.csv',data)])
        self.assertTrue(job['files'][0]['large'])
        result=combine.prepare(job['id'],selections(job))
        self.assertEqual((result['row_count'],result['column_count']),(26000,23))
        self.assertIsInstance(result['rows'],large.Rows)
        self.assertEqual(len(combine.summary(result)['preview']),12)
        self.assertEqual(next(iter(result['rows']))[0]['value'],'000000')
        # A repeat preview/run reuses counts only after rechecking source hashes.
        with patch.object(large,'records',side_effect=AssertionError('Preview must use verified cache')):
            again=combine.prepare(job['id'],selections(job))
            self.assertEqual(again['row_count'],26000)
        folder=combine.folder_for(job['id'])
        original=folder/'input_0.csv'
        with original.open('ab') as stream:stream.write(b'changed')
        with self.assertRaisesRegex(ValueError,'đã thay đổi'):
            combine.prepare(job['id'],selections(job))

    def test_result_split_keeps_every_row_and_original_csv_record(self):
        with patch.object(large,'MAX_ROWS',4),patch.object(xlsx,'MAX_ROWS',4),patch.object(combine,'MAX_CELLS',1):
            job=combine.upload([storage('Split.csv',csv_bytes([['ID']]+[[f'{i:03}'] for i in range(7)]))])
            sheets=job['files'][0]['sheets']
            self.assertEqual([s['rows'] for s in sheets],[8])
            choices=[dict(file_id=0,sheet=s['name'],header_row=s['header_row']) for s in sheets]
            payload=combine.prepare(job['id'],choices)
            self.assertEqual(payload['row_count'],7)
            self.assertEqual([row[-1]['value'] for row in payload['rows']],list(range(2,9)))
            result=combine.export(job['id'],choices)
            self.assertEqual(result['combined_sheets'],['Combined','Combined_2','Combined_3'])
            output=combine.folder_for(job['id'])/result['result_id']/result['output_name']
            book=openpyxl.load_workbook(output,read_only=True)
            try:
                values=[row[0] for name in result['combined_sheets'] for row in list(book[name].values)[1:]]
                self.assertEqual(values,[f'{i:03}' for i in range(7)])
            finally:book.close()

    def test_large_cleaning_matches_old_header_footer_and_value_rules(self):
        headers=['Factory Code','SO#','Audit Date',"Defective Q'ty(pair)","Total Defect Q'ty",'Rate %']
        rows=[headers,['JV','001',datetime(2026,10,2),2,3,'25.0%'],headers,['JV','002',datetime(2026,10,3),3,4,'0%'],['','','2','5','','']]
        job=combine.upload([storage('Audit.xlsx',workbook_bytes(rows))])
        small=combine.prepare(job['id'],selections(job))
        self.mark_large(job)
        big=combine.prepare(job['id'],selections(job))
        self.assertEqual(big['columns'],small['columns'])
        self.assertEqual(list(big['rows']),small['rows'])
        self.assertEqual(big['sources'],small['sources'])
        kept=combine.prepare(job['id'],selections(job),{'footer_mode':'keep'})
        self.assertEqual((kept['row_count'],kept['footer_rows']),(3,0))
        result=combine.export(job['id'],selections(job))
        output=combine.folder_for(job['id'])/result['result_id']/result['output_name']
        book=openpyxl.load_workbook(output)
        try:
            self.assertEqual(book['Combined']['C2'].value,datetime(2026,10,2))
            self.assertEqual((book['Combined']['F2'].value,book['Combined']['F2'].number_format),(.25,'0.0%'))
            self.assertEqual(book['Combined'].freeze_panes,'A2')
        finally:book.close()

    def test_background_preview_and_export_report_progress_and_complete(self):
        job=combine.upload([storage('async.csv',b'ID,Name\n001,A\n002,B\n')])
        self.mark_large(job)
        client=server.app.test_client()
        def finish(response):
            self.assertEqual(response.status_code,202,response.json)
            deadline=time.monotonic()+15
            task_url=response.json['task_url']
            while time.monotonic()<deadline:
                status=client.get(task_url).json
                if status['status']=='done':return status['result']
                if status['status']=='error':self.fail(status['error'])
                time.sleep(.01)
            self.fail('Background work timed out')
        preview=finish(client.post(f'/api/combine/{job["id"]}/preview',json={'selections':selections(job)}))
        self.assertEqual(preview['row_count'],2)
        result=finish(client.post(f'/api/combine/{job["id"]}/run',json={'selections':selections(job)}))
        response=client.get(result['download_url'])
        try:self.assertEqual(response.status_code,200)
        finally:response.close()

    def test_staged_background_upload_reads_csv_and_cleans_incoming_files(self):
        task=tasks.stage_upload([storage('stage.csv',b'ID,Name\n001,A\n')],{})
        deadline=time.monotonic()+15
        while time.monotonic()<deadline:
            status=tasks.snapshot(task['task_id'])
            if status['status']=='error':self.fail(status['error'])
            if status['status']=='done':break
            time.sleep(.01)
        self.assertEqual(status['status'],'done')
        self.assertNotIn('converted',status['result']['files'][0])
        self.assertFalse(list(combine.folder_for(status['result']['id']).glob('*_converted.xlsx')))
        self.assertEqual(list((combine.ROOT/'_incoming').iterdir()),[])


if __name__=='__main__':unittest.main()
