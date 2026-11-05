import io
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import openpyxl

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import spreadsheet_compare as compare
import spreadsheet_workflows as workflows
import sfc_dashboard_server as server


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.root=Path(self.temp.name)
        self.patch=patch.object(compare,'ROOT',self.root)
        self.patch.start()

    def tearDown(self):
        compare.load_workbook.cache_clear()
        self.patch.stop()
        self.temp.cleanup()

    def import_sheets(self,sheets,name):
        book=openpyxl.Workbook()
        book.remove(book.active)
        for title,rows in sheets.items():
            sheet=book.create_sheet(title)
            for row in rows: sheet.append(row)
        data=io.BytesIO()
        book.save(data)
        book.close()
        data.seek(0)
        return compare.import_file(data,name)

    def mqaa(self,period='From: 2026-10-19  To: 2026-10-24'):
        header=['Audit Date','Audit form','Factory','Line/Machine','MQAA%','Internal factory']
        rows=[header,['2026-10-19','Form A','VH','Line 1','100.0%','F1'],
              ['2026-10-19','Form B','VH','Line 1','0.0%','F1'],
              ['2026-10-20','Form A','VH','Line 2','94.4%','F2']]
        source=self.import_sheets({'Combined':rows+[['2026-10-20','Form A','VH2','Line 1','10.0%','Other plant']]},'Source.xlsx')
        summary=[['MQAA REPORT - Line'],[period],[],['','Line','Form A','Form B','Average'],
                 ['F1','Line 1','100.0%','0.0%','50.0%'],['F2','Line 2','94.4%','N/A','N/A'],
                 ['GRAND TOTAL','','97.2%','0.0%','']]
        report=self.import_sheets({'Raw QC':rows,'SUMMARY':summary},'QC.xlsx')
        return source,report

    def test_mqaa_preview_separates_scope_preserves_zero_and_finds_missing_average(self):
        source,report=self.mqaa()
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id']})
        self.assertEqual(plan['profile'],'mqaa')
        self.assertEqual(plan['scope']['factories'],['VH'])
        self.assertEqual(plan['scope']['source_rows'],4)
        self.assertEqual(plan['scope']['scoped_rows'],3)
        self.assertFalse((self.root/'results').exists())
        result=workflows.execute(plan['id'],True)
        issue=[row for row in result['results'] if row['status']=='missing_value']
        self.assertEqual([(row['cell'],row['expected']) for row in issue],[('E6',.944)])
        averages=[row for row in result['results'] if row['rule_index']==1 and row['cell']=='E5']
        self.assertEqual(averages[0]['expected'],.5)
        source_check=[row for row in result['results'] if row.get('other_cell')=='C5'][0]
        self.assertEqual(source_check['expected'],1)
        self.assertEqual(result['workflow']['reviewer'],'operator')

    def test_mqaa_detects_period_mismatch_and_reversed_roles(self):
        source,report=self.mqaa('From: 2026-10-26  To: 2026-10-31')
        with self.assertRaisesRegex(ValueError,'ngoài kỳ'):
            workflows.preview({'source_id':source['id'],'report_id':report['id']})
        with self.assertRaisesRegex(ValueError,'bị đảo'):
            workflows.preview({'source_id':report['id'],'report_id':source['id']})

    def test_next_week_renamed_sheets_and_reordered_source_columns_are_discovered(self):
        raw=[['From: 2026-10-26  To: 2026-10-31'],[],
             ['MQAA%','Line/Machine','Factory','Internal factory','Audit form','Audit Date'],
             ['90.0%','New Line','VH','F1','New Form','2026-10-26']]
        source=self.import_sheets({'Source week 44':raw},'Week44_source.xlsx')
        report=self.import_sheets({'QC raw week 44':raw,
            'Matrix week 44':[['MQAA REPORT - Line'],['From: 2026-10-26  To: 2026-10-31'],[],
                              ['Line','New Form','Average'],['New Line','90.0%','90.0%'],['GRAND TOTAL','90.0%','']]},'Week44_QC.xlsx')
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id']})
        self.assertEqual(plan['scope']['period'],['2026-10-26','2026-10-31'])
        result=workflows.execute(plan['id'],True)
        self.assertEqual(result['counts'],{'match':3})

    def test_does_not_deduplicate_repeated_audit_signatures(self):
        source,report=self.mqaa()
        source=self.import_sheets({'Combined':[
            ['Audit Date','Audit form','Factory','Line/Machine','MQAA%','Internal factory'],
            ['2026-10-19','Form A','VH','Line 1','100.0%','F1'],
            ['2026-10-19','Form A','VH','Line 1','100.0%','F1'],
            ['2026-10-20','Form A','VH','Line 1','0.0%','F1']]},'Repeated.xlsx')
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id'],'comparison_rounding_digits':None})
        self.assertEqual(plan['scope']['repeated_signature_rows'],1)
        result=workflows.execute(plan['id'],True)
        finding=next(row for row in result['results'] if row.get('other_cell')=='C5')
        self.assertAlmostEqual(finding['expected'],2/3)
        self.assertEqual(finding['valid_count'],3)

    def test_generic_maps_reordered_columns_and_reordered_records_by_headers(self):
        source=self.import_sheets({'Data':[['ID','Name','Score'],['A','Alpha',10],['B','Beta',20]]},'A.xlsx')
        report=self.import_sheets({'Data':[['Score','ID','Name'],[21,'B','Beta'],[10,'A','Alpha']]},'B.xlsx')
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id']})
        self.assertEqual(plan['profile'],'tables')
        self.assertEqual(plan['rules'][0]['left_keys'],'A')
        self.assertEqual(plan['rules'][0]['right_keys'],'B')
        result=workflows.execute(plan['id'],True)
        self.assertEqual(result['counts'],{'match':3,'difference':1})
        issue=next(row for row in result['results'] if row['status']=='difference')
        self.assertEqual((issue['cell'],issue['other_cell']),('C3','A2'))

    def test_generic_ambiguous_keys_require_manual_mapping(self):
        rows=[['Factory','Score'],['VH',10],['VH',20]]
        source=self.import_sheets({'Data':rows},'A.xlsx')
        report=self.import_sheets({'Data':rows},'B.xlsx')
        with self.assertRaisesRegex(ValueError,'quy tắc thủ công'):
            workflows.preview({'source_id':source['id'],'report_id':report['id']})

    def test_run_requires_review_and_rejects_stale_input(self):
        source,report=self.mqaa()
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id']})
        with self.assertRaisesRegex(ValueError,'Xác nhận'):
            workflows.execute(plan['id'])
        path=self.root/'inputs'/source['id']/'source.xlsx'
        path.write_bytes(path.read_bytes()+b'changed after review')
        with self.assertRaisesRegex(ValueError,'đã thay đổi'):
            workflows.execute(plan['id'],True)

    def test_workflow_api_preserves_local_origin_protection_and_report_metadata(self):
        source,report=self.mqaa()
        payload={'source_id':source['id'],'report_id':report['id']}
        with server.app.test_client() as client:
            self.assertEqual(client.post('/api/compare/workflows/preview',json=payload,headers={'Origin':'https://outside.example'}).status_code,403)
            response=client.post('/api/compare/workflows/preview',json=payload)
            self.assertEqual(response.status_code,200,response.get_json())
            plan=response.get_json()
            endpoint='/api/compare/workflows/'+plan['id']+'/run'
            self.assertEqual(client.post(endpoint,json={}).status_code,400)
            done=client.post(endpoint,json={'acknowledged':True})
            self.assertEqual(done.status_code,200,done.get_json())
            data=done.get_json()
            self.assertEqual(data['workflow']['profile'],'mqaa')
            self.assertEqual(data['issue_count'],1)
            self.assertTrue(data['html_url'].endswith('download.html'))

    def test_report_uses_qc_sheet_row_and_column_order_even_when_rules_are_reversed(self):
        from spreadsheet_report import prepare
        source=self.import_sheets({'Zulu':[['ID','Score'],['a',1],['b',2]],'Alpha':[['ID','Score'],['c',3]]},'Source.xlsx')
        report=self.import_sheets({'Zulu':[['ID','Score'],['a',11],['b',12]],'Alpha':[['ID','Score'],['c',13]]},'QC.xlsx')
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id']})
        result=compare.run(list(reversed(plan['rules'])))
        model=prepare(result,images=False)
        self.assertEqual([g['sheet'] for g in model['groups']],['Zulu','Alpha'])
        self.assertEqual([row['location'] for g in model['groups'] for row in g['rows']],['Zulu!B2','Zulu!B3','Alpha!B2'])

    def test_three_decimal_display_does_not_change_exact_comparison(self):
        from spreadsheet_report import display_value,exact_value,delta_value
        a=compare.Decimal('0.93678456789')
        b=compare.Decimal('0.93678456788')
        row={'expected_exact':str(a),'number_mode':'percent'}
        self.assertEqual(display_value(exact_value(row,'source')),'93.678%')
        self.assertEqual(display_value(str(b*100)+'%'),'93.678%')
        self.assertEqual(display_value('97.600%'),'97.6%')
        self.assertEqual(display_value('100.000%'),'100%')
        self.assertEqual(display_value('-0.100%'),'-0.1%')
        self.assertEqual(compare.compare_numbers(a,b,{},'percent',compare.Decimal(0),None)[0],'difference')
        self.assertNotEqual(display_value('1E-17%',is_delta=True),'0.000%')
        self.assertEqual(display_value(delta_value({'delta_exact':'-0.00025','number_mode':'percent'}),is_delta=True),'-0.025%')

    def test_workflow_records_rounding_precision_and_explains_it(self):
        source,report=self.mqaa()
        plan=workflows.preview({'source_id':source['id'],'report_id':report['id'],'comparison_rounding_digits':'1'})
        self.assertEqual(plan['comparison_rounding_digits'],1)
        self.assertTrue(all(rule['comparison_rounding_digits']==1 for rule in plan['rules']))
        self.assertTrue(any('ROUND_HALF_UP' in warning for warning in plan['warnings']))
        result=workflows.execute(plan['id'],True)
        self.assertEqual(result['comparison_policy'],'rounded')
        self.assertEqual(result['comparison_rounding_digits'],[1])


if __name__=='__main__': unittest.main()
