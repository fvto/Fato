"""Search across mixed sources with exact totals and bounded result pages."""
from decimal import Decimal
from pathlib import Path
import io
import json
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
from datetime import datetime

import openpyxl
from werkzeug.datastructures import FileStorage
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import spreadsheet_combine as combine
import spreadsheet_combine_search as search
import sfc_dashboard_server as server


def workbook(rows):
    book=openpyxl.Workbook()
    for row in rows:book.active.append(row)
    output=io.BytesIO();book.save(output);book.close();return output.getvalue()


class SearchTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.addCleanup(self.temp.cleanup)
        patcher=patch.object(combine,'ROOT',Path(self.temp.name)/'jobs');patcher.start();self.addCleanup(patcher.stop)
        csv=b'Mold Code,Size,Number of Pairs,FTY\nM001,9.0,1.5,VH\nM001,9,2,VH\nM001,10,invalid,JV\nM002,9,3,VH\n'
        excel=workbook([['FTY','Number of Pairs','Size','Mold Code'],['VH',.25,10,'M001'],['JV',2,8,'M002']])
        self.job=combine.upload([FileStorage(io.BytesIO(csv),filename='source.csv'),FileStorage(io.BytesIO(excel),filename='source.xlsx')])
        self.choices=[dict(file_id=f['id'],sheet=f['sheets'][0]['name'],header_row=1) for f in self.job['files']]

    def test_mixed_sources_multiple_filters_size_total_and_invalid_values(self):
        created=search.create(self.job['id'],self.choices,{})
        self.assertEqual(created['grand_total'],'8.75')
        self.assertEqual(created['invalid_quantities'],1)
        indexes={name:i for i,name in enumerate(created['columns'])}
        filters=[dict(column=indexes['Mold Code'],mode='exact',value='m001'),dict(column=indexes['FTY'],mode='exact',value='vh')]
        result=search.query(self.job['id'],created['index_id'],filters)
        self.assertEqual((result['matched'],result['grand_total'],result['invalid_quantities']),(3,'3.75',0))
        totals={group['size']:Decimal(group['total']) for group in result['groups']}
        self.assertEqual(totals,{'10':Decimal('.25'),'9':Decimal('3.5')})
        self.assertEqual(result['rows'][0][indexes['Mold Code']],'M001')
        result=search.query(self.job['id'],created['index_id'],[dict(column=indexes['Mold Code'],mode='contains',value='001')])
        self.assertEqual((result['matched'],result['grand_total'],result['invalid_quantities']),(4,'3.75',1))
        result=search.query(self.job['id'],created['index_id'],[dict(column=indexes['Size'],mode='exact',value='9.0')])
        self.assertEqual((result['matched'],result['grand_total']),(3,'6.5'))
        with patch.object(search.sqlite3,'connect',side_effect=AssertionError('An existing index should be reused')):
            self.assertEqual(search.create(self.job['id'],self.choices,{})['index_id'],created['index_id'])

    def test_csv_and_excel_dates_filter_as_the_same_day(self):
        csv=b'Mold Code,Size,Number of Pairs,Date\nM001,9,1,2026-01-01\n'
        excel=workbook([['Mold Code','Size','Number of Pairs','Date'],['M001',9,2,datetime(2026,1,1)]])
        job=combine.upload([FileStorage(io.BytesIO(csv),filename='a.csv'),FileStorage(io.BytesIO(excel),filename='b.xlsx')])
        choices=[dict(file_id=f['id'],sheet=f['sheets'][0]['name'],header_row=1) for f in job['files']]
        index=search.create(job['id'],choices,{})
        result=search.query(job['id'],index['index_id'],[dict(column=index['columns'].index('Date'),mode='exact',value='2026-01-01')])
        self.assertEqual((result['matched'],result['grand_total']),(2,'3'))

    def test_shoe_code_whitespace_and_size_ranges_are_separate_totals(self):
        csv=('Mold Code,Size,Size Range,Number of Pairs\n'
             'OS398578-1,5,5C1C2,2\nOS 398578-1,5.0,5C1C2,1.5\n'
             'OS398578-1,5,5E1E2,7\nOS398578-1,8,8A1A2-VQ,3\nOS398578-1,8,8A1A2,4\n').encode()
        job=combine.upload([FileStorage(io.BytesIO(csv),filename='shoe.csv')])
        choices=[dict(file_id=0,sheet='CSV',header_row=1)]
        index=search.create(job['id'],choices,{})
        filters=[dict(column=index['columns'].index('Mold Code'),mode='exact',value='OS 398578-1')]
        result=search.query(job['id'],index['index_id'],filters)
        groups={(g['size'],g['size_range']):Decimal(g['total']) for g in result['groups']}
        self.assertEqual(groups,{('5','5C1C2'):Decimal('3.5'),('5','5E1E2'):Decimal('7'),('8','8A1A2'):Decimal('4'),('8','8A1A2-VQ'):Decimal('3')})
        self.assertEqual(result['matched'],5)

    def test_unicode_sql_literal_filters_and_pagination(self):
        data='Mold Code,Size,Number of Pairs,Name\n'+''.join(f'M001,{i%3},0.1,Đế {i:03}\n' for i in range(205))
        job=combine.upload([FileStorage(io.BytesIO(data.encode('utf-8-sig')),filename='unicode.csv')])
        choices=[dict(file_id=0,sheet='CSV',header_row=1)]
        created=search.create(job['id'],choices,{})
        self.assertEqual(created['grand_total'],'20.5')
        index=created['columns'].index('Name')
        result=search.query(job['id'],created['index_id'],[dict(column=index,mode='contains',value='đẾ')],page=2)
        self.assertEqual((result['matched'],len(result['rows']),result['grand_total']),(205,100,'20.5'))
        result=search.query(job['id'],created['index_id'],[dict(column=index,mode='contains',value="' OR 1=1 --")])
        self.assertEqual(result['matched'],0)
        for invalid in [[{'column':-1,'mode':'exact','value':'x'}],[{'column':0,'mode':'bad','value':'x'}]]:
            with self.assertRaises(ValueError):search.query(job['id'],created['index_id'],invalid)

    def test_endpoints_build_and_query_background_index(self):
        client=server.app.test_client()
        restored=client.get(f'/api/combine/{self.job["id"]}/inputs')
        self.assertEqual(restored.status_code,200)
        self.assertEqual(len(restored.json['files']),2)
        self.assertNotIn('stored',restored.json['files'][0])
        def finish(response):
            self.assertEqual(response.status_code,202,response.json)
            url=response.json['task_url'];deadline=time.monotonic()+15
            while time.monotonic()<deadline:
                status=client.get(url).json
                if status['status']=='done':return status['result']
                if status['status']=='error':self.fail(status['error'])
                time.sleep(.01)
            self.fail('Background index task timed out')
        created=finish(client.post(f'/api/combine/{self.job["id"]}/search-index',json={'selections':self.choices,'options':{}}))
        result=finish(client.post(f'/api/combine/{self.job["id"]}/search',json={'index_id':created['index_id'],'filters':[],'page':1}))
        self.assertEqual((result['matched'],result['grand_total']),(6,'8.75'))

    def test_top3_issues_use_only_drr_quantity_and_exact_size_range(self):
        text=('Mold Code,Size,Size Range,Number of Pairs,Unit Type,Issue\n'
              'OS398578-1,8,8A1A2,100,NON-DRR,Ignore\n'
              'OS398578-1,8,8A1A2,1,DRR,Bubble\n'
              'OS398578-1,8,8A1A2,1.5,drr,Bubble\n'
              'OS398578-1,8,8A1A2,5,DRR,Burning mark\n'
              'OS398578-1,8,8A1A2,2.5,DRR,Tear\n'
              'OS398578-1,8,8A1A2,1,DRR,Other\n'
              'OS398578-1,8,8A1A2-VQ,999,DRR,Exclude VQ\n'
              'OS398578-1,8,8A1A2,bad,DRR,Bad quantity\n')
        job=combine.upload([FileStorage(io.BytesIO(text.encode()),filename='drr.csv')])
        created=search.create(job['id'],[dict(file_id=0,sheet='CSV',header_row=1)],{})
        filters=[dict(column=created['columns'].index('Mold Code'),mode='exact',value='OS 398578-1'),
                 dict(column=created['columns'].index('Size Range'),mode='exact',value='8A1A2')]
        result=search.query(job['id'],created['index_id'],filters)
        self.assertEqual(result['grand_total'],'111.0')
        self.assertEqual([(x['rank'],x['issue'],x['total']) for x in result['top_issues']],[(1,'Burning mark','5'),(2,'Bubble','2.5'),(3,'Tear','2.5')])
        self.assertEqual(result['invalid_quantities'],1)


if __name__=='__main__':unittest.main()
