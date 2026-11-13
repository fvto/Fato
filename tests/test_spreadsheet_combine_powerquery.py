from pathlib import Path
import csv,hashlib,io,json,sys,tempfile,unittest,zipfile
from unittest.mock import patch
from werkzeug.datastructures import FileStorage
import openpyxl
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import spreadsheet_combine as combine
import spreadsheet_combine_powerquery as pq
import sfc_dashboard_server as server


class PowerQueryTests(unittest.TestCase):
    def test_threshold_is_strict_and_large_export_avoids_xlsx(self):
        base=dict(rows=[],row_count=1_000_000,columns=[])
        self.assertEqual(combine.summary(base)['export_mode'],'xlsx')
        self.assertEqual(combine.summary(dict(base,row_count=1_000_001))['export_mode'],'power_query')
        with (patch.object(combine,'prepare',return_value=dict(base,row_count=1_000_001,output_name='Large.xlsx')),
             patch.object(pq,'export',return_value={'kind':'power_query'}) as exporter,
             patch.object(combine,'runtime',side_effect=AssertionError('XLSX must not be used'))):
            self.assertEqual(combine.export('job',[])['kind'],'power_query')
            exporter.assert_called_once()

    def test_package_preserves_all_rows_headers_bom_and_download(self):
        with tempfile.TemporaryDirectory() as tmp,patch.object(combine,'ROOT',Path(tmp)),patch.object(combine,'XLSX_ROW_THRESHOLD',2),patch.object(pq,'CHUNK_ROWS',2):
            content='Mold Code,Size,Number of Pairs,Text\n001,5,1.5,"a,b"\n002,6,2,=2+3\n003,7,3,Đế\n'.encode('utf-8-sig')
            job=combine.upload([FileStorage(io.BytesIO(content),filename='raw.csv')])
            choices=[dict(file_id=0,sheet='CSV',header_row=1)]
            result=combine.export(job['id'],choices,'Test Power Query')
            self.assertEqual(result['kind'],'power_query')
            client=server.app.test_client();response=client.get(result['download_url'])
            try:
                self.assertEqual(response.status_code,200)
                self.assertIn('.zip',response.headers['Content-Disposition'])
                with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
                    self.assertIsNone(archive.testzip())
                    manifest=json.loads(archive.read('manifest.json'))
                    self.assertEqual([p['rows'] for p in manifest['parts']],[2,1])
                    rows=[]
                    for part in manifest['parts']:
                        data=archive.read('Data/'+part['file'])
                        self.assertEqual(hashlib.sha256(data).hexdigest(),part['sha256'])
                        table=list(csv.reader(io.StringIO(data.decode('utf-8-sig'))))
                        self.assertEqual(table[0],result['columns']);rows+=table[1:]
                    self.assertEqual([r[0] for r in rows],['001','002','003'])
                    self.assertEqual(rows[1][3],'=2+3')
                    self.assertIn('Only Create Connection',archive.read('HUONG_DAN.txt').decode('utf-8-sig'))
                    self.assertIn('Number.FromText',archive.read('Combine.pq').decode('utf-8-sig'))
                    book=openpyxl.load_workbook(io.BytesIO(archive.read('Huong_dan.xlsx')))
                    try:self.assertIn('SourceFolder',book.defined_names)
                    finally:book.close()
            finally:response.close()


if __name__=='__main__':unittest.main()
