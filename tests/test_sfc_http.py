from pathlib import Path
import sys,tempfile,threading,unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import openpyxl
import sfc_http as http
import sfc_excel as excel

HEADER=['Factory Group Cd_UccId','Audit Date (dd-mmm-yyyy)','Internal AUDIT ID','Plant','Line']
def fixture(path,day,group='CL_VH'):
 book=openpyxl.Workbook();sheet=book.active;sheet.append(HEADER)
 sheet.append([group,day,day+'-id','=literal','001']);book.save(path);book.close()

class HttpExportTests(unittest.TestCase):
 def test_factory_group_exists_in_both_modes_so_check_internal_columns(self):
  with tempfile.TemporaryDirectory() as directory:
   path=Path(directory)/'report.xlsx';fixture(path,'2026-10-24')
   self.assertEqual(excel.validate_export(path,'2026-10-24','2026-10-24',True),1)
   with self.assertRaises(ValueError):excel.validate_export(path,nike=False)
   book=openpyxl.Workbook();sheet=book.active
   sheet.append(HEADER+['SO#','Country','Pass/Fail']);sheet.append(['CL_VH','2026-10-24','id','VH','line','SO1','VN','Pass'])
   book.save(path);book.close()
   self.assertEqual(excel.validate_export(path,'2026-10-24','2026-10-24',False),1)
   with self.assertRaises(ValueError):excel.validate_export(path,nike=True)
 def test_only_successful_inputs_no_other_submit_handlers(self):
  form=http.Form('<input name="__VIEWSTATE" value="state"><input type="submit" name="btnPrint" value="submit"><input type="submit" name="btnExcel"><input name="skip" disabled><input type="checkbox" name="unchecked"><input type="checkbox" name="checked" checked value="on"><input name="ctl00$rpnlMain$ContentPlaceHolder1$ckbNIKE" type="hidden" value="U">')
  self.assertEqual(set(form.fields),{'__VIEWSTATE','checked',http.PREFIX+'ckbNIKE'})
 def test_no_reuse_of_locked_browser_session(self):
  session=http.Session('VH',cookies=[{'name':'ASP.NET_SessionId','value':'busy-session'},{'name':'MyFactoryName','value':'VH'}])
  self.assertEqual([c.name for c in session.jar],['MyFactoryName'])
 def test_exact_dates_nike_button_and_validate_before_publish(self):
  with tempfile.TemporaryDirectory() as directory,patch.object(excel,'EXPORT_ROOT',Path(directory)):
   root=Path(directory);source=root/'source.xlsx';fixture(source,'2026-01-02')
   session=http.Session('VH',cookies=[]);session.form=http.Form('<input name="__VIEWSTATE" value="original-state">')
   calls=[]
   def request(url,fields):calls.append(fields);return source.read_bytes()
   session.request=request
   item=session.export_part('2026-01-02','2026-01-02',True,root/'verified')
   fields=calls[0]
   self.assertEqual(fields['rpnlMain_ContentPlaceHolder1_dtDateFrom_Raw'],'1767312000000')
   self.assertEqual(fields[http.PREFIX+'dtDateTo'],'2026-01-02')
   self.assertEqual(fields[http.PREFIX+'ckbNIKE'],'C')
   self.assertEqual(fields['__EVENTTARGET'],http.PREFIX+'btnExcel')
   self.assertEqual(item['rows'],1)
   session.request=lambda *args: self.fail('A verified cached part must not be requested again')
   cached=session.export_part('2026-01-02','2026-01-02',True,root/'verified')
   self.assertEqual(cached['sha256'],item['sha256'])
   session.request=request
   with self.assertRaises(ValueError):session.export_part('2026-01-03','2026-01-03',True,root/'wrong_date')
   self.assertFalse((root/'wrong_date.xlsx').exists())
   fixture(source,'2026-01-02','CL_JV1')
   with self.assertRaises(ValueError):session.export_part('2026-01-02','2026-01-02',True,root/'wrong_plant')
   self.assertFalse((root/'wrong_plant.xlsx').exists())
 def test_daily_coverage_reconciled_to_single_month_table(self):
  with tempfile.TemporaryDirectory() as directory,patch.object(excel,'EXPORT_ROOT',Path(directory)):
   root=Path(directory);session=http.Session('VH',cookies=[]);session.worker=http.Session('VH',cookies=[])
   active=set();lock=threading.Lock();calls=[]
   def export_part(worker,first,last,nike,path):
    with lock:
     self.assertNotIn(id(worker),active);active.add(id(worker));calls.append((id(worker),first,last))
    try:
     path=Path(str(path)+'.xlsx');path.parent.mkdir(parents=True,exist_ok=True);fixture(path,first)
     return dict(excel.file_info(path,1),plant='VH',nike=True,date_from=first,date_to=last)
    finally:
     with lock:active.remove(id(worker))
   with patch.object(http.Session,'export_part',export_part):
    updates=list(session.export_period('2026-01-01','2026-01-05',True,root/'month'))
   item=updates[-1]['file'];self.assertEqual(item['rows'],5)
   self.assertEqual(len({c[0] for c in calls}),2)
   self.assertEqual([i['date_from'] for i in item['parts']],['2026-01-0'+str(i) for i in range(1,6)])
   header,rows=excel.read_export(item['path']);self.assertEqual(len(rows),5)
   self.assertEqual(rows[0][3],'=literal')
   self.assertEqual(excel.validate_export(item['path'],'2026-01-01','2026-01-05',True),5)
 def test_empty_sfc_day_has_only_base_columns(self):
  with tempfile.TemporaryDirectory() as directory,patch.object(excel,'EXPORT_ROOT',Path(directory)):
   root=Path(directory);empty=root/'empty.xlsx';full=root/'full.xlsx'
   book=openpyxl.Workbook();book.active.append(HEADER[:-1]);book.save(empty);book.close()
   fixture(full,'2026-01-02')
   items=[dict(excel.file_info(p),plant='VH',nike=True,date_from='2026-01-01',date_to='2026-01-02') for p in (empty,full)]
   item=excel.combine_exports(items,root/'month.xlsx','VH',True,include_sources=False)
   h,rows=excel.read_export(item['path']);self.assertEqual(h,tuple(HEADER));self.assertEqual(len(rows),1)
   book=openpyxl.Workbook();book.active.append(['Unrelated']+HEADER[1:-1]);book.save(empty);book.close()
   with self.assertRaises(ValueError):excel.combine_exports(items,root/'invalid.xlsx','VH',True,include_sources=False)
   self.assertFalse((root/'invalid.xlsx').exists())

if __name__=='__main__':unittest.main()
