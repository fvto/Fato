"""Exercise the dashboard's mixed-source filtering and size grand total."""
from pathlib import Path
import io
import os
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import openpyxl
import spreadsheet_combine as combine
import sfc_dashboard_server as server
from werkzeug.serving import make_server


@unittest.skipUnless(os.environ.get('CHINGLUH_UI_TESTS')=='1','Set CHINGLUH_UI_TESTS=1 for Chrome')
class SearchBrowserTests(unittest.TestCase):
    def test_filter_each_column_and_sum_size_without_xlsx_export(self):
        from playwright.sync_api import sync_playwright,expect
        with tempfile.TemporaryDirectory() as temporary, patch.object(combine,'ROOT',Path(temporary)/'jobs'),patch.object(combine,'XLSX_ROW_THRESHOLD',2):
            http=make_server('127.0.0.1',0,server.app,threaded=True)
            thread=threading.Thread(target=http.serve_forever,daemon=True);thread.start()
            try:
                with sync_playwright() as p:
                    browser=p.chromium.launch(channel='chrome',headless=True)
                    page=browser.new_page(viewport=dict(width=1400,height=900));errors=[]
                    page.on('pageerror',lambda error:errors.append(str(error)))
                    page.goto(f'http://127.0.0.1:{http.server_port}/combine')
                    book=openpyxl.Workbook()
                    for row in [['FTY','Number of Pairs','Size','Mold Code','Unit Type','Issue'],['VH',.25,10,'M001','DRR','Tear'],['JV',2,8,'M002','DRR','Other']]:book.active.append(row)
                    excel=io.BytesIO();book.save(excel);book.close()
                    csv=b'Mold Code,Size,Number of Pairs,FTY,Unit Type,Issue\nM001,9.0,1.5,VH,DRR,Bubble\nM001,9,2,VH,NON-DRR,Ignore\nM001,10,invalid,JV,DRR,Bad\nM002,9,3,VH,DRR,Other\n'
                    page.locator('#combine-upload').set_input_files([dict(name='source.csv',mimeType='text/csv',buffer=csv),dict(name='source.xlsx',mimeType='application/octet-stream',buffer=excel.getvalue())])
                    expect(page.locator('#combine-file-count')).to_have_text('2 file')
                    page.locator('#combine-preview').click();expect(page.locator('#combine-result')).to_be_visible()
                    expect(page.locator('#combine-run')).to_have_text('Tạo gói Power Query →')
                    page.locator('#combine-search-create').click();expect(page.locator('#combine-search-ready')).to_be_visible(timeout=15000)
                    expect(page.locator('#combine-search-stats')).to_contain_text('6 dòng')
                    expect(page.locator('#combine-search-warning')).to_contain_text('1 dòng')
                    page.locator('.combine-filter input').first.fill('M001')
                    page.locator('#combine-search-run').click()
                    expect(page.locator('#combine-search-stats')).to_contain_text('4 dòng',timeout=15000)
                    expect(page.locator('#combine-size-totals')).to_contain_text('3.5')
                    expect(page.locator('#combine-search-table tbody tr')).to_have_count(4)
                    expect(page.locator('#combine-top-issues')).to_contain_text('Bubble')
                    expect(page.locator('#combine-top-issues')).not_to_contain_text('Ignore')
                    page.locator('#combine-filter-add').click()
                    expect(page.locator('#combine-size-totals')).to_be_hidden()
                    page.locator('.combine-filter').nth(1).locator('select').first.select_option(label='FTY')
                    page.locator('.combine-filter').nth(1).locator('input').fill('VH')
                    page.locator('#combine-search-run').click()
                    expect(page.locator('#combine-search-stats')).to_contain_text('3 dòng',timeout=15000)
                    expect(page.locator('#combine-search-warning')).to_be_hidden()
                    page.locator('#combine-search-clear').click()
                    expect(page.locator('#combine-search-stats')).to_contain_text('6 dòng')
                    page.locator('#combine-run').click()
                    expect(page.locator('#combine-download')).to_be_visible(timeout=15000)
                    expect(page.locator('#combine-download')).to_have_attribute('download',__import__('re').compile(r'\.zip$'))
                    page.set_viewport_size(dict(width=390,height=844))
                    self.assertTrue(page.evaluate('document.documentElement.scrollWidth <= innerWidth'))
                    self.assertEqual(errors,[]);browser.close()
            finally:http.shutdown();thread.join(timeout=5)


if __name__=='__main__':unittest.main()
