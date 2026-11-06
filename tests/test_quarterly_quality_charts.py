from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree

from report_tools.quarterly_quality.template_output import C, X, repair_excel_chart_references


class QuarterlyChartTests(unittest.TestCase):
    def test_ftt_region_chart_sources_use_weighted_totals(self):
        from openpyxl import Workbook
        from report_tools.quarterly_quality.editable_charts import Values
        values=Values.__new__(Values)
        values.book=Workbook()
        sheet=values.book.active
        sheet.title='FTT month-quarter'
        values.memo={}
        for row,amount in {2:100,3:900,9:90,10:810,17:5,18:90,
                           4:200,5:300,11:190,12:270,19:10,20:30}.items():
            sheet.cell(row,18,amount)
        self.assertAlmostEqual(values.cell(sheet.title,'R36'),.805)
        self.assertAlmostEqual(values.cell(sheet.title,'R37'),.84)
        self.assertTrue(sheet['R36'].value.startswith('='))
        self.assertTrue(sheet['R37'].value.startswith('='))

    def fixture(self, path):
        chart = f'''<chartSpace xmlns="{C[1:-1]}"><chart><plotArea><lineChart><ser>
        <cat><multiLvlStrRef><multiLvlStrCache>
        <lvl><pt idx="0"><v>Jun</v></pt><pt idx="17"><v>Nov</v></pt></lvl>
        <lvl><pt idx="0"><v>FY26</v></pt><pt idx="11"><v>FY27</v></pt></lvl>
        </multiLvlStrCache><f>'Month HFPA'!$H$2:$Y$4</f></multiLvlStrRef></cat>
        <val><numRef><f>'Month HFPA'!$H$23:$Y$23</f><numCache>
        <ptCount val="18"/><pt idx="17"><v>0.84</v></pt></numCache></numRef></val>
        </ser></lineChart></plotArea></chart></chartSpace>'''.encode()
        sheet = f'''<worksheet xmlns="{X[1:-1]}"><sheetData><row r="23">
        <c r="Y23" s="4"><f>IFERROR(1-Y14/Y5,"")</f><v/></c>
        </row></sheetData></worksheet>'''.encode()
        with ZipFile(path, 'w') as package:
            package.writestr('xl/charts/chart1.xml', chart)
            package.writestr('xl/worksheets/sheet1.xml', sheet)
            package.writestr('xl/drawings/drawing1.xml', b'unchanged anchor')

    def test_wps_category_schema_and_all_unrelated_parts_are_preserved(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'report.xlsx'
            self.fixture(path)
            with ZipFile(path) as package:
                before = {name: package.read(name) for name in package.namelist()}
            repair_excel_chart_references(path)
            with ZipFile(path) as package:
                self.assertEqual(set(package.namelist()), set(before))
                for name, content in before.items():
                    if name != 'xl/charts/chart1.xml':
                        self.assertEqual(package.read(name), content)
                root = etree.fromstring(package.read('xl/charts/chart1.xml'))
                reference = root.find('.//' + C + 'multiLvlStrRef')
                self.assertEqual([node.tag for node in reference], [C + 'f', C + 'multiLvlStrCache'])
                self.assertEqual(reference[0].text, "'Month HFPA'!$H$2:$Y$4")
                cache = reference[1]
                self.assertEqual(cache[0].tag, C + 'ptCount')
                self.assertEqual(cache[0].get('val'), '18')
                self.assertEqual(root.find('.//' + C + 'val').find('.//' + C + 'v').text, '0.84')
            repaired = path.read_bytes()
            repair_excel_chart_references(path)
            self.assertEqual(path.read_bytes(), repaired)

    def test_adding_charts_keeps_calculated_cell_cache_and_style(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / 'report.xlsx'
            self.fixture(path)
            cached = f'''<worksheet xmlns="{X[1:-1]}"><sheetData><row r="23">
            <c r="Y23"><f>IFERROR(1-Y14/Y5,"")</f><v>0.84</v></c>
            </row></sheetData></worksheet>'''.encode()
            repair_excel_chart_references(path, cached_sheets={'xl/worksheets/sheet1.xml': cached})
            with ZipFile(path) as package:
                cell = etree.fromstring(package.read('xl/worksheets/sheet1.xml')).find('.//' + X + 'c')
                self.assertEqual(cell.find(X + 'f').text, 'IFERROR(1-Y14/Y5,"")')
                self.assertEqual(cell.find(X + 'v').text, '0.84')
                self.assertEqual(cell.get('s'), '4')


if __name__ == '__main__':
    unittest.main()
