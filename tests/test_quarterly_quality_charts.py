from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree

from report_tools.quarterly_quality.template_output import C, X, repair_excel_chart_references


class QuarterlyChartTests(unittest.TestCase):
    def test_native_trend_legend_keeps_data_and_cell_names_and_is_idempotent(self):
        from report_tools.quarterly_quality.editable_charts import compatible_trend_legend, NS
        source = f'''<c:chartSpace xmlns:c="{C[1:-1]}"><c:chart><c:plotArea>
        <c:layout><c:manualLayout><c:y val=".17"/><c:h val=".7"/></c:manualLayout></c:layout>
        <c:lineChart><c:ser><c:tx><c:strRef><c:f>"VH"</c:f></c:strRef></c:tx>
        <c:val><c:numRef><c:f>'Data'!$B$2:$D$2</c:f><c:numCache><c:pt idx="0"><c:v>.8</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
        <c:ser><c:tx><c:strRef><c:f>'Data'!$A$3</c:f><c:strCache><c:pt idx="0"><c:v>VH2</c:v></c:pt></c:strCache></c:strRef></c:tx></c:ser>
        </c:lineChart></c:plotArea><c:legend><c:legendPos val="t"/><c:layout><c:manualLayout>
        <c:x val=".74"/><c:h val=".06"/></c:manualLayout></c:layout><c:overlay val="1"/>
        <c:spPr/></c:legend></c:chart></c:chartSpace>'''.encode()
        result = compatible_trend_legend(source)
        root = etree.fromstring(result)
        self.assertEqual(root.find('.//c:ser/c:tx/c:v', NS).text, 'VH')
        self.assertEqual(root.findall('.//c:ser/c:tx/c:strRef/c:f', NS)[0].text, "'Data'!$A$3")
        self.assertEqual(root.find('.//c:val/c:numRef/c:f', NS).text, "'Data'!$B$2:$D$2")
        self.assertEqual(root.find('.//c:val//c:pt/c:v', NS).text, '.8')
        legend = root.find('.//c:legend', NS)
        self.assertIsNone(legend.find('c:layout/c:manualLayout', NS))
        self.assertEqual(legend.find('c:overlay', NS).get('val'), '0')
        self.assertEqual([etree.QName(c).localname for c in legend], ['legendPos', 'layout', 'overlay', 'spPr', 'txPr'])
        self.assertEqual(compatible_trend_legend(result), result)

    def test_legend_repair_does_not_change_bar_charts(self):
        from report_tools.quarterly_quality.editable_charts import compatible_trend_legend
        source = f'<chartSpace xmlns="{C[1:-1]}"><chart><plotArea><barChart/></plotArea></chart></chartSpace>'.encode()
        self.assertEqual(compatible_trend_legend(source), source)

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
