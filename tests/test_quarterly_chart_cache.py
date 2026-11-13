from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook

from report_tools.quarterly_quality.chart_cache import NS, refresh_qms_chart_caches


class QuarterlyChartCacheTests(unittest.TestCase):
    def fixture(self, path, title='FY27 Q2 QMS HFPA ANALYSIS'):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Monthly Data'
        sheet.append(['Sep', 'Oct', 'Nov'])
        sheet.append([0.8, 0.9, 0])
        data = BytesIO()
        workbook.save(data)
        with ZipFile(data) as book:
            parts = {n: book.read(n) for n in book.namelist()}
        # Preserve a numeric string beyond the 15 digits emitted by Office charts.
        parts['xl/worksheets/sheet1.xml'] = parts['xl/worksheets/sheet1.xml'].replace(
            b'<v>0.8</v>', b'<v>0.8473384621256318</v>')
        embedded = BytesIO()
        with ZipFile(embedded, 'w') as book:
            for name, content in parts.items():
                book.writestr(name, content)
        chart = f'''<c:chartSpace xmlns:c="{NS['c']}"><c:chart><c:plotArea><c:lineChart>
        <c:ser><c:tx><c:strRef><c:f>VN</c:f><c:strCache><c:ptCount val="1"/>
        <c:pt idx="0"><c:v>VN</c:v></c:pt></c:strCache></c:strRef></c:tx>
        <c:cat><c:strRef><c:f>'Monthly Data'!$A$1:$C$1</c:f><c:strCache>
        <c:ptCount val="3"/><c:pt idx="0"><c:v>Sep</c:v></c:pt>
        <c:pt idx="1"><c:v>Oct</c:v></c:pt><c:pt idx="2"><c:v>Nov</c:v></c:pt>
        </c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>'Monthly Data'!$A$2:$C$2</c:f>
        <c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="3"/>
        <c:pt idx="0"><c:v>0.847338462125632</c:v></c:pt>
        <c:pt idx="1"><c:v>0.9</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt>
        </c:numCache></c:numRef></c:val></c:ser></c:lineChart></c:plotArea></c:chart>
        <c:externalData xmlns:r="{NS['r']}" r:id="rId1"/></c:chartSpace>'''
        relations = 'http://schemas.openxmlformats.org/package/2006/relationships'
        with ZipFile(path, 'w') as package:
            package.writestr('ppt/slides/slide7.xml',
                             f'<slide xmlns:a="{NS["a"]}"><a:t>{title}</a:t></slide>')
            package.writestr('ppt/slides/_rels/slide7.xml.rels',
                             f'<Relationships xmlns="{relations}"><Relationship Id="rId7" '
                             f'Type="{NS["r"]}/chart" Target="../charts/chart7.xml"/></Relationships>')
            package.writestr('ppt/charts/chart7.xml', chart)
            package.writestr('ppt/charts/_rels/chart7.xml.rels',
                             f'<Relationships xmlns="{relations}"><Relationship Id="rId1" '
                             f'Type="{NS["r"]}/package" Target="../embeddings/Workbook7.xlsx"/></Relationships>')
            package.writestr('ppt/embeddings/Workbook7.xlsx', embedded.getvalue())
            package.writestr('ppt/slides/slide1.xml', b'<slide>Unchanged consumer summary</slide>')

    def test_refresh_copies_exact_precision_and_preserves_zero_formulas_and_other_parts(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'report.pptx'
            self.fixture(path)
            with ZipFile(path) as package:
                before = {n: package.read(n) for n in package.namelist()}
            self.assertEqual(refresh_qms_chart_caches(path), 1)
            with ZipFile(path) as package:
                for name, content in before.items():
                    if name != 'ppt/charts/chart7.xml':
                        self.assertEqual(package.read(name), content, name)
                chart = etree.fromstring(package.read('ppt/charts/chart7.xml'))
                self.assertEqual(chart.xpath('.//c:numCache/c:pt/c:v/text()', namespaces=NS),
                                 ['0.8473384621256318', '0.9', '0'])
                self.assertEqual(chart.find('.//c:val/c:numRef/c:f', NS).text, "'Monthly Data'!$A$2:$C$2")
                self.assertEqual(chart.find('.//c:formatCode', NS).text, '0.0%')
                self.assertIsNone(chart.find('.//c:tx/c:strRef', NS))
                self.assertEqual(chart.find('.//c:tx/c:v', NS).text, 'VN')
            first = path.read_bytes()
            self.assertEqual(refresh_qms_chart_caches(path), 0)
            self.assertEqual(path.read_bytes(), first)

    def test_unrelated_slide_is_untouched(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'report.pptx'
            self.fixture(path, title='INBOUND/MARKET/PI CLAIM CASE TRACKING')
            before = path.read_bytes()
            self.assertEqual(refresh_qms_chart_caches(path), 0)
            self.assertEqual(path.read_bytes(), before)

    def test_wps_linked_chart_embeds_the_saved_source_without_following_its_url(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'report.pptx'
            self.fixture(path)
            with ZipFile(path) as package:
                before = {n: package.read(n) for n in package.namelist()}
            original_workbook = before.pop('ppt/embeddings/Workbook7.xlsx')
            workbook_path = Path(folder) / 'source.xlsx'
            workbook_path.write_bytes(original_workbook)
            rel_name = 'ppt/charts/_rels/chart7.xml.rels'
            root = etree.fromstring(before[rel_name])
            root[0].set('Target', 'https://example.invalid/untrusted.xlsx')
            root[0].set('TargetMode', 'External')
            before[rel_name] = etree.tostring(root)
            before['[Content_Types].xml'] = b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
            before['ppt/charts/chart7.xml'] = before['ppt/charts/chart7.xml'].replace(
                b"'Monthly Data'!", b"'[source.xlsx]Monthly Data'!")
            with ZipFile(path, 'w') as package:
                for name, data in before.items():
                    package.writestr(name, data)
            with self.assertRaisesRegex(ValueError, 'known saved source workbook'):
                refresh_qms_chart_caches(path)
            self.assertEqual(refresh_qms_chart_caches(path, linked_workbook=workbook_path), 1)
            with ZipFile(path) as package:
                relation = etree.fromstring(package.read(rel_name))[0]
                self.assertIsNone(relation.get('TargetMode'))
                self.assertTrue(relation.get('Type').endswith('/package'))
                self.assertEqual(relation.get('Target'), '../embeddings/quarterly_source.xlsx')
                self.assertEqual(package.read('ppt/embeddings/quarterly_source.xlsx'), original_workbook)
                types = etree.fromstring(package.read('[Content_Types].xml'))
                self.assertEqual(types[0].get('Extension'), 'xlsx')
            self.assertEqual(refresh_qms_chart_caches(path), 0)

    def test_wps_ranking_arrays_are_linked_to_the_original_labels(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / 'report.pptx'
            self.fixture(path)
            with ZipFile(path) as package:
                parts = {n: package.read(n) for n in package.namelist()}
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = 'Quarter HFPA '
            labels = ['NIKE INITIATOR', 'MODEL B', 'MODEL C']
            for row, label, value in zip((4, 5, 6), labels, (0.1, 0.2, 0.3)):
                sheet.cell(row, 14, label)
                sheet.cell(row, 17, value)
            source = Path(folder) / 'source.xlsx'
            workbook.save(source)
            rel = etree.fromstring(parts['ppt/charts/_rels/chart7.xml.rels'])
            rel[0].set('TargetMode', 'External')
            rel[0].set('Target', 'source.xlsx')
            parts['ppt/charts/_rels/chart7.xml.rels'] = etree.tostring(rel)
            parts['[Content_Types].xml'] = b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
            chart = etree.fromstring(parts['ppt/charts/chart7.xml'])
            chart.find('.//c:cat/c:strRef/c:f', NS).text = '{"NIKE\nINITIATOR","MODEL B","MODEL C"}'
            for point, label in zip(chart.findall('.//c:cat//c:pt/c:v', NS), ['NIKE\nINITIATOR', *labels[1:]]):
                point.text = label
            chart.find('.//c:val/c:numRef/c:f', NS).text = "'[source.xlsx]Quarter HFPA '!$Q$4:$Q$6"
            parts['ppt/charts/chart7.xml'] = etree.tostring(chart)
            with ZipFile(path, 'w') as package:
                for name, data in parts.items():
                    package.writestr(name, data)
            self.assertEqual(refresh_qms_chart_caches(path, linked_workbook=source), 1)
            with ZipFile(path) as package:
                chart = etree.fromstring(package.read('ppt/charts/chart7.xml'))
                self.assertEqual(chart.find('.//c:cat/c:strRef/c:f', NS).text, "'QMS Chart Labels'!$A$2:$A$4")
                self.assertEqual(chart.find('.//c:val/c:numRef/c:f', NS).text, "'Quarter HFPA '!$Q$4:$Q$6")
                self.assertEqual(chart.xpath('.//c:cat//c:pt/c:v/text()', namespaces=NS), ['NIKE\nINITIATOR', *labels[1:]])
                from openpyxl import load_workbook
                embedded = load_workbook(BytesIO(package.read('ppt/embeddings/quarterly_source.xlsx')), data_only=False)
                self.assertEqual(embedded['Quarter HFPA ']['N4'].value, 'NIKE INITIATOR')
                self.assertEqual(embedded['QMS Chart Labels']['A1'].value, "'Quarter HFPA '!$N$4:$N$6")
                self.assertEqual(embedded['QMS Chart Labels']['A2'].value, 'NIKE\nINITIATOR')
                embedded.close()
            self.assertEqual(refresh_qms_chart_caches(path), 0)


if __name__ == '__main__':
    unittest.main()
