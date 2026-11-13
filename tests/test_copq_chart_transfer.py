"""FTT transfer must preserve source data and reject stale/image substitutes."""
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from report_tools.copq.Source import copq_office_charts as transfer


class ChartTransferTests(unittest.TestCase):
    def test_source_chart_is_identified_by_data_not_object_order(self):
        self.assertEqual(transfer._chart_header(
            "=SERIES('Top Defects Summary'!$B$19,'Top Defects Summary'!$A$20:$A$22,'Top Defects Summary'!$B$20:$B$22,1)"),
            ('Top Defects Summary', 19))
        self.assertEqual(transfer._chart_header(
            "=SERIES('[FTT.xlsx]Top Defects Summary'!$B$13,Sheet1!$A$1:$A$3,Sheet1!$B$1:$B$3,1)"),
            ('Top Defects Summary', 13))
        self.assertIsNone(transfer._chart_header('unrecognized'))

    def test_comparison_catches_missing_fourth_defect_or_model(self):
        source = tuple((f'defect {i}', ('model 1', 'model 2'), (.2, .3)) for i in range(4))
        self.assertTrue(transfer._same_data(source, source))
        self.assertFalse(transfer._same_data(source, source[:3]))
        self.assertFalse(transfer._same_data(source, tuple((n, cats[:1], vals[:1]) for n, cats, vals in source)))
        self.assertFalse(transfer._same_data(source, source[:3] + (('defect 3', ('model 1', 'model 2'), (.2, .4)),)))

    def test_sparse_excel_points_do_not_create_ghost_models(self):
        xml = b'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart><c:plotArea><c:barChart><c:ser>
        <c:tx><c:v>fourth defect</c:v></c:tx>
        <c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strCache></c:strRef></c:cat>
        <c:val><c:numRef><c:numCache><c:ptCount val="2"/><c:pt idx="1"><c:v>0.333333333333</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>'''
        self.assertEqual(transfer._saved_signature(xml), (('fourth defect', ('A', 'B'), (0., .333333333333)),))

    def package(self, path, *, image=False):
        ns = transfer.NS
        slide = etree.Element('{'+ns['p']+'}sld', nsmap=ns)
        tree = etree.SubElement(slide, '{'+ns['p']+'}spTree')
        relns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        rels = etree.Element('{'+relns+'}Relationships')
        with ZipFile(path, 'w') as z:
            for i, (name, _) in enumerate(transfer.SLOTS.values(), 1):
                frame = etree.SubElement(tree, '{'+ns['p']+'}pic' if image and i == 1 else '{'+ns['p']+'}graphicFrame')
                nv = etree.SubElement(frame, '{'+ns['p']+'}nvGraphicFramePr')
                etree.SubElement(nv, '{'+ns['p']+'}cNvPr', name=name)
                etree.SubElement(frame, '{'+ns['c']+'}chart', {'{'+ns['r']+'}id': f'rId{i}'})
                etree.SubElement(rels, '{'+relns+'}Relationship', Id=f'rId{i}', Target=f'../charts/chart{i}.xml')
                chartrels = etree.Element('{'+relns+'}Relationships')
                etree.SubElement(chartrels, '{'+relns+'}Relationship', Id='rId1',
                                 Type=ns['r']+'/oleObject', Target='file:///private/staging/source.xlsx', TargetMode='External')
                z.writestr(f'ppt/charts/_rels/chart{i}.xml.rels', etree.tostring(chartrels))
                z.writestr(f'ppt/charts/chart{i}.xml', b'original native chart; must remain unchanged')
            z.writestr('ppt/slides/slide2.xml', etree.tostring(slide))
            z.writestr('ppt/slides/_rels/slide2.xml.rels', etree.tostring(rels))
            z.writestr('ppt/charts/unrelated.xml', b'unrelated COPQ chart')

    def test_edit_data_points_to_delivered_workbook_without_rebuilding_charts(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'report.pptx'
            workbook = Path(directory)/'FTT_Combined_Report.xlsx'
            self.package(path)
            self.assertEqual(len(transfer._workbook_links(path, workbook, True)), 4)
            with ZipFile(path) as z:
                self.assertEqual(z.read('ppt/charts/chart1.xml'), b'original native chart; must remain unchanged')
                self.assertEqual(z.read('ppt/charts/unrelated.xml'), b'unrelated COPQ chart')
                for i in range(1, 5):
                    rel = etree.fromstring(z.read(f'ppt/charts/_rels/chart{i}.xml.rels'))[0]
                    self.assertEqual(rel.get('Target'), workbook.resolve().as_uri())
            with self.assertRaises(transfer.ChartTransferError):
                transfer._workbook_links(path, workbook, False)

    def test_picture_substitute_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'report.pptx'
            self.package(path, image=True)
            with ZipFile(path) as z, self.assertRaises(transfer.ChartTransferError):
                transfer._chart_parts(z)


if __name__ == '__main__':
    unittest.main()
