from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from lxml import etree
from openpyxl import Workbook, load_workbook
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

from report_tools.quarterly_quality import template_output
from report_tools.quarterly_quality.quarterly_quality import Counts, PLANTS, Quarter, quarter_months


class QuarterlyQualityPresentationTests(unittest.TestCase):
    def setUp(self):
        months = quarter_months("FY27-Q2")
        values = {("HFPA", months[0], plant): Counts(100, 5) for plant in PLANTS}
        current = (Counts(100, 10), Counts(900, 180), Counts(200, 30), Counts(300, 0))
        values.update({("HFPA", months[3], plant): record for plant, record in zip(PLANTS, current)})
        self.quarter = Quarter("FY27-Q2", months[:3], months[3:], months[3:4], values, {}, {},
                               date_from="2026-09-01", date_to="2026-09-30",
                               prior_date_from="2026-06-01", prior_date_to="2026-06-30")

    def consumer_slide(self, deck):
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        slide.shapes.add_textbox(0, 0, Inches(5), Inches(1)).text = "CONSUMER PROMISE REVIEW"
        table = slide.shapes.add_table(13, 14, 0, Inches(1), Inches(12), Inches(6)).table
        table.cell(6, 8).text = "Quality (HFPA)"
        table.cell(7, 8).text = "Q1FY27"
        table.cell(7, 9).text = "vs PQ"
        table.cell(8, 0).text = "CLG"
        for row, label in ((8, "VH"), (9, "JV"), (10, "CLG")):
            table.cell(row, 2).text = label
            table.cell(row, 9).text = "stale"
            table.cell(row, 9).text_frame.paragraphs[0].runs[0].font.size = Pt(12)
        return slide, table

    def qms_slide(self, deck, metric, image):
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        spec = template_output.SHAPES[metric]
        title = slide.shapes.add_textbox(0, 0, Inches(5), Inches(1))
        title.text = f"FY27 Q1 QMS {metric} ANALYSIS"
        shapes = [(title, spec["title"])]
        for shape_id in (spec["period"], *spec["VN"], *spec["Indo"]):
            picture = slide.shapes.add_picture(BytesIO(image), 0, 0, width=Inches(3))
            shapes.append((picture, shape_id))
        for shape, shape_id in shapes:
            shape._element.xpath(".//p:cNvPr")[0].set("id", str(shape_id))
        return slide

    def make_deck(self, full):
        image = BytesIO()
        Image.new("RGB", (1300, 500), "white").save(image, "PNG")
        deck = Presentation()
        if full:
            self.consumer_slide(deck)
            for _ in range(5):
                slide = deck.slides.add_slide(deck.slide_layouts[6])
                slide.shapes.add_picture(BytesIO(image.getvalue()), 0, 0)
        else:
            deck.slides.add_slide(deck.slide_layouts[6])
        self.qms_slide(deck, "FTT", image.getvalue())
        self.qms_slide(deck, "HFPA", image.getvalue())
        return deck

    def test_consumer_rates_weight_each_factory_by_its_inspection_quantity(self):
        slide, table = self.consumer_slide(Presentation())
        template_output.update_consumer_promise(slide, self.quarter)
        self.assertEqual([table.cell(row, 8).text for row in (8, 9, 10)], ["81.0%", "94.0%", "85.3%"])
        self.assertEqual([table.cell(row, 9).text for row in (8, 9, 10)], ["95.0%"] * 3)
        self.assertEqual(table.cell(7, 8).text, "Q2FY27")
        self.assertEqual(table.cell(8, 8).text_frame.paragraphs[0].runs[0].font.size, Pt(12))

    def test_missing_prior_and_zero_inspections_have_no_numeric_rate(self):
        self.quarter.values = {("HFPA", "2026-09", plant): Counts(0, 0) for plant in PLANTS}
        slide, table = self.consumer_slide(Presentation())
        template_output.update_consumer_promise(slide, self.quarter)
        self.assertEqual([table.cell(row, column).text for row in (8, 9, 10) for column in (8, 9)], ["—"] * 6)

    def test_current_quarter_export_keeps_the_saved_prior_rates(self):
        self.quarter.prior_date_from = self.quarter.prior_date_to = None
        self.quarter.values = {key: value for key, value in self.quarter.values.items() if key[1] == "2026-09"}
        slide, table = self.consumer_slide(Presentation())
        prior = ["84.9%", "91.9%", "87.1%"]
        for row, value in zip((8, 9, 10), prior):
            table.cell(row, 9).text = value
        template_output.update_consumer_promise(slide, self.quarter)
        self.assertEqual([table.cell(row, 8).text for row in (8, 9, 10)], ["81.0%", "94.0%", "85.3%"])
        self.assertEqual([table.cell(row, 9).text for row in (8, 9, 10)], prior)

    def test_export_handles_full_and_compact_decks_without_changing_shared_source_images(self):
        for full in (False, True):
            with self.subTest(full=full), TemporaryDirectory() as folder:
                source, output = (Path(folder) / name for name in ("template.pptx", "result.pptx"))
                self.make_deck(full).save(source)
                template_output.write_powerpoint(source, output, self.quarter)
                deck = Presentation(output)
                self.assertEqual(len(deck.slides), 8 if full else 3)
                self.assertIn("FY27 Q2 QMS FTT ANALYSIS", [s.text for s in deck.slides[-2].shapes if s.has_text_frame])
                if full:
                    table = next(shape.table for shape in deck.slides[0].shapes if shape.has_table)
                    self.assertEqual(table.cell(10, 8).text, "85.3%")
                with ZipFile(source) as original, ZipFile(output) as result:
                    for name in original.namelist():
                        if name.startswith("ppt/media/") or full and name == "ppt/slides/slide2.xml":
                            self.assertEqual(original.read(name), result.read(name), name)
                    self.assertGreater(len(result.namelist()), len(original.namelist()))

    def test_missing_qms_slide_is_rejected(self):
        deck = Presentation()
        self.consumer_slide(deck)
        with self.assertRaisesRegex(ValueError, "QMS FTT"):
            template_output.quality_slides(deck)

    def test_local_templates_take_priority_and_configured_paths_are_honored(self):
        with TemporaryDirectory() as folder, TemporaryDirectory() as home:
            root = Path(folder)
            local, downloads = root / "Report/Quarterly", Path(home) / "Downloads"
            local.mkdir(parents=True)
            downloads.mkdir()
            names = ("Quarterly Report HFPA-FTT.xlsx", "Partner QBR Format Q2 FY27-CLG.pptx")
            for name in names:
                (local / name).touch()
                (downloads / name).touch()
            with patch.object(template_output, "APP_ROOT", root), patch.object(Path, "home", return_value=Path(home)):
                self.assertEqual(template_output.template_paths(), tuple(local / name for name in names))
                (local / names[1]).unlink()
                self.assertEqual(template_output.template_paths()[1], downloads / names[1])
                configured = root / "custom.pptx"
                configured.touch()
                (root / "report_tools_config.json").write_text(json.dumps({"quarterly_templates": {"powerpoint": "custom.pptx"}}))
                self.assertEqual(template_output.template_paths()[1], configured)

    def test_powerpoint_charts_have_hyperlinks_to_excel(self):
        image = BytesIO()
        Image.new("RGB", (1300, 500), "white").save(image, "PNG")
        deck = Presentation()
        deck.slides.add_slide(deck.slide_layouts[6])
        self.qms_slide(deck, "FTT", image.getvalue())
        self.qms_slide(deck, "HFPA", image.getvalue())

        with TemporaryDirectory() as folder:
            source = Path(folder) / "template.pptx"
            output = Path(folder) / "result.pptx"
            deck.save(source)
            template_output.write_powerpoint(source, output, self.quarter)

            with ZipFile(output) as package:
                rels_xml = package.read("ppt/slides/_rels/slide2.xml.rels")
                rels = etree.fromstring(rels_xml)
                hlink_targets = [rel.get("Target") for rel in rels if "hyperlink" in rel.get("Type", "")]
                self.assertIn("result.xlsx", hlink_targets)

    def test_excel_charts_style_b(self):
        wb = Workbook()
        ws_ftt = wb.active
        ws_ftt.title = "FTT month-quarter"
        ws_ftt.cell(1, 3, "Sep")
        ws_ftt.cell(41, 4, "TopModelVN")
        ws_ftt.cell(41, 7, 0.05)
        ws_ftt.cell(38, 4, "TopModelIndo")
        ws_ftt.cell(38, 7, 0.04)

        ws_hfpa_q = wb.create_sheet("Quarter HFPA ")
        ws_hfpa_q.cell(7, 14, "TopHFPAVN")
        ws_hfpa_q.cell(7, 17, 0.02)
        ws_hfpa_q.cell(4, 14, "TopHFPAIndo")
        ws_hfpa_q.cell(4, 17, 0.03)

        ws_hfpa_m = wb.create_sheet("Month HFPA")
        ws_hfpa_m.cell(4, 3, "Sep")

        with TemporaryDirectory() as folder:
            path = Path(folder) / "output.xlsx"
            wb.save(path)
            mappings = {"FTT": {"2026-09": 3}, "HFPA": {"2026-09": 3}}
            template_output.add_excel_charts(path, self.quarter, mappings)

            reloaded = load_workbook(path)
            self.assertGreaterEqual(len(reloaded["FTT month-quarter"]._charts), 2)
            self.assertGreaterEqual(len(reloaded["Quarter HFPA "]._charts), 2)
            self.assertEqual(reloaded["FTT month-quarter"]._charts[0].y_axis.numFmt.formatCode, "0.0%")
            self.assertIn('C9+C10-C17-C18', reloaded['FTT month-quarter']['C36'].value)
            self.assertIn('C11+C12-C19-C20', reloaded['FTT month-quarter']['C37'].value)
            reloaded.close()

    def test_powerpoint_export_automatically_converts_charts_using_peer_excel(self):
        from report_tools.quarterly_quality import editable_charts

        with TemporaryDirectory() as folder:
            root=Path(folder)
            source,output=root/'template.pptx',root/'report.pptx'
            self.make_deck(full=True).save(source)
            excel=output.with_suffix('.xlsx')
            excel.write_bytes(b'peer workbook fixture; converter is mocked')
            with patch.object(editable_charts,'convert') as convert, \
                    patch.object(template_output, 'trend_image') as trend_image, \
                    patch.object(template_output, 'rank_image') as rank_image:
                template_output.write_powerpoint(source,output,self.quarter)
                convert.assert_called_once()
                trend_image.assert_not_called()
                rank_image.assert_not_called()
                deck,workbook,target,styled_workbook=convert.call_args.args
                self.assertEqual((workbook,target,styled_workbook),(excel,output,excel))
                with ZipFile(deck) as package:
                    self.assertIn('ppt/slides/slide8.xml',package.namelist())

    def test_excel_charts_own_complete_design_and_small_readable_labels(self):
        from report_tools.quarterly_quality.editable_charts import (
            NS, CHART_SPECS, excel_chart_sources, prepare_excel_charts)

        for empty_rankings in (False, True):
            with self.subTest(empty_rankings=empty_rankings), TemporaryDirectory() as folder:
                root = Path(folder)
                book = Workbook()
                ftt = book.active
                ftt.title = 'FTT month-quarter'
                hfpa = book.create_sheet('Month HFPA')
                ranks = book.create_sheet('Quarter HFPA ')
                for col, month in enumerate(('Sep', 'Oct', 'Nov'), 3):
                    ftt.cell(1, col, month)
                    hfpa.cell(4, col, month)
                    for row in range(2, 21):
                        ftt.cell(row, col, 100 if row < 17 else 1)
                    for row in range(23, 30):
                        hfpa.cell(row, col, .9012345)
                    for row in range(31, 35):
                        ftt.cell(row, col, .7812345)
                long_name = 'WMNS NIKE EBERNON LOW PREMIUM MODEL NAME'
                if not empty_rankings:
                    for sheet, starts, column in ((ftt, (38, 41, 48, 51), 4),
                                                  (ranks, (4, 7, 14, 17), 14)):
                        for start in starts:
                            for i in range(3):
                                defect_names = (('Cleanness', 'Midsole/Outsole to upper bond gap', 'Over cement')
                                                if start % 2 == 0 else ('Thread end', 'Over cement', 'Cleanness'))
                                is_model = start in (38, 41, 4, 7)
                                sheet.cell(start+i, column, long_name if is_model else defect_names[i])
                                sheet.cell(start+i, column+3, .04123456789)
                excel = root/'report.xlsx'
                book.save(excel)
                book.close()
                columns = {metric: dict(zip(('2026-09', '2026-10', '2026-11'), (3, 4, 5)))
                           for metric in ('FTT', 'HFPA')}
                template_output.add_excel_charts(excel, self.quarter, columns)
                prepare_excel_charts(excel)
                calculated = load_workbook(excel, data_only=False)
                self.assertEqual(calculated['FTT month-quarter']['C36'].number_format, '0.0%')
                calculated.close()
                with ZipFile(excel) as workbook:
                    sources = excel_chart_sources(workbook)
                    self.assertEqual(len(sources), 12)
                    for metric, region, kind, *_ in CHART_SPECS:
                        source = etree.fromstring(workbook.read(sources[metric, region, kind]))
                        formats = source.findall('.//c:dLbls/c:numFmt', NS)
                        self.assertTrue(formats)
                        self.assertTrue(all(fmt.get('formatCode') == '0.0%' and fmt.get('sourceLinked') == '0'
                                            for fmt in formats))
                        sizes = source.xpath('.//a:rPr/@sz | .//a:defRPr/@sz', namespaces=NS)
                        self.assertLessEqual(max(map(int, sizes)), 1050)
                        if kind != 'trend':
                            self.assertEqual(source.find('c:chart/c:title/c:tx/c:rich/a:p/a:pPr', NS).get('algn'), 'ctr')
                            self.assertIsNone(source.find('c:chart/c:title/c:layout/c:manualLayout', NS))
                            self.assertEqual(source.find('c:chart/c:title/c:overlay', NS).get('val'), '0')
                            plot_top = float(source.find('.//c:plotArea/c:layout/c:manualLayout/c:y', NS).get('val'))
                            self.assertGreaterEqual(plot_top, .36)
                        if kind != 'trend' and not empty_rankings:
                            self.assertIsNotNone(source.find('.//c:ser/c:cat/c:strLit', NS))
                            self.assertIsNone(source.find('.//c:ser/c:cat/c:strRef', NS))
                            labels = source.xpath('.//c:ser/c:cat//c:pt/c:v/text()', namespaces=NS)
                            fills = {int(point.find('c:idx', NS).get('val')):
                                     point.find('c:spPr/a:solidFill/a:srgbClr', NS).get('val')
                                     for point in source.findall('.//c:ser/c:dPt', NS)}
                            if kind == 'model':
                                self.assertEqual(labels[0].replace('\n', ' '), long_name)
                                self.assertIn('\n', labels[0])
                                expected = template_output.RANK_FORMAT[metric, region, False][-1].lstrip('#')
                                self.assertEqual(list(fills.values()), [expected] * 3)
                            else:
                                from report_tools.hfpa.agents.color_manager import ColorManager
                                _, expected, _, _ = ColorManager(str(Path(__file__).resolve().parents[1] /
                                    'report_tools/reinspection/Color_Template.xlsx')).load_existing_template()
                                for index, label in enumerate(labels):
                                    self.assertEqual(fills[index], expected[ColorManager.normalize_name(label)])
                            self.assertIn('0.04123456789', source.xpath('.//c:val//c:pt/c:v/text()', namespaces=NS))
                            self.assertEqual(source.xpath('.//c:dLbl/c:tx/c:rich//a:t/text()', namespaces=NS), ['4.1%'] * 3)
                            self.assertEqual(source.xpath('.//c:catAx/c:tickLblSkip/@val', namespaces=NS), ['1'])
                        if kind == 'trend':
                            line_colors = source.xpath('.//c:ser/c:spPr/a:ln/a:solidFill/a:srgbClr/@val', namespaces=NS)
                            self.assertEqual(line_colors, [colour.lstrip('#') for colour in template_output.TREND_FORMAT[metric, region][-2]])
                            self.assertEqual(source.xpath('.//c:ser/c:marker/c:symbol/@val', namespaces=NS), ['none'] * 3)
                            legend = source.find('.//c:legend', NS)
                            self.assertEqual(legend.find('c:overlay', NS).get('val'), '0')
                            self.assertIsNone(legend.find('c:layout/c:manualLayout', NS))
                            self.assertTrue(all(fmt.get('formatCode') == '0.0%' for fmt in
                                                source.findall('.//c:dLbl/c:numFmt', NS)))
                            primary = source.find('.//c:ser', NS)
                            colour = template_output.TREND_FORMAT[metric, region][-1].lstrip('#')
                            self.assertEqual(primary.xpath('c:dLbls/c:dLbl/c:spPr/a:solidFill/a:srgbClr/@val', namespaces=NS), [colour] * 3)
                            self.assertEqual(primary.xpath('c:dLbls/c:dLbl/c:tx/c:rich//a:highlight/a:srgbClr/@val', namespaces=NS), [colour] * 3)
                            if metric == 'HFPA':
                                self.assertEqual(primary.xpath('c:dLbls/c:dLbl/c:dLblPos/@val', namespaces=NS), ['ctr'] * 3)

    def test_consumer_summary_table_image_generates_crisp_excel_style_screenshot(self):
        png_bytes = template_output.consumer_summary_table_image(self.quarter)
        im = Image.open(BytesIO(png_bytes))
        self.assertEqual(im.size, (547, 486))

    def test_consumer_corner_sticker_updates_date_box_and_table_image(self):
        template = "Report/Quarterly/Partner QBR Format Q2 FY27-CLG.pptx"
        if not Path(template).is_file():
            return
        with TemporaryDirectory() as folder:
            output = Path(folder) / "result.pptx"
            template_output.write_powerpoint(template, output, self.quarter)
            prs = Presentation(output)
            slide = prs.slides[0]
            s5 = next((s for s in slide.shapes if s.shape_id == 5), None)
            s6 = next((s for s in slide.shapes if s.shape_id == 6), None)
            self.assertIsNotNone(s5)
            self.assertIsNotNone(s6)
            self.assertIn("Sep to", s5.text_frame.text)
            self.assertEqual(s6.image.size, (547, 486))

    def test_sanitize_chart_links_normalizes_external_targets_to_peer_excel(self):
        from report_tools.quarterly_quality.office_charts import _sanitize_chart_links, _check_links, OfficeChartCopyError
        with TemporaryDirectory() as folder:
            folder = Path(folder)
            temp_source = folder / "office-qbr-12345" / "source.xlsx"
            pptx_path = folder / "test.pptx"
            with ZipFile(pptx_path, "w") as z:
                z.writestr("ppt/slides/slide1.xml", "<p:sld/>")
                rel_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" '
                    f'Target="{temp_source.as_uri()}" TargetMode="External"/>'
                    f'</Relationships>'
                ).encode("utf-8")
                z.writestr("ppt/charts/_rels/chart1.xml.rels", rel_xml)

            # Before sanitize: _check_links must raise error because of temp folder
            with self.assertRaises(OfficeChartCopyError):
                _check_links(pptx_path, temp_source)

            # Sanitize to peer file name
            _sanitize_chart_links(pptx_path, "HFPA-FTT-FY27-Q2.xlsx")

            # After sanitize: _check_links must succeed
            _check_links(pptx_path, temp_source)

            # Verify target content inside zip
            with ZipFile(pptx_path) as z:
                rels_content = z.read("ppt/charts/_rels/chart1.xml.rels").decode("utf-8")
                self.assertIn('Target="HFPA-FTT-FY27-Q2.xlsx"', rels_content)
                self.assertNotIn("office-qbr", rels_content)


if __name__ == "__main__":
    unittest.main()
