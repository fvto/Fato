from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile

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


if __name__ == "__main__":
    unittest.main()
