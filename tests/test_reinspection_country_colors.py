"""Country slice/legend colours must survive different factory rankings."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from report_tools.reinspection import pptx_generator as generator


NS = {'c': generator.NS_C, 'a': generator.NS_A}
TEMPLATE = Path(__file__).resolve().parents[1] / 'report_tools/reinspection/RE-INS REPORT.APR.2026.pptx'
COUNTRIES = (
    ('USA', 'CHINA', 'BELGIUM', 'JAPAN', 'MEXICO', 'OTHERS'),
    ('JAPAN', 'USA', 'CHINA', 'SOUTH KOREA', 'FRANCE', 'OTHERS'),
    ('BELGIUM', 'MEXICO', 'USA', 'JAPAN', 'CHINA', 'OTHERS'),
    ('FRANCE', 'JAPAN', 'SOUTH KOREA', 'CHINA', 'USA', 'OTHERS'),
)


def colours_by_country(series):
    labels = [point.find('c:v', NS).text for point in series.findall('c:cat//c:pt', NS)]
    fills = {int(point.find('c:idx', NS).get('val')):
             point.find('c:spPr/a:solidFill/a:srgbClr', NS).get('val')
             for point in series.findall('c:dPt', NS)}
    return {name: fills[index] for index, name in enumerate(labels)}


class ReinspectionCountryColourTests(unittest.TestCase):
    def template_chart(self, chart=8):
        with ZipFile(TEMPLATE) as package:
            return package.read(f'ppt/charts/chart{chart}.xml')

    def test_four_rankings_use_the_same_country_colours_in_every_series(self):
        shared = {}
        for chart, categories in enumerate(COUNTRIES, 8):
            source = self.template_chart(chart)
            before = ET.fromstring(source)
            values = [40, 25, 15, 10, 7, 3]
            output = generator._update_pie3d_chart(source, categories, values, [v/100 for v in values])
            root = ET.fromstring(output)
            self.assertEqual(ET.tostring(root.find('.//c:legend', NS)),
                             ET.tostring(before.find('.//c:legend', NS)))
            for series_index, series in enumerate(root.findall('.//c:ser', NS)):
                colours = colours_by_country(series)
                self.assertEqual(len(series.findall('c:dPt', NS)), len(categories))
                for country, colour in colours.items():
                    if country in shared:
                        self.assertEqual(colour, shared[country])
                    shared[country] = colour
                    self.assertEqual(colour, generator.COUNTRY_COLORS[country])
                original_line = before.findall('.//c:ser', NS)[series_index].find('c:dPt/c:spPr/a:ln', NS)
                current_line = series.find('c:dPt/c:spPr/a:ln', NS)
                self.assertEqual(ET.tostring(current_line) if current_line is not None else None,
                                 ET.tostring(original_line) if original_line is not None else None)
                children = [element.tag.rsplit('}', 1)[-1] for element in series]
                self.assertLess(max(i for i, tag in enumerate(children) if tag == 'dPt'), children.index('cat'))
            actual = root.findall('.//c:ser', NS)[0].findall('c:val//c:pt/c:v', NS)
            self.assertEqual([int(v.text) for v in actual], values)
        self.assertEqual(shared['OTHERS'], 'CBCBCB')

    def test_new_countries_aliases_and_shorter_rankings_do_not_reuse_old_colours(self):
        categories = ['  usa\xa0', 'United States', 'NEW COUNTRY', ' new\xa0country ', 'OTHERS']
        source = self.template_chart()
        # Also exercise the 2D fallback handled by the same updater.
        root = ET.fromstring(source)
        root.find('.//c:pie3DChart', NS).tag = f'{{{generator.NS_C}}}pieChart'
        output = generator._update_pie3d_chart(ET.tostring(root), categories, [1]*5, [.2]*5)
        for series in ET.fromstring(output).findall('.//c:ser', NS):
            colours = colours_by_country(series)
            self.assertEqual(colours[categories[0]], colours[categories[1]])
            self.assertEqual(colours[categories[2]], colours[categories[3]])
            self.assertEqual(colours[categories[0]], generator.COUNTRY_COLORS['USA'])
        shorter = generator._update_pie3d_chart(output, ['NEW COUNTRY', 'USA'], [3, 7], [.3, .7])
        for series in ET.fromstring(shorter).findall('.//c:ser', NS):
            self.assertEqual(len(series.findall('c:dPt', NS)), 2)
            self.assertEqual(colours_by_country(series)['NEW COUNTRY'], colours['NEW COUNTRY'])

    def test_export_pipeline_applies_country_colours_to_all_four_factory_charts(self):
        sites = {}
        for site, countries in zip(('VH', 'VH2', 'JV', 'JV2'), COUNTRIES):
            sites[site] = dict(qc_total_fail=100, top_countries=list(zip(countries[:5], (40, 25, 15, 10, 7))),
                               top5_defects=[], top_models=[], grand_total_5models=1,
                               models_alphabetical=[], ftt_top3=[], hfpa_top3=[], re_qc_top3=[], re_ma_top3=[])
        with TemporaryDirectory() as temporary, patch.object(generator, 'sync_and_register_defects', return_value={}):
            output = Path(temporary)/'Re-Inspection.pptx'
            generator.export_reinspection_presentation(sites, TEMPLATE, output, month_label='October 2026', log_fn=lambda *_: None)
            with ZipFile(output) as package:
                for chart in range(8, 12):
                    root = ET.fromstring(package.read(f'ppt/charts/chart{chart}.xml'))
                    for series in root.findall('.//c:ser', NS):
                        for country, colour in colours_by_country(series).items():
                            self.assertEqual(colour, generator.COUNTRY_COLORS[country])


if __name__ == '__main__':
    unittest.main()
