from pathlib import Path
from io import BytesIO
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import MagicMock, PropertyMock, patch
from zipfile import ZipFile

from report_tools.quarterly_quality import office_charts
from report_tools.quarterly_quality.editable_charts import CHART_SPECS


class OfficeTransferTests(unittest.TestCase):
    def test_publish_keeps_native_saved_bytes_and_never_edits_pasted_chart(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source, output = root/'source.xlsx', root/'report.pptx'
            source.write_bytes(b'Excel fixture')
            excel, powerpoint, workbook, presentation = [MagicMock() for _ in range(4)]
            excel.Workbooks.Open.return_value = workbook
            powerpoint.Presentations.Open.return_value = presentation
            presentation.PageSetup.SlideWidth = 960
            presentation.PageSetup.SlideHeight = 540
            charts = {}
            for metric, region, kind, _, sheet, _ in CHART_SPECS:
                chart = MagicMock()
                chart.Parent.Name = sheet
                chart.Left = chart.Top = 0
                charts[metric, region, kind] = chart
            slide, pasted = MagicMock(), MagicMock()
            slide.Shapes.Count = 6
            slide.Shapes.Item.return_value.HasChart = True
            slide.Shapes.PasteSpecial.return_value.Count = 1
            slide.Shapes.PasteSpecial.return_value.Item.return_value = pasted
            pasted.Type = 3
            type(pasted).Chart = PropertyMock(side_effect=AssertionError('PPT chart must not be edited'))
            saved = []

            def native_save(filename, kind):
                self.assertEqual(kind, 24)
                with ZipFile(filename, 'w') as package:
                    package.writestr('native-writer.txt', 'Native Office bytes, unchanged')
                saved.append(Path(filename).read_bytes())

            presentation.SaveAs.side_effect = native_save
            with patch.object(office_charts, '_application', side_effect=[excel, powerpoint]), \
                    patch.object(office_charts, '_charts', return_value=charts), \
                    patch.object(office_charts, '_slide', return_value=slide), \
                    patch.object(office_charts, '_shape_by_id', return_value=MagicMock()), \
                    patch.object(office_charts, '_position'), \
                    patch.object(office_charts.time, 'sleep'):
                self.assertEqual(office_charts.copy_charts(BytesIO(b'fixture'), source, output), 12)
            self.assertEqual(output.read_bytes(), saved[0])
            self.assertEqual(slide.Shapes.PasteSpecial.call_count, 12)
            for call in slide.Shapes.PasteSpecial.call_args_list:
                self.assertEqual(call.kwargs, dict(DataType=0, Link=False))
            workbook.SaveAs.assert_called_once_with(str(source.resolve()), 51)


if __name__ == '__main__':
    unittest.main()
