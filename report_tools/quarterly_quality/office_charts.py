"""Use Office's actual chart Copy/Paste and native PPTX writer."""
from pathlib import Path
import re
import shutil
import tempfile
import time
from zipfile import ZipFile
from xml.etree import ElementTree


class OfficeChartCopyError(RuntimeError):
    pass


def _retry(action, timeout=15):
    deadline = time.monotonic() + timeout
    while True:
        try:
            return action()
        except Exception as error:
            if getattr(error, 'hresult', None) not in (-2147418111, -2147417846) or time.monotonic() >= deadline:
                raise
            time.sleep(.2)


def _application(client, names, collection):
    failures = []
    for name in names:
        try:
            app = _retry(lambda: client.DispatchEx(name))
            _retry(lambda: getattr(app, collection).Count)
            return app
        except Exception as error:
            failures.append(error)
        try:
            app = _retry(lambda: client.Dispatch(name))
            _retry(lambda: getattr(app, collection).Count)
            return app
        except Exception as error:
            failures.append(error)
    raise OfficeChartCopyError('Không mở được Excel/WPS hoặc PowerPoint/WPS để Copy/Paste chart. '
                              'Cần cài ứng dụng Office trên máy và đóng các hộp thoại đang chờ.') from failures[-1]


def _signature(chart):
    sheets, rows = set(), []
    series = chart.SeriesCollection()
    for index in range(1, series.Count + 1):
        formula = str(series.Item(index).Formula)
        refs = re.findall(r"(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ ]*))!\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?\d+)?", formula)
        if not refs:
            return None
        quoted, plain, row = refs[-1]
        sheet = re.sub(r'\[[^]]+\]', '', quoted or plain).replace("''", "'")
        sheets.add(sheet)
        rows.append(int(row))
    return (next(iter(sheets)), tuple(rows)) if len(sheets) == 1 else None


def _charts(workbook, specs):
    expected = {(sheet, rows): (metric, region, kind) for metric, region, kind, _, sheet, rows in specs}
    found = {}
    for index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets.Item(index)
        objects = sheet.ChartObjects()
        for number in range(1, objects.Count + 1):
            obj = objects.Item(number)
            key = expected.get(_retry(lambda: _signature(obj.Chart)))
            if key is None:
                continue
            if key in found:
                raise OfficeChartCopyError('Có chart Excel trùng nguồn dữ liệu; chưa đưa vào PPT.')
            found[key] = obj
    if set(found) != set(expected.values()):
        raise OfficeChartCopyError('Excel chưa có đủ 12 chart quý để Copy/Paste sang PPT.')
    return found


def _slide(presentation, metric):
    result = []
    for index in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides.Item(index)
        for number in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes.Item(number)
            if shape.HasTextFrame and shape.TextFrame.HasText:
                try:
                    text = shape.TextFrame.TextRange.Text
                except Exception:
                    # WPS reports HasTextFrame for some non-text shapes.
                    continue
                if f'QMS {metric} ANALYSIS' in text:
                    result.append(slide)
                    break
    if len(result) != 1:
        raise OfficeChartCopyError(f'Không tìm được đúng một slide QMS {metric}.')
    return result[0]


def _shape_by_id(slide, ident):
    for index in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes.Item(index)
        if shape.Id == ident:
            return shape
    return None


def _position(shape, box):
    # PPT position/size only. Never access pasted.Chart for styling or data.
    x, y, width, height = box
    shape.LockAspectRatio = 0
    shape.Left, shape.Top = x, y
    shape.Width, shape.Height = width, height


def _check_links(path, source):
    """Read-only check: Office must not leave links to the private workbook."""
    needle = str(Path(source).parent).replace('\\', '/').lower()
    with ZipFile(path) as package:
        for name in package.namelist():
            if not name.endswith('.rels'):
                continue
            for relation in ElementTree.fromstring(package.read(name)):
                target = relation.get('Target', '').replace('\\', '/').lower()
                if relation.get('TargetMode') == 'External' and needle in target:
                    raise OfficeChartCopyError('Chart vẫn liên kết tới Excel tạm. Chưa xuất PPT để tránh lỗi khi gửi người khác.')


def _summary_sheet(workbook, quarter, columns):
    """Build the screenshot in Excel; every month is selected from the run."""
    from openpyxl.utils.cell import get_column_letter
    from .quarterly_quality import REGIONS

    name = 'QBR HFPA Summary'
    try:
        sheet = workbook.Worksheets.Item(name)
        sheet.Cells.Clear()
    except Exception:
        sheet = workbook.Worksheets.Add()
        sheet.Name = name
    area = sheet.Range('A1:E4')
    area.Font.Name = 'Times New Roman'
    area.Font.Size = 20
    area.HorizontalAlignment = -4108
    area.VerticalAlignment = -4108
    area.Interior.Color = 16777215
    area.RowHeight = 27
    area.ColumnWidth = 12
    area.Borders.LineStyle = 1
    area.Borders.Color = 14474460
    area.Borders.Weight = 1
    sheet.Cells(1, 1).Value = 'HFPA'
    sheet.Cells(1, 5).Value = 'Avg'
    for row, region in ((2, 'VN'), (3, 'Indo')):
        sheet.Cells(row, 1).Value = region
        for col, month in enumerate(quarter.months, 2):
            sheet.Cells(1, col).Value = int(month[-2:])
            records = [quarter.values.get(('HFPA', month, plant)) for plant in REGIONS[region]]
            cell = sheet.Cells(row, col)
            if month not in quarter.covered or any(record is None for record in records):
                cell.Value = None
                continue
            source = get_column_letter(columns[month])
            first = 5 if region == 'VN' else 7
            cell.Formula = (f'=IFERROR((1-SUM(\'Month HFPA\'!{source}{first+9}:{source}{first+10})/'
                            f'SUM(\'Month HFPA\'!{source}{first}:{source}{first+1}))*100,"")')
        sheet.Cells(row, 5).Formula = f'=IF(COUNT(B{row}:D{row})=3,AVERAGE(B{row}:D{row}),"")'
    sheet.Range('B2:E3').NumberFormat = '0.0'
    sheet.Range('E1:E3').BorderAround(1, 3, Color=0)
    sheet.Range('A4:D4').ClearContents()
    sheet.Range('E4').Formula = '=IF(COUNT(E2:E3)=2,AVERAGE(E2:E3),"")'
    sheet.Range('E4').NumberFormat = '0'
    sheet.Calculate()
    return sheet.Range('A1:E4')


def _paste_summary(presentation, area, image_path):
    """Only copy/paste and placement happen in PPT; no table is drawn there."""
    slides = []
    for i in range(1, presentation.Slides.Count + 1):
        slide = presentation.Slides.Item(i)
        for j in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes.Item(j)
            if not shape.HasTextFrame:
                continue
            try:
                text = str(shape.TextFrame.TextRange.Text).upper()
            except Exception:
                continue  # WPS reports HasTextFrame for some non-text objects.
            if 'CONSUMER PROMISE REVIEW' in text:
                slides.append(slide)
                break
    for slide in slides:
        slot = _shape_by_id(slide, 6)
        if slot is None:
            raise OfficeChartCopyError('Thiếu vị trí note HFPA trên slide Consumer Promise Review.')
        left, top, width = float(slot.Left), float(slot.Top), float(slot.Width)
        area.Parent.Activate()
        area.CopyPicture(1, 2)  # xlScreen, xlBitmap: a real Excel screenshot.
        # WPS does not implement bitmap PasteSpecial consistently. Preserve
        # Excel's exact clipboard screenshot as PNG, then paste that picture.
        import win32clipboard
        import struct
        from io import BytesIO
        from PIL import Image

        def read_bitmap():
            win32clipboard.OpenClipboard()
            try:
                return win32clipboard.GetClipboardData(8)  # CF_DIB
            finally:
                win32clipboard.CloseClipboard()
        dib = _retry(read_bitmap)
        header = struct.pack('<2sIHHI', b'BM', len(dib) + 14, 0, 0, 0)
        with Image.open(BytesIO(header + dib)) as picture:
            picture.save(image_path, 'PNG')
            image_width, image_height = picture.size
        shape = slide.Shapes.AddPicture(str(image_path), False, True, left, top,
                                        width, width * image_height / image_width)
        shape.Name = 'HFPA quarterly summary from Excel'
        slot.Delete()


def copy_charts(pptx, xlsx, output_pptx, output_xlsx=None, *, quarter=None):
    from .editable_charts import CHART_SPECS, chart_rectangle, RANK_PLOT_X, RANK_PLOT_WIDTH
    try:
        import pythoncom
        import win32com.client as client
    except ImportError as error:
        raise OfficeChartCopyError('Thiếu pywin32 để Copy/Paste chart bằng Office. Cài lại dependencies của Dashboard.') from error
    pythoncom.CoInitialize()
    excel = powerpoint = workbook = presentation = temporary_files = candidate = None
    try:
        excel = _application(client, ('Excel.Application', 'ket.Application'), 'Workbooks')
        powerpoint = _application(client, ('PowerPoint.Application', 'kwpp.Application'), 'Presentations')
        # WPS can deadlock opening presentations without an application window.
        _retry(lambda: setattr(powerpoint, 'Visible', True))
        temporary_files = tempfile.TemporaryDirectory(prefix='office-qbr-', ignore_cleanup_errors=True)
        with temporary_files as temporary:
            temporary = Path(temporary)
            source = temporary/'source.pptx'
            source.write_bytes(pptx.read() if hasattr(pptx, 'read') else Path(pptx).read_bytes())
            # Save beside the delivered Excel so Office can write native peer
            # workbook links. The final bytes are never modified afterwards.
            candidate = Path(output_pptx).resolve().with_name('.' + Path(output_pptx).stem + '.native.pptx')
            # Prepare a private copy, then save the finished Excel at its
            # delivered path before Copy, so WPS cannot retain a private link.
            excel_copy = temporary/'source.xlsx'
            shutil.copyfile(xlsx, excel_copy)
            workbook = _retry(lambda: excel.Workbooks.Open(str(excel_copy.resolve()), UpdateLinks=0))
            presentation = _retry(lambda: powerpoint.Presentations.Open(str(source.resolve()), False, False, True))
            width, height = float(presentation.PageSetup.SlideWidth), float(presentation.PageSetup.SlideHeight)
            charts = _charts(workbook, CHART_SPECS)
            summary = None
            if quarter is not None:
                from openpyxl import load_workbook
                from .template_output import month_columns
                book = load_workbook(excel_copy, read_only=True, data_only=True)
                try:
                    columns = month_columns(book['Month HFPA'], 3, 4, None)
                finally:
                    book.close()
                summary = _retry(lambda: _summary_sheet(workbook, quarter, columns))
            boxes = {}
            origins = {}
            for key, obj in charts.items():
                sheet = str(obj.Parent.Name)
                previous = origins.get(sheet, (float(obj.Left), float(obj.Top)))
                origins[sheet] = min(previous[0], float(obj.Left)), min(previous[1], float(obj.Top))
            for metric, region, kind, *_ in CHART_SPECS:
                box = chart_rectangle(width * 12700, height * 12700, region, kind)
                boxes[metric, region, kind] = tuple(value / 12700 for value in box)
                obj = charts[metric, region, kind]
                # Design stays in Excel. Match the final aspect ratio before
                # Copy so text and labels are not distorted after Paste.
                _retry(lambda: setattr(obj, 'Width', boxes[metric, region, kind][2]))
                _retry(lambda: setattr(obj, 'Height', boxes[metric, region, kind][3]))
                left, top = origins[str(obj.Parent.Name)]
                # Separate the Excel charts as well, including the wider trend
                # aspect ratio. HFPA rankings live on a separate worksheet.
                _retry(lambda: setattr(obj, 'Left', left + (650 if region == 'Indo' else 0)))
                offset = {'trend': 0, 'model': 240, 'defect': 370}[kind]
                _retry(lambda: setattr(obj, 'Top', top + offset))
            published_excel = Path(output_xlsx or xlsx).resolve()
            alerts = excel.DisplayAlerts
            try:
                excel.DisplayAlerts = False
                _retry(lambda: workbook.SaveAs(str(published_excel), 51))
            finally:
                excel.DisplayAlerts = alerts
            for metric in ('FTT', 'HFPA'):
                slide = _slide(presentation, metric)
                for owner, region, kind, ident, *_ in CHART_SPECS:
                    if owner != metric:
                        continue
                    slot = _shape_by_id(slide, ident)
                    if slot is None:
                        raise OfficeChartCopyError(f'Thiếu vị trí chart {metric}/{region}/{kind} trong mẫu PPT.')
                    obj = charts[metric, region, kind]
                    # This is the same Copy -> Paste operation as the user.
                    # No CopyPicture, image conversion or chart-XML insertion.
                    def _paste_chart():
                        obj.Parent.Activate()
                        obj.Copy()
                        time.sleep(0.08)
                        p = slide.Shapes.PasteSpecial(DataType=0, Link=False)
                        if p.Count >= 1:
                            return p.Item(1)
                        raise OfficeChartCopyError('Clipboard chưa sẵn sàng để Paste.')

                    shape = _retry(_paste_chart, timeout=10)
                    if shape.Type != 3 or not shape.HasChart:
                        raise OfficeChartCopyError('Office chưa Paste thành chart chỉnh sửa được. Chưa xuất PPT.')
                    shape.Name = f'{metric} {region} {kind}'
                    _position(shape, boxes[metric, region, kind])
                    slot.Delete()
                for region, ident in (('VN', 8 if metric == 'FTT' else 22), ('Indo', 12 if metric == 'FTT' else 21)):
                    arrow = _shape_by_id(slide, ident)
                    if arrow is None:
                        continue
                    x, y, w, h = boxes[metric, region, 'model']
                    defect_y = boxes[metric, region, 'defect'][1]
                    gap = defect_y - y - h
                    aw, ah = width*.025, min(height*.035, gap*.7)
                    centre = x + w*(RANK_PLOT_X + RANK_PLOT_WIDTH/6)
                    _position(arrow, (centre-aw/2, y+h+(gap-ah)/2, aw, ah))
            if summary is not None:
                _paste_summary(presentation, summary, temporary/'hfpa-summary.png')
                _retry(workbook.Save)
            # Office owns the final package. Never patch its XML, links or
            # chart caches after this save: that can trigger PowerPoint Repair.
            _retry(lambda: presentation.SaveAs(str(candidate.resolve()), 24))
            _retry(presentation.Close)
            presentation = None
            _retry(lambda: workbook.Close(False))
            workbook = None
            _check_links(candidate, excel_copy)
            presentation = _retry(lambda: powerpoint.Presentations.Open(str(candidate.resolve()), True, False, True))
            for metric in ('FTT', 'HFPA'):
                saved = _slide(presentation, metric)
                count = sum(bool(saved.Shapes.Item(i).HasChart) for i in range(1, saved.Shapes.Count + 1))
                if count != 6:
                    raise OfficeChartCopyError(f'PowerPoint mở lại không đủ 6 chart {metric}.')
            _retry(presentation.Close)
            presentation = None
            shutil.copyfile(candidate, output_pptx)
            candidate.unlink()
        return 12
    except OfficeChartCopyError:
        raise
    except Exception as error:
        raise OfficeChartCopyError('Copy/Paste chart bằng Office chưa hoàn tất. Đóng các hộp thoại Excel/PowerPoint/WPS rồi xuất lại.') from error
    finally:
        for document, close_args in ((presentation, ()), (workbook, (False,))):
            if document is not None:
                try:
                    if document is presentation:
                        document.Saved = True
                    document.Close(*close_args)
                except Exception:
                    pass
        if temporary_files is not None:
            temporary_files.cleanup()
        if candidate is not None and candidate.is_file():
            try:
                candidate.unlink()
            except OSError:
                pass
        pythoncom.CoUninitialize()
