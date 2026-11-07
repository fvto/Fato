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
            app = _retry(lambda: client.Dispatch(name))
            _retry(lambda: getattr(app, collection).Count)
            return app
        except Exception as error:
            failures.append(error)
        try:
            app = _retry(lambda: client.DispatchEx(name))
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


def _sanitize_chart_links(path, target_filename):
    """Normalize external chart relationships to the peer Excel file name."""
    temp_pkg = Path(path).with_suffix('.relclean.pptx')
    with ZipFile(path) as zin, ZipFile(temp_pkg, 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith('.rels') and 'charts' in item.filename.replace('\\', '/'):
                root = ElementTree.fromstring(data)
                changed = False
                for rel in root:
                    if rel.get('TargetMode') == 'External':
                        rel.set('Target', target_filename)
                        changed = True
                if changed:
                    data = ElementTree.tostring(root, xml_declaration=True, encoding='utf-8')
            zout.writestr(item, data)
    shutil.move(str(temp_pkg), str(path))


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


def copy_charts(pptx, xlsx, output_pptx, output_xlsx=None):
    from .editable_charts import CHART_SPECS, chart_rectangle, RANK_PLOT_X, RANK_PLOT_WIDTH
    try:
        import pythoncom
        import win32com.client as client
    except ImportError as error:
        raise OfficeChartCopyError('Thiếu pywin32 để Copy/Paste chart bằng Office. Cài lại dependencies của Dashboard.') from error
    pythoncom.CoInitialize()
    excel = powerpoint = workbook = presentation = temporary_files = None
    try:
        excel = _application(client, ('Excel.Application', 'ket.Application'), 'Workbooks')
        powerpoint = _application(client, ('PowerPoint.Application', 'kwpp.Application'), 'Presentations')
        temporary_files = tempfile.TemporaryDirectory(prefix='office-qbr-', ignore_cleanup_errors=True)
        with temporary_files as temporary:
            temporary = Path(temporary)
            source = temporary/'source.pptx'
            source.write_bytes(pptx.read() if hasattr(pptx, 'read') else Path(pptx).read_bytes())
            candidate = temporary/'finished.pptx'
            # Work on a private Excel copy; publish the Office-saved source at
            # the end only after both outputs have been saved and reopened.
            excel_copy = temporary/'source.xlsx'
            shutil.copyfile(xlsx, excel_copy)
            workbook = _retry(lambda: excel.Workbooks.Open(str(excel_copy.resolve())))
            presentation = _retry(lambda: powerpoint.Presentations.Open(str(source.resolve()), False, False, False))
            width, height = float(presentation.PageSetup.SlideWidth), float(presentation.PageSetup.SlideHeight)
            charts = _charts(workbook, CHART_SPECS)
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
            _retry(workbook.Save)
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
                        p = slide.Shapes.Paste()
                        if p.Count >= 1:
                            return p.Item(1)
                        raise OfficeChartCopyError('Clipboard chưa sẵn sàng để Paste.')

                    shape = _retry(_paste_chart, timeout=10)
                    if shape.Type in (11, 13):
                        raise OfficeChartCopyError('Office đã Paste thành ảnh thay vì đối tượng chart. Chưa xuất PPT.')
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
            # Save the native presentation in place; Office authors the entire package.
            _retry(presentation.Save)
            _retry(presentation.Close)
            presentation = None
            _retry(lambda: workbook.Close(False))
            workbook = None
            _sanitize_chart_links(source, Path(output_xlsx or xlsx).name)
            _check_links(source, excel_copy)
            shutil.copyfile(source, output_pptx)
            shutil.copyfile(excel_copy, output_xlsx or xlsx)
        return 12
    except OfficeChartCopyError:
        raise
    except Exception as error:
        raise OfficeChartCopyError('Copy/Paste chart bằng Office chưa hoàn tất. Đóng các hộp thoại Excel/PowerPoint/WPS rồi xuất lại.') from error
    finally:
        for document, close_args in ((presentation, ()), (workbook, (False,))):
            if document is not None:
                try:
                    document.Close(*close_args)
                except Exception:
                    pass
        if temporary_files is not None:
            temporary_files.cleanup()
        pythoncom.CoUninitialize()
