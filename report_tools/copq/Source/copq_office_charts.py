"""Transfer the existing FTT charts through Office Copy/Paste.

The Excel chart is authoritative. No categories, series, values, colors or
labels are rebuilt in PowerPoint. The cell-based keys are copied from Excel
as pictures; the four charts themselves must remain editable chart objects.
"""
from pathlib import Path
import io
import math
import re
import shutil
import struct
import tempfile
import time
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree


SLOTS = {
    'VH': ('Chart 78', 66),
    'VH2': ('Chart 80', 89),
    'JV': ('Chart 92', 95),
    'JV2': ('Chart 96', 99),
}
MODEL_SLOTS = {'VH': 'Chart 24', 'VH2': 'Chart 35', 'JV': 'Chart 38', 'JV2': 'Chart 40'}
SHEET = 'Top Defects Summary'
NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


class ChartTransferError(RuntimeError):
    pass


def _retry(action, seconds=10):
    deadline = time.monotonic() + seconds
    while True:
        try:
            return action()
        except Exception:
            if time.monotonic() >= deadline:
                raise
            time.sleep(.15)


def _application(client, names, collection):
    failures = []
    for name in names:
        try:
            app = client.DispatchEx(name)
            getattr(app, collection).Count
            return app
        except Exception as error:
            failures.append(error)
    raise ChartTransferError('Không mở được Excel/PowerPoint hoặc WPS để Copy/Paste biểu đồ COPQ.') from failures[-1]


def _source_blocks(sheet):
    blocks = {}
    for row in range(1, int(sheet.UsedRange.Rows.Count) + 1):
        site = sheet.Cells(row, 1).Value
        if site not in SLOTS:
            continue
        if site in blocks:
            raise ChartTransferError(f'FTT combined có hai bảng dữ liệu của site {site}.')
        count = 0
        while sheet.Cells(row, count + 2).Value:
            count += 1
        if not count:
            raise ChartTransferError(f'FTT combined thiếu danh sách lỗi của {site}.')
        # Legend header is merged across its two columns. Discover its location
        # by its content rather than assuming a fixed column U.
        legends = [col for col in range(2, int(sheet.UsedRange.Columns.Count) + 1)
                   if sheet.Cells(row, col).Value == f'Site {site}']
        if len(legends) != 1:
            raise ChartTransferError(f'Không xác định được bảng chú giải Excel của {site}.')
        blocks[site] = (row, count, legends[0])
    if set(blocks) != set(SLOTS):
        raise ChartTransferError('FTT combined chưa có đủ bảng Top Defects Summary của VH, VH2, JV, JV2.')
    return blocks


def _chart_header(formula):
    # The first SERIES argument refers to the defect header, e.g. $B$13.
    match = re.search(r"(?:'((?:[^']|'')+)'|([^,()!]+))!\$?[A-Z]+\$?(\d+)", str(formula))
    if not match:
        return None
    sheet = re.sub(r'\[[^]]+\]', '', match[1] or match[2]).replace("''", "'")
    return (sheet, int(match[3]))


def _source_charts(sheet, blocks):
    rows = {block[0]: site for site, block in blocks.items()}
    charts = {}
    objects = sheet.ChartObjects()
    for index in range(1, objects.Count + 1):
        obj = objects.Item(index)
        key = _chart_header(obj.Chart.SeriesCollection(1).Formula)
        if key is None or key[0] != SHEET or key[1] not in rows:
            continue
        site = rows[key[1]]
        if site in charts:
            raise ChartTransferError(f'FTT combined có nhiều biểu đồ cùng nguồn cho {site}.')
        charts[site] = obj
    if set(charts) != set(SLOTS):
        raise ChartTransferError('FTT combined chưa có đủ 4 biểu đồ Excel để Copy/Paste; không dùng lại dữ liệu mẫu.')
    return charts


def _signature(chart, sheet):
    series = chart.SeriesCollection()
    result = []
    for i in range(1, series.Count + 1):
        item = series.Item(i)
        refs = re.findall(r"!((?:\$?[A-Z]+\$?\d+)(?::\$?[A-Z]+\$?\d+)?)", str(item.Formula))
        if len(refs) != 3:
            raise ChartTransferError('Không đọc được vùng dữ liệu biểu đồ Excel để đối chiếu sau Paste.')
        def cells(ref):
            value = sheet.Range(ref).Value
            return tuple(row[0] for row in value) if isinstance(value, tuple) else (value,)
        result.append((str(cells(refs[0])[0]), tuple(str(v or '') for v in cells(refs[1])),
                       tuple(float(v or 0) for v in cells(refs[2]))))
    return tuple(result)


def _saved_signature(xml):
    def points(element, numeric=False):
        if element is None:
            raise ChartTransferError('Biểu đồ đã dán không có dữ liệu cache để đối chiếu.')
        count = element.find('c:ptCount', NS)
        if count is None:
            raise ChartTransferError('Biểu đồ đã dán thiếu số lượng điểm dữ liệu.')
        values = [0.0 if numeric else ''] * int(count.get('val'))
        for point in element.findall('c:pt', NS):
            text = point.findtext('c:v', default='', namespaces=NS)
            values[int(point.get('idx'))] = float(text or 0) if numeric else text
        return tuple(values)
    root = etree.fromstring(xml)
    result = []
    for series in root.findall('.//c:ser', NS):
        title = series.findtext('c:tx/c:strRef/c:strCache/c:pt/c:v', namespaces=NS)
        if title is None:
            title = series.findtext('c:tx/c:v', default='', namespaces=NS)
        cats = series.find('c:cat/c:strRef/c:strCache', NS)
        if cats is None:
            cats = series.find('c:cat/c:strLit', NS)
        vals = series.find('c:val/c:numRef/c:numCache', NS)
        if vals is None:
            vals = series.find('c:val/c:numLit', NS)
        result.append((title, points(cats), points(vals, numeric=True)))
    return tuple(result)


def _same_data(first, second):
    return (len(first) == len(second) and all(
        a[:2] == b[:2] and len(a[2]) == len(b[2]) and
        all(math.isclose(x, y, rel_tol=1e-10, abs_tol=1e-12) for x, y in zip(a[2], b[2]))
        for a, b in zip(first, second)))


def _shape(slide, *, name=None, ident=None):
    found = [slide.Shapes.Item(i) for i in range(1, slide.Shapes.Count + 1)
             if (name is not None and slide.Shapes.Item(i).Name == name) or
             (ident is not None and slide.Shapes.Item(i).Id == ident)]
    if len(found) != 1:
        raise ChartTransferError(f'Không tìm được đúng một vị trí {name or ident} trong mẫu COPQ.')
    return found[0]


def _position(shape, box):
    shape.LockAspectRatio = 0
    shape.Left, shape.Top, shape.Width, shape.Height = box


def _fit_source_chart(obj, box):
    """Scale the Excel chart to its template slot before copying it.

    WPS does not scale chart fonts when a pasted chart is resized. Apply the
    proportional font sizes in Excel so the delivered workbook and PPT agree.
    Only geometry and typography change; source data and colors stay intact.
    """
    chart = obj.Chart
    ratio = min(box[2] / float(obj.Width), box[3] / float(obj.Height))
    title_size = max(6, min(8, float(chart.ChartTitle.Font.Size) * ratio))
    obj.Width, obj.Height = box[2], box[3]
    chart.ChartTitle.Font.Size = title_size
    labels = chart.Axes(1).TickLabels
    labels.Font.Size = 4.5
    labels.Orientation = -4128  # xlHorizontal, also supported by WPS
    for i in range(1, chart.SeriesCollection().Count + 1):
        chart.SeriesCollection(i).DataLabels().Font.Size = 4.5


def _legend_picture(sheet, block, destination):
    """Copy Excel's rendered cells, without recreating their text or colors."""
    import win32clipboard
    from PIL import Image

    row, count, col = block
    sheet.Activate()
    sheet.Range(sheet.Cells(row, col), sheet.Cells(row + count, col + 1)).CopyPicture(1, 2)

    def read_bitmap():
        win32clipboard.OpenClipboard()
        try:
            return win32clipboard.GetClipboardData(8)  # CF_DIB, rendered by Excel
        finally:
            win32clipboard.CloseClipboard()

    dib = _retry(read_bitmap)
    # BITMAPFILEHEADER precedes the existing DIB. Pillow determines its pixel
    # offset from the DIB header (including palettes and bitfield masks).
    header = struct.pack('<2sIHHI', b'BM', len(dib) + 14, 0, 0, 0)
    with Image.open(io.BytesIO(header + dib)) as picture:
        picture.save(destination)
        return picture.width, picture.height


def _chart_parts(package):
    import posixpath
    slide = etree.fromstring(package.read('ppt/slides/slide2.xml'))
    relations = etree.fromstring(package.read('ppt/slides/_rels/slide2.xml.rels'))
    rels = {rel.get('Id'): rel.get('Target') for rel in relations}
    result = {}
    for frame in slide.findall('.//p:graphicFrame', NS):
        props = frame.find('p:nvGraphicFramePr/p:cNvPr', NS)
        chart = frame.find('.//c:chart', NS)
        if props is not None and chart is not None and props.get('name') in {v[0] for v in SLOTS.values()}:
            target = rels[chart.get('{'+NS['r']+'}id')]
            result[props.get('name')] = posixpath.normpath(posixpath.join('ppt/slides', target))
    if len(result) != 4:
        raise ChartTransferError('PowerPoint sau khi lưu chưa có đủ 4 đối tượng biểu đồ chỉnh sửa được.')
    return result


def _workbook_links(pptx, workbook_path, allow_linked):
    """Only relocate native chart links; never rewrite chart XML or data."""
    changed, linked = {}, []
    with ZipFile(pptx) as package:
        for shape_name, part in _chart_parts(package).items():
            rel_path = str(Path(part).parent / '_rels' / (Path(part).name + '.rels')).replace('\\', '/')
            if rel_path not in package.namelist():
                raise ChartTransferError(f'Biểu đồ {shape_name} không có nguồn dữ liệu chỉnh sửa được.')
            root = etree.fromstring(package.read(rel_path))
            data_links = [rel for rel in root if rel.get('Type', '').rsplit('/', 1)[-1] in ('package', 'oleObject')]
            if not data_links:
                raise ChartTransferError(f'Biểu đồ {shape_name} thiếu workbook để chỉnh sửa dữ liệu.')
            for rel in data_links:
                if rel.get('TargetMode') == 'External':
                    if not allow_linked:
                        raise ChartTransferError('Office/WPS đang dán biểu đồ liên kết. Chưa nhúng được dữ liệu vào PPT; không xuất thay bằng ảnh.')
                    # WPS cannot resolve relative chart workbook links. Use the
                    # delivered workbook's actual location, never the private
                    # staging copy. Both files remain in the report output.
                    rel.set('Target', Path(workbook_path).resolve().as_uri())
                    linked.append(shape_name)
                    changed[rel_path] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        if changed:
            candidate = Path(pptx).with_suffix('.links.pptx')
            with ZipFile(candidate, 'w', ZIP_DEFLATED) as output:
                for item in package.infolist():
                    output.writestr(item, changed.get(item.filename, package.read(item.filename)))
    if changed:
        candidate.replace(pptx)
    return linked


def _publish(source, destination):
    try:
        from pipeline_common import save_with_fallback
    except ImportError:
        from report_tools.copq.pipeline_common import save_with_fallback
    return Path(save_with_fallback(lambda filename: shutil.copy2(source, filename), str(destination))).resolve()


def copy_ftt_charts(pptx, xlsx, *, allow_linked=False, preview=None):
    """Replace the four template slots using actual Office Copy/Paste."""
    import pythoncom
    import win32com.client as client

    pptx, xlsx = Path(pptx).resolve(), Path(xlsx).resolve()
    if not xlsx.is_file():
        raise ChartTransferError('Không tìm thấy FTT combined để Copy/Paste biểu đồ.')
    pythoncom.CoInitialize()
    workbook = presentation = temporary = None
    try:
        excel = _application(client, ('Excel.Application', 'ket.Application'), 'Workbooks')
        powerpoint = _application(client, ('PowerPoint.Application', 'kwpp.Application'), 'Presentations')
        temporary = tempfile.TemporaryDirectory(prefix='copq-paste-', dir=pptx.parent, ignore_cleanup_errors=True)
        with temporary as directory:
            directory = Path(directory)
            candidate = directory / 'native.pptx'
            shutil.copy2(pptx, candidate)
            excel_copy = directory / xlsx.name
            shutil.copy2(xlsx, excel_copy)
            workbook = _retry(lambda: excel.Workbooks.Open(str(excel_copy), 0, False))
            sheet = workbook.Worksheets.Item(SHEET)
            blocks = _source_blocks(sheet)
            charts = _source_charts(sheet, blocks)
            presentation = _retry(lambda: powerpoint.Presentations.Open(str(candidate), False, False, False))
            slide = presentation.Slides.Item(2)
            signatures = {}
            source_top = float(sheet.Cells(max(row + max(3, count) for row, count, _ in blocks.values()) + 3, 1).Top)
            for index, (site, (name, legend_id)) in enumerate(SLOTS.items()):
                slot = _shape(slide, name=name)
                legend = _shape(slide, ident=legend_id)
                model = _shape(slide, name=MODEL_SLOTS[site])
                # Keep the chart inside its existing BY MODEL panel, including
                # the VH slot that protrudes beyond the original panel edge.
                right = min(float(slot.Left + slot.Width), float(model.Left + model.Width))
                bottom = min(float(slot.Top + slot.Height), float(model.Top + model.Height))
                box = (float(slot.Left), float(slot.Top), right - float(slot.Left), bottom - float(slot.Top))
                legend_box = (float(legend.Left), float(legend.Top), float(legend.Width))
                obj = charts[site]
                signatures[name] = _signature(obj.Chart, sheet)
                _fit_source_chart(obj, box)
                obj.Left = (index % 2) * 280
                obj.Top = source_top + (index // 2) * 155
                # Honor the key-column dimensions specified by the Excel
                # generator (older workbooks overwrote them with width 25).
                sheet.Columns.Item(blocks[site][2]).ColumnWidth = 5
                sheet.Columns.Item(blocks[site][2] + 1).ColumnWidth = 34

                def paste_chart():
                    sheet.Activate()
                    obj.Copy()
                    time.sleep(.1)
                    pasted = slide.Shapes.Paste()
                    if pasted is None or pasted.Count != 1:
                        raise ChartTransferError(f'Clipboard chưa dán được biểu đồ {site}.')
                    return pasted.Item(1)

                pasted = _retry(paste_chart)
                if not pasted.HasChart or pasted.Type != 3:
                    raise ChartTransferError(f'Office dán {site} thành ảnh hoặc đối tượng không chỉnh sửa được.')
                # Microsoft Office can detach a linked chart natively. WPS may
                # not implement ChartData.IsLinked/BreakLink; package validation
                # below enforces the selected delivery mode in either case.
                try:
                    if pasted.Chart.ChartData.IsLinked:
                        pasted.Chart.ChartData.BreakLink()
                except Exception:
                    pass
                _position(pasted, box)
                slot.Delete()
                pasted.Name = name

                picture_path = directory / f'legend-{site}.png'
                width, height = _legend_picture(sheet, blocks[site], picture_path)
                picture = slide.Shapes.AddPicture(str(picture_path), False, True,
                                                  legend_box[0], legend_box[1],
                                                  legend_box[2], legend_box[2] * height / width)
                picture.Name = f'FTT legend {site}'
                legend.Delete()
            _retry(presentation.Save)
            presentation.Close()
            presentation = None
            _retry(workbook.Save)
            workbook.Close(False)
            workbook = None
            xlsx = _publish(excel_copy, xlsx)
            linked = _workbook_links(candidate, xlsx, allow_linked)
            with ZipFile(candidate) as package:
                for name, part in _chart_parts(package).items():
                    if not _same_data(signatures[name], _saved_signature(package.read(part))):
                        raise ChartTransferError(f'Biểu đồ {name} sau Paste khác dữ liệu Excel.')
            # Reopen using Office and compare the saved chart data as well.
            presentation = _retry(lambda: powerpoint.Presentations.Open(str(candidate), True, False, False))
            slide = presentation.Slides.Item(2)
            for name in signatures:
                saved = _shape(slide, name=name)
                if not saved.HasChart:
                    raise ChartTransferError(f'Biểu đồ {name} thay đổi sau khi lưu/mở lại PPT.')
                if name in linked:
                    saved.Chart.ChartData.Activate()
                    linked_book = saved.Chart.ChartData.Workbook
                    try:
                        if Path(linked_book.FullName).resolve() != xlsx:
                            raise ChartTransferError(f'Edit Data của {name} không mở đúng FTT combined đi kèm.')
                    finally:
                        linked_book.Close(False)
            if preview:
                slide.Export(str(Path(preview).resolve()), 'PNG', 1800, 1013)
            presentation.Close()
            presentation = None
            pptx = _publish(candidate, pptx)
            return {'charts': 4, 'legends': 4, 'linked_charts': linked,
                    'source': xlsx.name, 'output': str(pptx), 'method': 'Office Copy/Paste'}
    except ChartTransferError:
        raise
    except Exception as error:
        raise ChartTransferError('Copy/Paste biểu đồ COPQ bằng Office chưa hoàn tất. Đóng hộp thoại đang chờ trong Office/WPS rồi xuất lại.') from error
    finally:
        for document, args in ((presentation, ()), (workbook, (False,))):
            if document is not None:
                try:
                    document.Close(*args)
                except Exception:
                    pass
        if temporary is not None:
            temporary.cleanup()
        pythoncom.CoUninitialize()
