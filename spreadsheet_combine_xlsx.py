"""Streaming OOXML values: bounded XML buffers, literal strings, and ZIP64."""
from collections import OrderedDict
from datetime import date, datetime
from functools import lru_cache
import hashlib
import json
import math
import posixpath
import zipfile
from xml.sax.saxutils import escape, quoteattr

from lxml import etree as ET
from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.datetime import from_excel, to_excel, CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
TAG = '{' + NS + '}'
MAX_ROWS = 1_048_576
MAX_UNPACKED = 8 * 1024**3
V_TAG, F_TAG, T_TAG = TAG+'v', TAG+'f', TAG+'t'


@lru_cache(maxsize=16384)
def column_index(letters):
    return column_index_from_string(letters)-1


def canonical(row):
    values = []
    for cell in row:
        value = cell['value'] if isinstance(cell, dict) else cell
        if isinstance(value, (datetime, date)):
            value = to_excel(value)
        if value == '':
            value = None
        if type(value) in (float, int):
            value = ['n', format(value, '.12g')]
        values.append(value)
    return (json.dumps(values, ensure_ascii=False, separators=(',', ':')) + '\n').encode('utf-8')


class Writer:
    def __init__(self, path):
        self.archive = zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED,
                                       compresslevel=1, allowZip64=True)
        self.sheets = []
        self.formats = OrderedDict()
        self.current = None
        self.buffer = bytearray()

    def _write(self, text):
        self.buffer.extend(text.encode('utf-8'))
        if len(self.buffer) >= 1024 * 1024:
            self.current['stream'].write(self.buffer)
            self.buffer.clear()

    def sheet(self, name, headers=None, widths=None):
        self.end_sheet()
        item = dict(name=name, rows=0, columns=0, digest=hashlib.sha256(), headers=headers,
                    stream=self.archive.open(f'xl/worksheets/sheet{len(self.sheets)+1}.xml', 'w', force_zip64=True))
        self.sheets.append(item)
        self.current = item
        self._write(f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{NS}"><sheetViews><sheetView workbookViewId="0" showGridLines="0">')
        if headers:
            self._write('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
        self._write('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="20"/>')
        if headers:
            self._write('<cols>')
            for i, label in enumerate(headers, 1):
                width = max(14, min(52, len(label) * .8))
                if widths:
                    width = widths[i-1]
                elif 'source file' in label.casefold() or label.casefold() == 'file':
                    width = 46
                elif label.casefold() == 'sha256':
                    width = 72
                self._write(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')
            self._write('</cols>')
        self._write('<sheetData>')
        if headers:
            self.row([dict(value=value, format='General') for value in headers], header=True)

    def row(self, cells, header=False):
        item = self.current
        item['rows'] += 1
        if item['rows'] > MAX_ROWS:
            raise ValueError('Trang tính vượt giới hạn dòng Excel.')
        item['columns'] = max(item['columns'], len(cells))
        item['digest'].update(canonical(cells))
        number = item['rows']
        parts = [f'<row r="{number}"' + (' ht="46" customHeight="1"' if header else '') + '>']
        for index, entry in enumerate(cells, 1):
            value = entry.get('value') if isinstance(entry, dict) else entry
            fmt = entry.get('format', 'General') if isinstance(entry, dict) else 'General'
            style = 1 if header else 0
            if not header and fmt and fmt != 'General':
                if fmt not in self.formats:
                    self.formats[fmt] = len(self.formats) + 2
                style = self.formats[fmt]
            if value is None or value == '':
                continue
            ref = f'{get_column_letter(index)}{number}'
            prefix = f'<c r="{ref}" s="{style}"'
            if isinstance(value, str):
                if len(value) > 32767:
                    raise ValueError('Một ô vượt giới hạn 32.767 ký tự của Excel.')
                # XML normalizes literal CR/CRLF unless CR is a character ref.
                parts.append(prefix + ' t="inlineStr"><is><t xml:space="preserve">' + escape(value).replace('\r','&#13;') + '</t></is></c>')
            elif isinstance(value, bool):
                parts.append(prefix + f' t="b"><v>{int(value)}</v></c>')
            else:
                if isinstance(value, (datetime, date)):
                    value = to_excel(value)
                if not isinstance(value, (int, float)) or not math.isfinite(value):
                    raise ValueError('Giá trị ô không hợp lệ khi xuất Excel.')
                parts.append(prefix + f'><v>{value!r}</v></c>')
        parts.append('</row>')
        self._write(''.join(parts))

    def end_sheet(self):
        if not self.current:
            return
        item = self.current
        self._write('</sheetData>')
        if item['headers'] and item['rows']:
            self._write(f'<autoFilter ref="A1:{get_column_letter(len(item["headers"]))}{item["rows"]}"/>')
        self._write('</worksheet>')
        item['stream'].write(self.buffer)
        self.buffer.clear()
        item['stream'].close()
        self.current = None

    def close(self):
        self.end_sheet()
        types = ['<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
        book = [f'<workbook xmlns="{NS}" xmlns:r="{REL}"><sheets>']
        rels = ['<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        for index, sheet in enumerate(self.sheets, 1):
            types.append(f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            book.append(f'<sheet name={quoteattr(sheet["name"])} sheetId="{index}" r:id="rId{index}"/>')
            rels.append(f'<Relationship Id="rId{index}" Type="{REL}/worksheet" Target="worksheets/sheet{index}.xml"/>')
        types.append('</Types>'); book.append('</sheets></workbook>')
        rels.append(f'<Relationship Id="styles" Type="{REL}/styles" Target="styles.xml"/></Relationships>')
        styles = [f'<styleSheet xmlns="{NS}"><numFmts count="{len(self.formats)}">']
        for index, fmt in enumerate(self.formats, 164):
            styles.append(f'<numFmt numFmtId="{index}" formatCode={quoteattr(fmt)}/>')
        styles.append('</numFmts><fonts count="2"><font><sz val="10"/><color rgb="FF202020"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF263D4A"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
        styles.append(f'<cellXfs count="{2+len(self.formats)}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>')
        for index, _ in enumerate(self.formats, 164):
            styles.append(f'<xf numFmtId="{index}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center"/></xf>')
        styles.append('</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
        self.archive.writestr('[Content_Types].xml', ''.join(types))
        self.archive.writestr('xl/workbook.xml', ''.join(book))
        self.archive.writestr('xl/_rels/workbook.xml.rels', ''.join(rels))
        self.archive.writestr('xl/styles.xml', ''.join(styles))
        self.archive.writestr('_rels/.rels', f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        self.archive.close()

    def abort(self):
        if self.current:
            self.current['stream'].close()
            self.current = None
        self.archive.close()


class Source:
    def __init__(self, path, max_unpacked=MAX_UNPACKED):
        self.archive = zipfile.ZipFile(path)
        try:
            if max_unpacked is not None and sum(info.file_size for info in self.archive.infolist()) > max_unpacked:
                raise ValueError('Nội dung Excel sau giải nén vượt giới hạn 8 GB/file.')
            parser = ET.XMLParser(resolve_entities=False, no_network=True)
            book = ET.fromstring(self.archive.read('xl/workbook.xml'), parser)
            relationships = ET.fromstring(self.archive.read('xl/_rels/workbook.xml.rels'), parser)
            targets = {r.get('Id'): r.get('Target') for r in relationships}
            self.sheets = []
            for sheet in book.find(TAG + 'sheets'):
                target = targets[sheet.get('{' + REL + '}id')]
                target = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
                rows, columns = 0, 0
                with self.archive.open(target) as stream:
                    for _, element in ET.iterparse(stream, events=('start',), resolve_entities=False, no_network=True):
                        if element.tag == TAG + 'dimension':
                            last = element.get('ref', 'A1').split(':')[-1]
                            letters = ''.join(c for c in last if c.isalpha())
                            rows = int(''.join(c for c in last if c.isdigit()) or 0)
                            columns = column_index_from_string(letters or 'A')
                        if element.tag == TAG + 'sheetData':
                            break
                self.sheets.append(dict(name=sheet.get('name'), path=target, rows=rows, columns=columns))
            prop = book.find(TAG + 'workbookPr')
            self.epoch = CALENDAR_MAC_1904 if prop is not None and prop.get('date1904') in ('1','true') else CALENDAR_WINDOWS_1900
            self.formats = ['General']
            if 'xl/styles.xml' in self.archive.namelist():
                styles = ET.fromstring(self.archive.read('xl/styles.xml'), parser)
                custom = {int(f.get('numFmtId')): f.get('formatCode') for f in styles.findall(TAG+'numFmts/'+TAG+'numFmt')}
                self.formats = [custom.get(int(x.get('numFmtId', 0)), BUILTIN_FORMATS.get(int(x.get('numFmtId', 0)), 'General')) for x in styles.findall(TAG+'cellXfs/'+TAG+'xf')] or ['General']
            self.date_formats = {fmt for fmt in self.formats if is_date_format(fmt)}
            self.strings = []
            if 'xl/sharedStrings.xml' in self.archive.namelist():
                with self.archive.open('xl/sharedStrings.xml') as stream:
                    for _, item in ET.iterparse(stream, events=('end',), tag=TAG+'si', resolve_entities=False, no_network=True):
                        self.strings.append(''.join(item.itertext()) if not item.findall('.//'+TAG+'t') else ''.join(t.text or '' for t in item.findall('.//'+TAG+'t')))
                        item.clear()
                        while item.getprevious() is not None:
                            del item.getparent()[0]
        except Exception:
            self.archive.close()
            raise

    def close(self):
        self.archive.close()

    def rows(self, name, end=None):
        sheet = next((s for s in self.sheets if s['name'] == name), None)
        if sheet is None:
            raise ValueError('Trang tính không còn tồn tại.')
        previous = 0
        with self.archive.open(sheet['path']) as stream:
            for _, row in ET.iterparse(stream, events=('end',), tag=TAG+'row', resolve_entities=False, no_network=True, huge_tree=True):
                number = int(row.get('r', previous + 1))
                if end is not None and number > end:
                    break
                while previous + 1 < number:
                    previous += 1
                    yield [], []
                values, formats = [], []
                for cell in row:
                    if cell.tag != TAG+'c':
                        continue
                    ref = cell.get('r', '')
                    letters = ref.rstrip('0123456789')
                    index = column_index(letters) if letters else len(values)
                    if index >= 16384:
                        raise ValueError('Trang tính vượt giới hạn cột Excel.')
                    while len(values) <= index:
                        values.append(None); formats.append('General')
                    style = int(cell.get('s', 0))
                    fmt = self.formats[style] if style < len(self.formats) else 'General'
                    kind = cell.get('t')
                    value,formula=None,False
                    for child in cell:
                        if child.tag==V_TAG:
                            value=child.text
                        elif child.tag==F_TAG:
                            formula=True
                    if formula and value is None:
                        raise ValueError(f'{name}!{ref}: công thức chưa có kết quả đã lưu. Mở Excel, tính lại và lưu file.')
                    if kind == 'e':
                        raise ValueError(f'{name}!{ref}: ô lỗi Excel.')
                    if kind == 's':
                        value = self.strings[int(value)] if value is not None else None
                    elif kind == 'inlineStr':
                        value = ''.join(t.text or '' for t in cell.iter(T_TAG))
                    elif kind == 'b':
                        value = value == '1'
                    elif kind == 'd' and value:
                        value = datetime.fromisoformat(value)
                    elif kind not in ('str', 'inlineStr') and value is not None:
                        value = float(value) if '.' in value or 'e' in value.lower() else int(value)
                        if fmt in self.date_formats:
                            value = from_excel(value, self.epoch)
                    values[index], formats[index] = value, fmt
                previous = number
                yield values, formats
                row.clear()
                while row.getprevious() is not None:
                    del row.getparent()[0]
        limit = min(end, sheet['rows']) if end is not None and sheet['rows'] else sheet['rows']
        while previous < limit:
            previous += 1
            yield [], []
