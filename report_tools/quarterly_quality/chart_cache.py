"""Keep native QMS chart caches identical to their embedded worksheet values."""

from copy import copy
from io import BytesIO
from pathlib import Path
import posixpath
import re
from zipfile import ZipFile

from lxml import etree
from openpyxl.utils.cell import get_column_letter, range_boundaries


NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
}


def _part(owner, target):
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target)).lstrip('/')


def _rels_name(owner):
    return posixpath.dirname(owner) + '/_rels/' + posixpath.basename(owner) + '.rels'


def _rels(package, owner):
    name = _rels_name(owner)
    return {rel.get('Id'): rel for rel in etree.fromstring(package.read(name))}


def _cells(data):
    """Read the original numeric text; parsing floats loses Excel's precision."""
    with ZipFile(BytesIO(data)) as book:
        strings = []
        if 'xl/sharedStrings.xml' in book.namelist():
            strings = [''.join(si.xpath('.//x:t/text()', namespaces=NS))
                       for si in etree.fromstring(book.read('xl/sharedStrings.xml'))]
        relationships = _rels(book, 'xl/workbook.xml')
        sheets = {}
        for sheet in etree.fromstring(book.read('xl/workbook.xml')).findall('x:sheets/x:sheet', NS):
            relation = relationships[sheet.get('{'+NS['r']+'}id')]
            if relation.get('TargetMode') == 'External':
                raise ValueError('Embedded chart worksheet must be internal.')
            name = _part('xl/workbook.xml', relation.get('Target'))
            values = {}
            for cell in etree.fromstring(book.read(name)).findall('x:sheetData/x:row/x:c', NS):
                value = cell.find('x:v', NS)
                text = value.text if value is not None else None
                kind = cell.get('t', 'n')
                if kind == 's' and text is not None:
                    text = strings[int(text)]
                elif kind == 'inlineStr':
                    text = ''.join(cell.xpath('.//x:t/text()', namespaces=NS))
                elif kind in ('e', 'b'):
                    text = None
                values[cell.get('r')] = text
            sheets[sheet.get('name')] = values
        return sheets


def _reference(sheets, formula):
    owner, address = formula.rsplit('!', 1)
    owner = owner.strip("'").replace("''", "'")
    if any(char in owner for char in '[]:/\\'):
        raise ValueError('Chart data must refer to its embedded workbook.')
    x1, y1, x2, y2 = range_boundaries(address)
    if x1 != x2 and y1 != y2:
        raise ValueError('Chart data range must be a single row or column.')
    return [sheets[owner].get(f'{get_column_letter(x)}{y}')
            for y in range(y1, y2 + 1) for x in range(x1, x2 + 1)]


def _display_label_sheet(data, sheet_name, values):
    """Retain WPS's wrapped labels alongside the unchanged source worksheets."""
    with ZipFile(BytesIO(data)) as book:
        workbook = etree.fromstring(book.read('xl/workbook.xml'))
        relationships = etree.fromstring(book.read('xl/_rels/workbook.xml.rels'))
        types = etree.fromstring(book.read('[Content_Types].xml'))
        sheets = workbook.find('x:sheets', NS)
        ident = max(int(sheet.get('sheetId')) for sheet in sheets) + 1
        rid = 'rIdQMSLabels'
        while any(rel.get('Id') == rid for rel in relationships):
            rid += 'x'
        target = f'worksheets/qms_labels_{ident}.xml'
        etree.SubElement(sheets, '{'+NS['x']+'}sheet', name=sheet_name, sheetId=str(ident),
                         state='hidden', attrib={'{'+NS['r']+'}id': rid})
        etree.SubElement(relationships, '{'+etree.QName(relationships).namespace+'}Relationship',
                         Id=rid, Type=NS['r']+'/worksheet', Target=target)
        etree.SubElement(types, '{'+etree.QName(types).namespace+'}Override',
                         PartName='/xl/'+target,
                         ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        worksheet = etree.Element('{'+NS['x']+'}worksheet', nsmap={None:NS['x']})
        rows = etree.SubElement(worksheet, '{'+NS['x']+'}sheetData')
        for row in range(1, 5):
            record = etree.SubElement(rows, '{'+NS['x']+'}row', r=str(row))
            for address, text in values.items():
                if int(re.search(r'\d+$', address)[0]) != row:
                    continue
                cell = etree.SubElement(record, '{'+NS['x']+'}c', r=address, t='inlineStr')
                inline = etree.SubElement(cell, '{'+NS['x']+'}is')
                etree.SubElement(inline, '{'+NS['x']+'}t').text = text
        updates = {'xl/workbook.xml':etree.tostring(workbook),
                   'xl/_rels/workbook.xml.rels':etree.tostring(relationships),
                   '[Content_Types].xml':etree.tostring(types),
                   'xl/'+target:etree.tostring(worksheet)}
        result = BytesIO()
        with ZipFile(result, 'w') as output:
            for info in book.infolist():
                output.writestr(copy(info), updates.get(info.filename, book.read(info.filename)))
            output.writestr('xl/'+target, updates['xl/'+target])
        return result.getvalue()


def refresh_qms_chart_caches(path, *, linked_workbook=None):
    """Refresh QMS charts only, preserving workbook formulas and slide design.

    Office rounds chart caches to 15 significant digits while its embedded XLSX
    can retain more digits. Copy cached worksheet text directly instead of
    recalculating or rounding either source. WPS can paste a linked chart; embed
    the known, saved source workbook before its temporary folder is removed.
    Never follow an external URL from the presentation itself. Return the number
    of refreshed charts.
    """
    updates = {}
    workbooks = {}
    changed_charts = 0
    linked_data = linked_name = None
    label_sheets = {}
    with ZipFile(path) as package:
        charts = set()
        for name in package.namelist():
            if not re.fullmatch(r'ppt/slides/slide\d+\.xml', name):
                continue
            slide = etree.fromstring(package.read(name))
            title = ' '.join(slide.xpath('.//a:t/text()', namespaces=NS))
            if not re.search(r'\bQMS\s+(?:FTT|HFPA)\s+ANALYSIS\b', title, re.I):
                continue
            for relation in _rels(package, name).values():
                if relation.get('Type', '').endswith('/chart') and relation.get('TargetMode') != 'External':
                    charts.add(_part(name, relation.get('Target')))
        for name in sorted(charts):
            chart = etree.fromstring(package.read(name))
            external = chart.find('c:externalData', NS)
            if external is None:
                raise ValueError(f'{name}: chart has no embedded workbook.')
            relation = _rels(package, name)[external.get('{'+NS['r']+'}id')]
            changed = False
            if relation.get('TargetMode') == 'External':
                if linked_workbook is None:
                    raise ValueError(f'{name}: linked chart requires its known saved source workbook.')
                if linked_data is None:
                    linked_data = Path(linked_workbook).read_bytes()
                    index = 1
                    linked_name = 'ppt/embeddings/quarterly_source.xlsx'
                    while linked_name in package.namelist():
                        index += 1
                        linked_name = f'ppt/embeddings/quarterly_source_{index}.xlsx'
                    updates[linked_name] = linked_data
                    content_types = etree.fromstring(package.read('[Content_Types].xml'))
                    ct = 'http://schemas.openxmlformats.org/package/2006/content-types'
                    if not any(e.get('Extension') == 'xlsx' for e in content_types):
                        etree.SubElement(content_types, '{'+ct+'}Default', Extension='xlsx',
                                         ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        updates['[Content_Types].xml'] = etree.tostring(content_types)
                relation.set('Type', NS['r']+'/package')
                relation.set('Target', posixpath.relpath(linked_name, posixpath.dirname(name)))
                relation.attrib.pop('TargetMode', None)
                updates[_rels_name(name)] = etree.tostring(relation.getparent())
                workbook_name = linked_name
                changed = True
            else:
                workbook_name = _part(name, relation.get('Target'))
            if workbook_name not in workbooks:
                workbooks[workbook_name] = _cells(linked_data if workbook_name == linked_name
                                                else package.read(workbook_name))
            for series in chart.findall('.//c:ser', NS):
                # WPS adds [source.xlsx] even when the exact saved source has
                # just been embedded. Point these references inside that book.
                if workbook_name == linked_name:
                    for formula in series.findall('.//c:f', NS):
                        if '!' not in (formula.text or ''):
                            continue
                        owner, address = formula.text.rsplit('!', 1)
                        owner = owner.strip("'").replace("''", "'")
                        prefix = '['+Path(linked_workbook).name+']'
                        if owner.casefold().startswith(prefix.casefold()):
                            owner = owner[len(prefix):]
                            if owner not in workbooks[workbook_name]:
                                raise ValueError(f'{name}: linked source worksheet is missing.')
                            formula.text = "'"+owner.replace("'", "''")+"'!"+address
                            changed = True
                # WPS also stores wrapped ranking labels as an array formula.
                # Link those names to the matching original ranking table.
                category = series.find('c:cat/c:strRef/c:f', NS)
                if category is not None and (category.text or '').startswith('{'):
                    value_formula = series.find('c:val/c:numRef/c:f', NS)
                    owner, address = value_formula.text.rsplit('!', 1)
                    sheet_name = owner.strip("'").replace("''", "'")
                    x1, y1, x2, y2 = range_boundaries(address)
                    label_column = {'FTT month-quarter': (7, 'D'), 'Quarter HFPA ': (17, 'N')}.get(sheet_name)
                    if label_column is None or x1 != x2 or x1 != label_column[0]:
                        raise ValueError(f'{name}: cannot locate the source ranking labels.')
                    formula = f'{owner}!${label_column[1]}${y1}:${label_column[1]}${y2}'
                    labels = _reference(workbooks[workbook_name], formula)
                    cached = category.getparent().xpath('c:strCache/c:pt/c:v/text()', namespaces=NS)
                    if [''.join(v.split()) for v in cached] != [''.join((v or '').split()) for v in labels]:
                        raise ValueError(f'{name}: ranking labels do not match the saved source table.')
                    if cached != labels:
                        if workbook_name not in label_sheets:
                            helper_name = 'QMS Chart Labels'
                            while helper_name in workbooks[workbook_name]:
                                helper_name += 'x'
                            helper_values = {}
                            workbooks[workbook_name][helper_name] = helper_values
                            label_sheets[workbook_name] = helper_name, helper_values
                        helper_name, helper_values = label_sheets[workbook_name]
                        column = get_column_letter(len(helper_values) // 4 + 1)
                        helper_values[column+'1'] = formula
                        helper_values.update({f'{column}{i+2}':value for i,value in enumerate(cached)})
                        category.text = f"'{helper_name}'!${column}$2:${column}$4"
                    else:
                        category.text = formula
                    changed = True
                title = series.find('c:tx', NS)
                title_ref = title.find('c:strRef', NS) if title is not None else None
                title_formula = title_ref.find('c:f', NS) if title_ref is not None else None
                # Office sometimes saves a literal legend label as a formula
                # such as VN or "VH". Only real worksheet references use strRef.
                if title_formula is not None and '!' not in (title_formula.text or ''):
                    formula = title_formula.text or ''
                    label = (formula[1:-1].replace('""', '"') if re.fullmatch(r'"(?:[^"]|"")*"', formula)
                             else formula if re.fullmatch(r'[A-Za-z][A-Za-z0-9 _.-]*', formula) else None)
                    cached = title_ref.xpath('c:strCache/c:pt/c:v/text()', namespaces=NS)
                    if label is None or cached != [label]:
                        raise ValueError(f'{name}: ambiguous literal series name.')
                    title.remove(title_ref)
                    etree.SubElement(title, '{'+NS['c']+'}v').text = label
                    changed = True
                for role in ('tx', 'cat', 'val'):
                    parent = series.find('c:'+role, NS)
                    if parent is None:
                        continue
                    for ref in parent:
                        kind = etree.QName(ref).localname
                        if kind not in ('numRef', 'strRef'):
                            continue
                        formula = ref.find('c:f', NS)
                        if formula is None or not formula.text:
                            raise ValueError(f'{name}: chart reference has no formula.')
                        values = _reference(workbooks[workbook_name], formula.text)
                        cache_tag = 'numCache' if kind == 'numRef' else 'strCache'
                        cache = ref.find('c:'+cache_tag, NS)
                        if cache is None:
                            cache = etree.SubElement(ref, '{'+NS['c']+'}'+cache_tag)
                        expected = [(str(index), value) for index, value in enumerate(values)
                                    if value not in (None, '')]
                        actual = [(pt.get('idx'), pt.findtext('c:v', namespaces=NS))
                                  for pt in cache.findall('c:pt', NS)]
                        count = cache.find('c:ptCount', NS)
                        if actual == expected and count is not None and count.get('val') == str(len(values)):
                            continue
                        changed = True
                        for point in cache.findall('c:pt', NS):
                            cache.remove(point)
                        if count is None:
                            count = etree.SubElement(cache, '{'+NS['c']+'}ptCount')
                        count.set('val', str(len(values)))
                        for index, value in expected:
                            point = etree.SubElement(cache, '{'+NS['c']+'}pt', idx=index)
                            etree.SubElement(point, '{'+NS['c']+'}v').text = value
            if changed:
                updates[name] = etree.tostring(chart)
                changed_charts += 1
        if updates:
            for workbook_name, (sheet_name, values) in label_sheets.items():
                original = linked_data if workbook_name == linked_name else package.read(workbook_name)
                updates[workbook_name] = _display_label_sheet(original, sheet_name, values)
            output = BytesIO()
            with ZipFile(output, 'w') as result:
                for info in package.infolist():
                    result.writestr(copy(info), updates.get(info.filename, package.read(info.filename)))
                for name, data in updates.items():
                    if name not in package.namelist():
                        result.writestr(name, data)
    if updates:
        Path(path).write_bytes(output.getvalue())
    return changed_charts
