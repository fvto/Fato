"""Local, configurable spreadsheet checks; never execute workbook formulas or code."""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from functools import lru_cache
from itertools import zip_longest
from pathlib import Path
from uuid import uuid4
import csv
import hashlib
import io
import json
import math
import re
import shutil
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string, range_boundaries, coordinate_to_tuple

ROOT = Path(__file__).resolve().parent / 'Report' / 'Comparisons'
EXPORTS = Path(__file__).resolve().parent / 'Report' / 'HFPAExports'
MAX_BYTES = 40 * 1024 * 1024
MAX_CELLS = 500_000
MAX_RESULTS = 50_000
MISSING = {'', 'n/a', 'na', 'n.a.', '-', '—'}
ERRORS = {'#DIV/0!', '#VALUE!', '#REF!', '#N/A', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'}
OPERATIONS = {'average', 'sum', 'min', 'max', 'count', 'weighted_average', 'ratio_of_sums'}


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    pending = path.with_suffix('.partial.json')
    pending.write_text(json.dumps(data, ensure_ascii=False, allow_nan=False), encoding='utf-8')
    pending.replace(path)


def identifier(value):
    if not re.fullmatch(r'[a-f0-9]{32}', str(value or '')):
        raise ValueError('Unknown workbook or result. Import the file again.')
    return str(value)


def display(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    return value


@lru_cache(maxsize=8)
def load_workbook(root, workbook_id):
    folder = Path(root) / 'inputs' / identifier(workbook_id)
    meta = json.loads((folder / 'metadata.json').read_text(encoding='utf-8'))
    path = folder / ('source' + meta['extension'])
    sheets, total = {}, 0
    if path.suffix == '.xls':
        import xlrd
        book = xlrd.open_workbook(str(path))
        try:
            for sheet in book.sheets():
                total += sheet.nrows * sheet.ncols
                if total > MAX_CELLS:
                    raise ValueError('Workbook exceeds 500,000 cells. Import a smaller workbook.')
                cells = {}
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        cell = sheet.cell(row, col)
                        if cell.ctype == xlrd.XL_CELL_EMPTY:
                            continue
                        value = cell.value
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            value = xlrd.xldate_as_datetime(value, book.datemode)
                        elif cell.ctype == xlrd.XL_CELL_ERROR:
                            value = xlrd.error_text_from_code.get(value, '#VALUE!')
                        cells[(row + 1, col + 1)] = {'value': display(value), 'formula': None, 'format': 'General'}
                sheets[sheet.name] = {'cells': cells, 'rows': sheet.nrows, 'columns': sheet.ncols}
        finally:
            book.release_resources()
    else:
        with zipfile.ZipFile(path) as archive:
            if sum(item.file_size for item in archive.infolist()) > 160 * 1024 * 1024:
                raise ValueError('Expanded workbook is too large. Import a smaller workbook.')
        formulas = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
        try:
            cached = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
            try:
                for sheet in formulas:
                    if (sheet.max_row or 0) > MAX_CELLS or (sheet.max_column or 0) > 16384:
                        raise ValueError('Sheet dimensions are too large. Remove unused formatted rows first.')
                    cells, last_row, last_col = {}, 0, 0
                    for row_index, (formula_row, cached_row) in enumerate(zip_longest(
                            sheet.iter_rows(), cached[sheet.title].iter_rows(), fillvalue=()), 1):
                        total += max(len(formula_row), len(cached_row))
                        if total > MAX_CELLS:
                            raise ValueError('Workbook exceeds 500,000 cells. Import a smaller workbook.')
                        for col_index, (cell, saved) in enumerate(zip_longest(formula_row, cached_row), 1):
                            if cell is None or cell.value is None:
                                continue
                            formula = str(cell.value) if cell.data_type == 'f' else None
                            value = saved.value if formula and saved is not None else cell.value
                            cells[(row_index, col_index)] = {
                                'value': display(value), 'formula': formula,
                                'format': cell.number_format or 'General'}
                            last_row, last_col = max(last_row, row_index), max(last_col, col_index)
                    sheets[sheet.title] = {'cells': cells, 'rows': last_row, 'columns': last_col}
            finally:
                cached.close()
        finally:
            formulas.close()
    return dict(meta, sheets=sheets)


def workbook(workbook_id):
    try:
        return load_workbook(str(ROOT.resolve()), identifier(workbook_id))
    except FileNotFoundError as exc:
        raise ValueError('Workbook is no longer available. Import it again.') from exc


def cell(sheet, row, col):
    return sheet['cells'].get((row, col), {'value': None, 'formula': None, 'format': 'General'})


def address(row, col):
    return f'{get_column_letter(col)}{row}'


def bounds(value):
    if not re.fullmatch(r'\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?', str(value or '')):
        raise ValueError('Enter an Excel cell range, for example C5:O55.')
    left, top, right, bottom = range_boundaries(value)
    if left > right or top > bottom or right > 16384 or bottom > 1048576:
        raise ValueError('Invalid Excel range.')
    if (right - left + 1) * (bottom - top + 1) > MAX_CELLS:
        raise ValueError('Selected range exceeds 500,000 cells.')
    return left, top, right, bottom


def source(spec):
    book = workbook(spec.get('workbook'))
    if spec.get('sheet') not in book['sheets']:
        raise ValueError('Choose a sheet from the selected workbook.')
    return book, book['sheets'][spec['sheet']]


def numeric(item, mode='number'):
    value = item['value']
    if item.get('formula') and value is None:
        return 'uncached', None
    if isinstance(value, str) and value.strip().upper() in ERRORS:
        return 'error', None
    if value is None or (isinstance(value, str) and value.strip().lower() in MISSING):
        return 'missing', None
    if isinstance(value, bool):
        return 'invalid', None
    percent = isinstance(value, str) and value.strip().endswith('%')
    text = str(value).strip().replace('\u00a0', '').replace(' ', '')
    if percent:
        text = text[:-1]
    if re.fullmatch(r'[+-]?\d{1,3}(,\d{3})+(\.\d+)?', text):
        text = text.replace(',', '')
    elif ',' in text and '.' not in text and text.count(',') == 1:
        text = text.replace(',', '.')
    if not re.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?', text):
        return 'invalid', None
    try:
        number = Decimal(text)
    except InvalidOperation:
        return 'invalid', None
    if not number.is_finite():
        return 'invalid', None
    if percent:
        number /= 100
    elif mode == 'percentage_points' and '%' not in item.get('format', ''):
        number /= 100
    if mode == 'percent' and not percent and '%' not in item.get('format', '') and abs(number) > 1:
        return 'invalid', None
    return 'number', number


def settings(rule):
    mode = rule.get('number_mode', 'number')
    if mode not in {'number', 'percent', 'percentage_points'}:
        raise ValueError('Unknown number interpretation.')
    try:
        tolerance = Decimal(str(rule.get('tolerance', '0')))
    except InvalidOperation as exc:
        raise ValueError('Tolerance must be a nonnegative number.') from exc
    if not tolerance.is_finite() or tolerance < 0:
        raise ValueError('Tolerance must be a nonnegative number.')
    if tolerance != 0:
        raise ValueError('Sai số phải bằng 0. Không cho phép khoảng sai số.')
    if mode != 'number':
        tolerance /= 100
    digits = rule.get('comparison_rounding_digits')
    if digits not in (None, ''):
        if isinstance(digits, bool) or not re.fullmatch(r'[0-9]', str(digits)):
            raise ValueError('Số chữ số làm tròn trước khi so sánh phải từ 0 đến 9.')
        digits = int(digits)
    else:
        digits = None
    return mode, tolerance, digits


def decimals(item):
    text = str(item.get('value') or '').strip()
    if text.endswith('%'):
        text = text[:-1].replace(',', '.')
        return len(text.split('.', 1)[1]) if '.' in text else 0
    fmt = item.get('format', '').split(';')[0]
    if '%' in fmt:
        match = re.search(r'\.([0#?]+)', fmt)
        return len(match[1]) if match else 0
    return None


def round_for_comparison(number,mode,digits):
    if number is None or digits is None: return number
    quantum=Decimal(1).scaleb(-(digits+(2 if mode!='number' else 0)))
    with localcontext() as context:
        context.prec=max(28,len(number.as_tuple().digits)+digits+6,number.adjusted()+digits+6)
        return number.quantize(quantum,rounding=ROUND_HALF_UP)


def compare_numbers(expected, actual, item, mode, tolerance, digits):
    expected=round_for_comparison(expected,mode,digits)
    actual=round_for_comparison(actual,mode,digits)
    delta = actual - expected
    if delta == 0:
        return 'match', delta, 'Values match.' if digits is None else f'Values match after rounding to {digits} decimals; tolerance 0.'
    return 'difference', delta, 'Values differ. Tolerance 0.' if digits is None else f'Values differ after rounding to {digits} decimals; tolerance 0.'


def finding(rule, status, book, sheet, location, **extra):
    return {'rule': rule.get('name') or rule['kind'], 'status': status, 'file': book['name'],
            'workbook': book['id'], 'sheet': sheet, 'cell': location, **extra}


def calculation(rule):
    book, sheet = source(rule['left'])
    mode, tolerance, digits = settings(rule)
    operation = rule.get('operation', 'average')
    if operation not in OPERATIONS:
        raise ValueError('Choose a supported calculation.')
    if rule.get('missing_policy', 'skip') not in {'skip', 'error'}:
        raise ValueError('Choose how missing inputs should be handled.')
    direction = rule.get('direction', 'rows')
    if direction not in {'rows', 'columns'}:
        raise ValueError('Choose rows or columns.')
    left, top, right, bottom = bounds(rule.get('source_range'))
    tl, tt, tr, tb = bounds(rule.get('target_range'))
    if not (tr < left or tl > right or tb < top or tt > bottom):
        raise ValueError('Expected-value cells must be outside the input range.')
    groups = [[(row, col) for col in range(left, right + 1)] for row in range(top, bottom + 1)] if direction == 'rows' else [
        [(row, col) for row in range(top, bottom + 1)] for col in range(left, right + 1)]
    targets = [(row, col) for row in range(tt, tb + 1) for col in range(tl, tr + 1)]
    if len(groups) != len(targets) or (direction == 'rows' and tl != tr) or (direction == 'columns' and tt != tb):
        raise ValueError('Use one expected-value cell for each input row or column.')
    if len(groups) > MAX_RESULTS:
        raise ValueError('Calculation exceeds 50,000 checks. Narrow the ranges.')
    labels = []
    if rule.get('label_range'):
        ll, lt, lr, lb = bounds(rule['label_range'])
        labels = [cell(sheet, row, col)['value'] for row in range(lt, lb + 1) for col in range(ll, lr + 1)]
        if len(labels) != len(groups):
            raise ValueError('Label range must contain one cell per check.')
    weight_groups = None
    if operation in {'weighted_average', 'ratio_of_sums'}:
        wl, wt, wr, wb = bounds(rule.get('weights_range'))
        if wr - wl != right - left or wb - wt != bottom - top:
            raise ValueError('Weights / denominator range must have the same dimensions as the input range.')
        if not (tr < wl or tl > wr or tb < wt or tt > wb):
            raise ValueError('Expected-value cells must be outside the weights / denominator range.')
        weight_groups = [[(r + wt - top, c + wl - left) for r, c in group] for group in groups]
    results = []
    for index, (group, (row, col)) in enumerate(zip(groups, targets)):
        target = cell(sheet, row, col)
        state, actual = numeric(target, mode)
        numbers, weights, ignored, invalid, inputs = [], [], 0, [], []
        for position, (r, c) in enumerate(group):
            current = cell(sheet, r, c)
            kind, number = numeric(current, mode)
            inputs.append({'cell': address(r, c), 'value': current['value'], 'formula': current.get('formula')})
            if kind == 'missing' or (rule.get('ignore_zero') is True and kind == 'number' and number == 0):
                ignored += 1
                if rule.get('missing_policy', 'skip') == 'error' and kind == 'missing':
                    invalid.append(f'{address(r, c)}: missing value')
                continue
            if kind != 'number':
                invalid.append(f'{address(r, c)}: {kind} value {current["value"]!r}')
                continue
            if weight_groups is not None:
                weight_row, weight_col = weight_groups[index][position]
                weight_cell = cell(sheet, weight_row, weight_col)
                inputs.append({'cell': address(weight_row, weight_col), 'value': weight_cell['value'], 'role': 'weight / denominator', 'formula': weight_cell.get('formula')})
                weight_kind, weight = numeric(weight_cell, 'number')
                if weight_kind != 'number':
                    invalid.append(f'{address(weight_row, weight_col)}: missing or invalid weight / denominator')
                    continue
                weights.append(weight)
            numbers.append(number)
        base = dict(label=str(labels[index] or '') if labels else address(row, col), actual=display(target['value']),
                    formula=target.get('formula'), valid_count=len(numbers), ignored_count=ignored,
                    operation=operation, inputs=inputs, expected=None, delta=None)
        status, reason, expected = None, '', None
        if invalid:
            status, reason = 'invalid_input', '; '.join(invalid)
        elif numbers or operation == 'count':
            if operation == 'average': expected = sum(numbers) / len(numbers)
            elif operation == 'sum': expected = sum(numbers)
            elif operation == 'min': expected = min(numbers)
            elif operation == 'max': expected = max(numbers)
            elif operation == 'count': expected = Decimal(len(numbers))
            elif sum(weights) == 0: status, reason = 'no_data', 'The total weight / denominator is zero.'
            elif operation == 'weighted_average': expected = sum(value * weight for value, weight in zip(numbers, weights)) / sum(weights)
            else: expected = sum(numbers) / sum(weights)
        else:
            status, reason = 'no_data', 'No valid numeric input. Missing values are not treated as zero.'
        if state in {'error', 'uncached'}:
            status = 'formula_error' if state == 'error' else 'formula_uncached'
            reason = 'Expected-value cell contains an Excel error.' if state == 'error' else 'Formula has no saved result. Recalculate and save it in Excel / WPS before checking.'
        elif status is None and expected is not None:
            if state == 'missing': status, reason = 'missing_value', 'Numeric inputs exist, but the expected-value cell is blank or N/A.'
            elif state != 'number': status, reason = 'invalid_value', 'Expected-value cell is not a valid number in the chosen number interpretation.'
            else:
                status, delta, reason = compare_numbers(expected, actual, target, mode, tolerance, digits)
                base['delta'] = float(delta)
                base['delta_exact'] = str(delta)
        elif status == 'no_data' and state == 'number':
            status, reason = 'missing_input', 'A numeric result is present, but there are no usable inputs / denominator.'
        if expected is not None:
            rounded_expected=round_for_comparison(expected,mode,digits)
            rounded_actual=round_for_comparison(actual,mode,digits) if state=='number' else None
            base['expected'] = float(rounded_expected)
            base['expected_exact'] = str(rounded_expected)
            base['actual_exact'] = str(rounded_actual) if rounded_actual is not None else None
            base['expected_unrounded']=str(expected)
            base['actual_unrounded']=str(actual) if state=='number' else None
            if digits is not None and rounded_actual is not None: base['actual']=float(rounded_actual)
            method = {'average': 'Average', 'sum': 'Sum', 'min': 'Minimum', 'max': 'Maximum', 'count': 'Numeric count',
                      'weighted_average': 'Weighted average', 'ratio_of_sums': 'Ratio of sums'}[operation]
            reason = f'{method}: {len(numbers)} numeric inputs, {ignored} skipped. ' + reason
        results.append(finding(rule, status, book, rule['left']['sheet'], address(row, col), reason=reason, **base))
    return results


def columns(value):
    parts = [part.strip().upper() for part in str(value or '').split(',') if part.strip()]
    if not parts or any(not re.fullmatch(r'[A-Z]{1,3}', part) for part in parts):
        raise ValueError('Enter column letters separated by commas, for example B,D.')
    indices = [column_index_from_string(part) for part in parts]
    if len(set(indices)) != len(indices) or max(indices) > 16384:
        raise ValueError('Column list contains duplicates or invalid columns.')
    return indices


def comparison(rule):
    left_book, left_sheet = source(rule['left'])
    right_book, right_sheet = source(rule['right'])
    mode, tolerance, digits = settings(rule)
    left_bounds, right_bounds = bounds(rule.get('left_range')), bounds(rule.get('right_range'))
    left_columns, right_columns = columns(rule.get('left_columns')), columns(rule.get('right_columns'))
    if len(left_columns) != len(right_columns):
        raise ValueError('Choose the same number of value columns on both sides.')
    if ((left_bounds[3] - left_bounds[1] + 1) + (right_bounds[3] - right_bounds[1] + 1)) * len(left_columns) > MAX_RESULTS:
        raise ValueError('Comparison exceeds 50,000 checks. Narrow the ranges or columns.')
    matching = rule.get('matching', 'keys')
    if matching not in {'keys', 'position'} or rule.get('value_type', 'auto') not in {'auto', 'numeric', 'text'}:
        raise ValueError('Unknown comparison mode.')
    key_left = columns(rule.get('left_keys')) if matching == 'keys' else []
    key_right = columns(rule.get('right_keys')) if matching == 'keys' else []
    if len(key_left) != len(key_right):
        raise ValueError('Choose the same number of key columns on both sides.')
    for region, selected in [(left_bounds, left_columns + key_left), (right_bounds, right_columns + key_right)]:
        if any(col < region[0] or col > region[2] for col in selected):
            raise ValueError('Key and value columns must be inside the selected data range.')
    def key_value(value):
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            text = format(Decimal(str(value)).normalize(), 'f')
        else:
            text = str(value or '').strip() if value != 0 else '0'
        return text.casefold() if rule.get('ignore_case', True) else text
    def keyed(sheet, region, keys):
        mapping = defaultdict(list)
        for row in range(region[1], region[3] + 1):
            if all(cell(sheet, row, col)['value'] in (None, '') and not cell(sheet, row, col).get('formula') for col in range(region[0], region[2] + 1)):
                continue
            if matching == 'position':
                key = (str(row - region[1] + 1),)
            else:
                key = tuple(key_value(cell(sheet, row, col)['value']) for col in keys)
                if any(numeric(cell(sheet, row, col))[0] in {'uncached', 'error'} for col in keys):
                    key = ('',)
            mapping[key].append(row)
        return mapping
    left_map = keyed(left_sheet, left_bounds, key_left)
    right_map = keyed(right_sheet, right_bounds, key_right)
    results = []
    for key in dict.fromkeys([*left_map, *right_map]):
        left_rows, right_rows = left_map.get(key, []), right_map.get(key, [])
        label = ' | '.join(key)
        provenance = {'other_file': right_book['name'], 'other_sheet': rule['right']['sheet'], 'other_cell': ''}
        if matching == 'keys' and any(not part for part in key):
            results.append(finding(rule, 'invalid_key', left_book, rule['left']['sheet'], '', label=label,
                reason='Blank / uncached matching key; rows were not paired.', left_rows=left_rows, right_rows=right_rows, **provenance))
            continue
        if len(left_rows) > 1 or len(right_rows) > 1:
            results.append(finding(rule, 'ambiguous_key', left_book, rule['left']['sheet'], '', label=label,
                reason='Duplicate key. Add another key column; no first-match or many-to-many pairing was used.', left_rows=left_rows, right_rows=right_rows, **provenance))
            continue
        if not left_rows or not right_rows:
            results.append(finding(rule, 'unmatched', left_book, rule['left']['sheet'], '', label=label,
                reason='Row exists only in ' + ('reference' if left_rows else 'comparison') + '.', left_rows=left_rows, right_rows=right_rows, **provenance))
            continue
        for left_col, right_col in zip(left_columns, right_columns):
            left_item, right_item = cell(left_sheet, left_rows[0], left_col), cell(right_sheet, right_rows[0], right_col)
            left_kind, left_number = numeric(left_item, mode)
            right_kind, right_number = numeric(right_item, mode)
            delta = None
            if 'uncached' in (left_kind, right_kind): status, reason = 'formula_uncached', 'A formula has no saved result; recalculate and save the workbook first.'
            elif 'error' in (left_kind, right_kind): status, reason = 'formula_error', 'One of the cells contains an Excel error.'
            elif left_kind == right_kind == 'missing': status, reason = 'no_data', 'Both cells are blank / N/A.'
            elif 'missing' in (left_kind, right_kind): status, reason = 'missing_value', 'Only one side has a value.'
            elif rule.get('value_type') == 'text' or (rule.get('value_type', 'auto') == 'auto' and 'number' not in (left_kind, right_kind)):
                status = 'match' if key_value(left_item['value']) == key_value(right_item['value']) else 'difference'
                reason = 'Text values match.' if status == 'match' else 'Text values differ.'
            elif left_kind == right_kind == 'number':
                status, delta, reason = compare_numbers(left_number, right_number, right_item, mode, tolerance, digits)
            else: status, reason = 'invalid_value', 'Numeric value cannot be interpreted consistently. Check the number interpretation.'
            results.append(finding(rule, status, left_book, rule['left']['sheet'], address(left_rows[0], left_col),
                label=label, actual=right_item['value'], expected=left_item['value'], delta=float(delta) if delta is not None else None,
                other_file=right_book['name'], other_sheet=rule['right']['sheet'], other_cell=address(right_rows[0], right_col), reason=reason,
                formula=left_item.get('formula'), other_formula=right_item.get('formula')))
            if left_kind == right_kind == 'number':
                rounded_left=round_for_comparison(left_number,mode,digits) if rule.get('value_type')!='text' else left_number
                rounded_right=round_for_comparison(right_number,mode,digits) if rule.get('value_type')!='text' else right_number
                results[-1].update(expected_exact=str(rounded_left), actual_exact=str(rounded_right),
                                   expected_unrounded=str(left_number),actual_unrounded=str(right_number),
                                   delta_exact=str(delta) if delta is not None else None)
                if digits is not None and rule.get('value_type')!='text':
                    results[-1].update(expected=float(rounded_left),actual=float(rounded_right))
    return results


def suggestions(book_id, name, sheet):
    """Suggest visible headers/ranges; the user reviews the mathematical rule."""
    cells, max_row, max_col = sheet['cells'], sheet['rows'], sheet['columns']
    header_row = max(range(1, min(20, max_row) + 1), key=lambda row: sum(
        isinstance(cell(sheet, row, col)['value'], str) and cell(sheet, row, col)['value'].strip().lower() not in MISSING
        for col in range(1, min(max_col, 100) + 1)), default=1)
    candidates = []
    operation_names = {'average': 'average', 'avg': 'average', 'mean': 'average', 'total': 'sum', 'sum': 'sum', 'minimum': 'min', 'maximum': 'max', 'count': 'count'}
    totals = [row for row in range(header_row + 1, max_row + 1) if any(str(cell(sheet, row, col)['value'] or '').strip().lower() in {'grand total', 'total'} for col in range(1, min(3, max_col) + 1))]
    last_data = min(totals) - 1 if totals else max_row
    for col in range(2, max_col + 1):
        label = str(cell(sheet, header_row, col)['value'] or '').strip().lower()
        if label not in operation_names or last_data <= header_row:
            continue
        numeric_cols = [c for c in range(1, col) if sum(numeric(cell(sheet, row, c))[0] == 'number' for row in range(header_row + 1, min(last_data, header_row + 12) + 1))]
        if not numeric_cols:
            continue
        first = min(numeric_cols)
        candidates.append({'kind': 'calculation', 'name': f'Check {get_column_letter(col)} · {label}',
            'left': {'workbook': book_id, 'sheet': name}, 'operation': operation_names[label], 'direction': 'rows',
            'source_range': f'{get_column_letter(first)}{header_row + 1}:{get_column_letter(col - 1)}{last_data}',
            'target_range': f'{get_column_letter(col)}{header_row + 1}:{get_column_letter(col)}{last_data}',
            'label_range': f'{get_column_letter(max(1, first - 1))}{header_row + 1}:{get_column_letter(max(1, first - 1))}{last_data}',
            'number_mode': 'percent' if any('%' in str(cell(sheet, header_row + 1, c)['value']) or '%' in cell(sheet, header_row + 1, c)['format'] for c in numeric_cols) else 'number',
            'tolerance': '0', 'rounding_digits': '', 'missing_policy': 'skip', 'ignore_zero': False})
    if totals and last_data > header_row:
        footer = min(totals)
        total_cols = [col for col in range(2, max_col + 1) if numeric(cell(sheet, footer, col))[0] == 'number']
        if total_cols:
            first, last = min(total_cols), max(total_cols)
            candidates.append({'kind': 'calculation', 'name': 'Footer totals · review Average / Sum method',
                'left': {'workbook': book_id, 'sheet': name}, 'operation': 'average', 'direction': 'columns',
                'source_range': f'{get_column_letter(first)}{header_row + 1}:{get_column_letter(last)}{last_data}',
                'target_range': f'{get_column_letter(first)}{footer}:{get_column_letter(last)}{footer}',
                'label_range': f'{get_column_letter(first)}{header_row}:{get_column_letter(last)}{header_row}',
                'number_mode': 'percent' if any('%' in str(cell(sheet, footer, col)['value']) or '%' in cell(sheet, footer, col)['format'] for col in total_cols) else 'number',
                'tolerance': '0', 'rounding_digits': '', 'missing_policy': 'skip', 'ignore_zero': False})
    return header_row, candidates


def describe(book):
    sheets = []
    for name, sheet in book['sheets'].items():
        header_row, proposed = suggestions(book['id'], name, sheet)
        sheets.append({'name': name, 'rows': sheet['rows'], 'columns': sheet['columns'], 'header_row': header_row,
            'data_range': f'A{header_row + 1}:{get_column_letter(max(1, sheet["columns"]))}{max(header_row + 1, sheet["rows"])}',
            'preview': [[{'cell': address(row, col), **cell(sheet, row, col)} for col in range(1, min(sheet['columns'], 20) + 1)] for row in range(1, min(sheet['rows'], 12) + 1)],
            'suggestions': proposed})
    return {'id': book['id'], 'name': book['name'], 'sheets': sheets}


def import_file(stream, name):
    extension = Path(name).suffix.lower()
    if extension not in {'.xlsx', '.xlsm', '.xls'}:
        raise ValueError('Choose an .xlsx, .xlsm or .xls workbook.')
    workbook_id = uuid4().hex
    folder = ROOT / 'inputs' / workbook_id
    folder.mkdir(parents=True)
    path = folder / ('source' + extension)
    try:
        with path.open('wb') as output:
            remaining = MAX_BYTES + 1
            while remaining:
                chunk = stream.read(min(1024 * 1024, remaining))
                if not chunk:
                    break
                output.write(chunk)
                remaining -= len(chunk)
        if path.stat().st_size > MAX_BYTES:
            raise ValueError('Maximum workbook size is 40 MB.')
        meta = {'id': workbook_id, 'name': str(name).replace('\\', '/').rsplit('/', 1)[-1], 'extension': extension,
                'sha256': hashlib.sha256(path.read_bytes()).hexdigest()}
        write_json(folder / 'metadata.json', meta)
        return describe(workbook(workbook_id))
    except Exception:
        load_workbook.cache_clear()
        if folder.resolve().is_relative_to((ROOT / 'inputs').resolve()):
            shutil.rmtree(folder)
        raise


def catalog(limit=80):
    files = []
    paths = sorted((p for p in EXPORTS.glob('*/*') if p.is_file() and p.suffix.lower() in {'.xls', '.xlsx', '.xlsm'} and '.partial' not in p.name), key=lambda p: p.stat().st_mtime, reverse=True)
    if limit is not None:
        paths = paths[:limit]
    for path in paths:
        relative = path.relative_to(EXPORTS).as_posix()
        key = hashlib.sha256(relative.encode()).hexdigest()[:32]
        files.append({'id': key, 'name': path.name, 'relative': relative, 'size': path.stat().st_size})
    return files


def delete_export(catalog_id):
    selection = next((item for item in catalog() if item['id'] == catalog_id), None)
    if selection is None:
        raise ValueError('File export không còn trong danh sách. Hãy tải lại trang.')
    path = (EXPORTS / selection['relative']).resolve()
    if not path.is_relative_to(EXPORTS.resolve()):
        raise ValueError('Export path is outside the report folder.')
    try:
        path.unlink()
    except FileNotFoundError as error:
        raise ValueError('File export đã được xóa. Hãy tải lại danh sách.') from error
    except OSError as error:
        raise ValueError('Không thể xóa file. Hãy đóng file trong Excel / WPS rồi thử lại.') from error
    return {'id': selection['id'], 'name': selection['name']}


def delete_exports(catalog_ids):
    if not isinstance(catalog_ids, list) or not catalog_ids or any(
            not isinstance(key, str) or not re.fullmatch(r'[a-f0-9]{32}', key) for key in catalog_ids):
        raise ValueError('Hãy tải lại danh sách trước khi xóa tất cả file export.')
    requested = set(catalog_ids)
    selections = [item for item in catalog(limit=None) if item['id'] in requested]
    # Validate every target before removing any files. Never remove directories.
    targets = []
    for item in selections:
        path = (EXPORTS / item['relative']).resolve()
        if not path.is_relative_to(EXPORTS.resolve()):
            raise ValueError('Export path is outside the report folder.')
        targets.append((item, path))
    deleted, failed = [], []
    missing = len(requested) - len(selections)
    for item, path in targets:
        try:
            path.unlink()
            deleted.append(item['id'])
        except FileNotFoundError:
            missing += 1
        except OSError:
            failed.append({'id': item['id'], 'name': item['name'],
                           'error': 'Không thể xóa file. Hãy đóng file trong Excel / WPS rồi thử lại.'})
    return {'deleted_ids': deleted, 'deleted_count': len(deleted), 'missing_count': missing, 'failed': failed}


def qc_ordered_findings(rows,rules):
    books={}
    for row in rows:
        rule=rules[row['rule_index']-1]
        spec=rule.get('right') or rule['left']
        if spec['workbook'] not in books: books[spec['workbook']]=workbook(spec['workbook'])
    book_order={key:index for index,key in enumerate(books)}
    sheet_order={key:{name:index for index,name in enumerate(book['sheets'])} for key,book in books.items()}
    def key(row):
        rule=rules[row['rule_index']-1]
        spec=rule.get('right') or rule['left']
        address=row.get('other_cell') or row.get('cell')
        position=coordinate_to_tuple(address) if address else (min(row.get('right_rows') or row.get('left_rows') or [10**9]),0)
        return book_order[spec['workbook']],sheet_order[spec['workbook']][row.get('other_sheet') or row['sheet']],*position,row['rule_index']
    return sorted(rows,key=key)


def run(rules):
    if not isinstance(rules, list) or not 1 <= len(rules) <= 20:
        raise ValueError('Add between 1 and 20 checks.')
    results, input_cells = [], 0
    for index, rule in enumerate(rules, 1):
        if not isinstance(rule, dict) or rule.get('kind') not in {'calculation', 'comparison', 'summary_comparison', 'formula_audit'}:
            raise ValueError('Unknown check type.')
        for field in (['source_range', 'target_range', 'weights_range'] if rule['kind'] == 'calculation' else ['left_range', 'right_range']):
            if rule.get(field):
                left, top, right, bottom = bounds(rule[field])
                input_cells += (right - left + 1) * (bottom - top + 1)
        if input_cells > 1_000_000:
            raise ValueError('The run covers more than 1,000,000 cells. Narrow the ranges or split the checks.')
        if rule['kind'] in {'summary_comparison', 'formula_audit'}:
            from spreadsheet_workflows import summary_comparison, formula_audit
            produced = summary_comparison(rule) if rule['kind'] == 'summary_comparison' else formula_audit(rule)
        else:
            produced = calculation(rule) if rule['kind'] == 'calculation' else comparison(rule)
        results.extend(dict(item, rule_index=index, number_mode=rule.get('number_mode', 'number')) for item in produced)
        if len(results) > MAX_RESULTS:
            raise ValueError('Checks exceed 50,000 results. Narrow the ranges and run again.')
    result_id = uuid4().hex
    precisions=sorted({settings(rule)[2] for rule in rules if settings(rule)[2] is not None})
    result = {'id': result_id, 'comparison_policy': 'rounded' if precisions else 'exact',
              'comparison_rounding_digits':precisions, 'created_at': datetime.now().isoformat(timespec='seconds'),
              'rules': rules, 'counts': dict(Counter(item['status'] for item in results)), 'total': len(results), 'results': results}
    result['results']=qc_ordered_findings(result['results'],rules)
    write_json(ROOT / 'results' / f'{result_id}.json', result)
    return result


def read_result(result_id):
    try:
        return json.loads((ROOT / 'results' / (identifier(result_id) + '.json')).read_text(encoding='utf-8'))
    except FileNotFoundError as exc:
        raise ValueError('Comparison result is no longer available.') from exc


def result_csv(result):
    stream = io.StringIO(newline='')
    writer = csv.writer(stream)
    fields = ['rule_index', 'rule', 'status', 'label', 'file', 'sheet', 'cell', 'other_file', 'other_sheet',
              'other_cell', 'actual', 'expected', 'delta', 'number_mode', 'valid_count', 'ignored_count', 'reason', 'formula', 'other_formula', 'left_rows', 'right_rows']
    writer.writerow(fields)
    for row in result['results']:
        values = []
        for key in fields:
            value = row.get(key, '')
            if isinstance(value, (list, dict)):
                value = json.dumps(value, ensure_ascii=False)
            if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')):
                value = "'" + value
            values.append(value)
        writer.writerow(values)
    return '\ufeff' + stream.getvalue()
