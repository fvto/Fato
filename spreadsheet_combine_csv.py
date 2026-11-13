"""Bounded CSV ingestion; retain text values and use the existing table cleaner."""
import codecs
import csv
import io
import re
from collections import Counter

from charset_normalizer import from_bytes

ENCODINGS = {'auto', 'utf-8-sig', 'utf-16', 'utf-32', 'cp1258', 'cp1252',
             'cp1250', 'iso8859-1', 'shift_jis', 'gb18030', 'big5'}
DELIMITERS = {'auto', ',', ';', '\t', '|'}
SAMPLE_BYTES = 128 * 1024
INVALID_CONTROL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def verified_encoding(path, requested='auto'):
    """Validate the complete file, including legacy bytes after an ASCII prefix."""
    with path.open('rb') as stream:
        raw = stream.read(SAMPLE_BYTES)
    detected, guessed = detect_encoding(raw) if requested == 'auto' else (requested, False)
    candidates = [detected]
    if requested == 'auto' and not raw.startswith((codecs.BOM_UTF8,codecs.BOM_UTF16_LE,codecs.BOM_UTF16_BE,codecs.BOM_UTF32_LE,codecs.BOM_UTF32_BE)):
        candidates += ['cp1252', 'cp1258', 'cp1250', 'iso8859-1']
    for encoding in dict.fromkeys(candidates):
        decoder = codecs.getincrementaldecoder(encoding)(errors='strict')
        try:
            with path.open('rb') as stream:
                while chunk := stream.read(1024 * 1024):
                    decoder.decode(chunk)
                decoder.decode(b'', final=True)
            return encoding, guessed or encoding != detected
        except UnicodeError:
            continue
    raise ValueError('Không đọc được CSV với encoding đã chọn. Chọn lại encoding hoặc lưu lại dạng UTF-8.')


def configuration(path, csv_options=None):
    config = options(csv_options)
    encoding, guessed = verified_encoding(path, config['encoding'])
    with path.open('rb') as stream:
        text = sample_text(stream.read(SAMPLE_BYTES), encoding)
    if '\0' in text:
        raise ValueError('CSV chứa byte NUL. Chọn đúng encoding (UTF-16 cần BOM) hoặc lưu lại dạng UTF-8.')
    lines = text.splitlines()
    first = lines[0] if lines else ''
    directive = len(first) == 5 and first.lower().startswith('sep=') and first[-1] in DELIMITERS - {'auto'}
    delimiter = config['delimiter']
    if delimiter == 'auto':
        delimiter = first[-1] if directive else detect_delimiter(text)
    warnings = []
    if guessed:
        warnings.append(f'Encoding {encoding} được suy đoán sau khi kiểm tra toàn bộ file; kiểm tra lại ký tự trong bản xem trước.')
    return dict(encoding=encoding, delimiter=delimiter, directive=directive, warnings=warnings)


def iter_records(path, config):
    """Stream logical records without a cell-count or memory-sized list."""
    with path.open(encoding=config['encoding'], errors='strict', newline='') as stream:
        if config['directive']:
            stream.readline()
            yield []
        reader = csv.reader(stream, delimiter=config['delimiter'], strict=True)
        try:
            for row in reader:
                validate_record(row, reader.line_num)
                yield row
        except csv.Error as error:
            raise ValueError(f'CSV không hợp lệ gần dòng {reader.line_num}: {error}. Kiểm tra dấu nháy và dấu phân cách.') from error


def validate_record(row, number):
    if len(row) > 16384:
        raise ValueError(f'CSV gần dòng {number}: vượt giới hạn cột Excel.')
    # Do length checks in C and scan small groups instead of invoking the regex
    # for every cell. At most 64 Excel-sized fields are joined at a time.
    if max(map(len, row), default=0) > 32767 or any(
            INVALID_CONTROL.search('\t'.join(row[start:start+64]))
            for start in range(0, len(row), 64)):
        raise ValueError(f'CSV gần dòng {number}: ô quá dài hoặc có ký tự điều khiển không hợp lệ.')


def options(value):
    value = {} if value is None else value
    if (not isinstance(value, dict) or set(value) - {'encoding', 'delimiter'}
            or not isinstance(value.get('encoding', 'auto'), str)
            or value.get('encoding', 'auto') not in ENCODINGS
            or not isinstance(value.get('delimiter', 'auto'), str)
            or value.get('delimiter', 'auto') not in DELIMITERS):
        raise ValueError('Tùy chọn encoding hoặc dấu phân cách CSV không hợp lệ.')
    return dict(encoding=value.get('encoding', 'auto'), delimiter=value.get('delimiter', 'auto'))


def sample_text(raw, encoding):
    # The bounded sample can end halfway through a multibyte character.
    return codecs.getincrementaldecoder(encoding)(errors='strict').decode(raw, final=False)


def detect_encoding(raw):
    for bom, encoding in [(codecs.BOM_UTF32_LE, 'utf-32'), (codecs.BOM_UTF32_BE, 'utf-32'),
                          (codecs.BOM_UTF8, 'utf-8-sig'),
                          (codecs.BOM_UTF16_LE, 'utf-16'), (codecs.BOM_UTF16_BE, 'utf-16')]:
        if raw.startswith(bom):
            return encoding, False
    try:
        sample_text(raw, 'utf-8-sig')
        return 'utf-8-sig', False
    except UnicodeError:
        match = from_bytes(raw).best()
        if match is None or match.percent_chaos > 20:
            raise ValueError('Không nhận diện được encoding CSV. Chọn encoding trong Tùy chọn CSV hoặc lưu lại dạng UTF-8.')
        return match.encoding, True


def detect_delimiter(text):
    try:
        return csv.Sniffer().sniff(text, delimiters=',;\t|').delimiter
    except csv.Error:
        # Report titles, blank lines and footers often defeat Sniffer. Rank
        # repeated widths, ignoring one-column preambles and a truncated tail.
        scores = []
        for delimiter in ',;\t|':
            widths = Counter()
            try:
                reader = csv.reader(io.StringIO(text, newline=''), delimiter=delimiter, strict=True)
                for index, row in enumerate(reader):
                    if len(row) > 1:
                        widths[len(row)] += 1
                    if index >= 99:
                        break
            except csv.Error:
                pass
            width, count = max(widths.items(), key=lambda pair: (pair[1], pair[0]), default=(1, 0))
            scores.append((count, width, delimiter))
        return max(scores, key=lambda item: item[:2])[2] if any(s[0] for s in scores) else ','


def read_csv(path, max_cells, csv_options=None):
    config = options(csv_options)
    with path.open('rb') as stream:
        raw = stream.read(SAMPLE_BYTES)
    encoding, guessed = verified_encoding(path, config['encoding'])
    try:
        text = sample_text(raw, encoding)
        if '\0' in text:
            raise ValueError('CSV chứa byte NUL. Chọn đúng encoding (UTF-16 cần BOM) hoặc lưu lại dạng UTF-8.')
        first_line = text.splitlines()[0] if text.splitlines() else ''
        directive = len(first_line) == 5 and first_line.lower().startswith('sep=') and first_line[-1] in DELIMITERS - {'auto'}
        delimiter = config['delimiter']
        if delimiter == 'auto':
            delimiter = first_line[-1] if directive else detect_delimiter(text)
        rows, formats, cells = [], [], 0
        with path.open(encoding=encoding, errors='strict', newline='') as stream:
            if directive:
                stream.readline()
                # Retain the directive's position for source record numbers.
                rows.append([])
                formats.append([])
            reader = csv.reader(stream, delimiter=delimiter, strict=True)
            for row in reader:
                # Blank records also consume budget, so millions of blank
                # lines cannot bypass the memory bound.
                cells += max(1, len(row))
                if cells > max_cells:
                    raise ValueError('File vượt 500.000 ô. Hãy chia nhỏ dữ liệu.')
                validate_record(row, reader.line_num)
                rows.append(row)
                formats.append(['General'] * len(row))
    except UnicodeError as error:
        raise ValueError('Không đọc được CSV với encoding đã chọn/nhận diện. Chọn lại encoding trong Tùy chọn CSV hoặc lưu lại dạng UTF-8.') from error
    except csv.Error as error:
        raise ValueError(f'CSV không hợp lệ gần dòng {reader.line_num if "reader" in locals() else 1}: {error}. Kiểm tra dấu nháy và dấu phân cách.') from error
    warnings = []
    if guessed:
        warnings.append(f'Encoding {encoding} được suy đoán; kiểm tra dấu tiếng Việt/ký tự đặc biệt. Nếu sai, chọn encoding CSV rồi tải lại.')
    return dict(name='CSV', rows=rows, formats=formats,
                csv=dict(encoding=encoding, delimiter=delimiter), warnings=warnings)
