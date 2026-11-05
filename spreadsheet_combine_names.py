"""Report-independent naming from shared filename context and explicit periods."""
from datetime import date
from pathlib import Path
import re
import unicodedata

PERIOD = re.compile(r'(?<!\d)(?:\d{4}[-_.]\d{2}[-_.]\d{2}|(?:19|20)\d{6}|\d{4}[-_.]\d{2})(?!\d)')


def context(name):
    stem = re.sub(r'\s*\(\d+\)$', '', Path(name).stem)
    spans, periods = [], []
    for match in PERIOD.finditer(stem):
        raw = match[0]
        normalized = re.sub(r'[_\.]', '-', raw)
        try:
            if len(raw) == 8 and raw.isdigit():
                normalized = f'{raw[:4]}-{raw[4:6]}-{raw[6:]}'
            if len(normalized) == 7:
                date.fromisoformat(normalized + '-01')
            else:
                date.fromisoformat(normalized)
        except ValueError:
            return None
        spans.append(match.span())
        periods.append(normalized)
    for start, end in reversed(spans):
        stem = stem[:start] + ' ' + stem[end:]
    if len(periods) > 1 and (len({len(p) for p in periods}) != 1 or periods[0] > periods[-1]):
        return None
    tokens = re.findall(r'[^\W_]+', unicodedata.normalize('NFKC', stem), re.UNICODE)
    while tokens and tokens[-1].casefold() in {'combined', 'merged'}:
        tokens.pop()
    return dict(tokens=tokens, periods=periods)


def common_context(contexts):
    rows = [c['tokens'] for c in contexts]
    prefix = []
    for group in zip(*rows):
        if len({token.casefold() for token in group}) != 1 or group[0].isdigit():
            break
        prefix.append(group[0])
    if not prefix:
        return ''
    tails = [row[len(prefix):] for row in rows]
    suffix = []
    for group in zip(*(reversed(row) for row in tails)):
        if len({token.casefold() for token in group}) != 1 or group[0].isdigit():
            break
        suffix.insert(0, group[0])
    return '_'.join(prefix + suffix)


def naming_details(sources):
    names = list(dict.fromkeys(s['file'] for s in sources))
    sheets = list(dict.fromkeys(s['sheet'] for s in sources))
    contexts = [context(name) for name in names]
    common = common_context(contexts) if all(c is not None for c in contexts) else ''
    label = common or (Path(names[0]).stem if len(names) == 1 else sheets[0] if len(sheets) == 1 else 'Excel')
    label = re.sub(r'[<>:"/\\|?*\x00-\x1f\x7f]', '_', label)[:90].strip(' ._') or 'Excel'
    periods = [p for c in contexts if c for p in c['periods']]
    complete_period = bool(common and periods and all(c and c['periods'] for c in contexts))
    # Monthly and daily periods are kept distinct; never infer the last day of
    # a month or interpret a week number as an exact date.
    same_grain = len({len(p) for p in periods}) == 1
    if complete_period and same_grain:
        first, last = min(periods), max(periods)
        period = first if first == last and len(first) == 7 else f'{first}_{last}'
        stem = f'{label}_{period}_Combined'
        basis = 'Tên chung và kỳ trong tên các file nguồn đã chọn.'
    else:
        stem = f'{label}_{len(names)}files_Combined' if len(names) > 1 else f'{label}_Combined'
        basis = 'Tên chung của file nguồn.' if common else 'Tên trang tính chung hoặc tên file nguồn.'
        basis += ' Không thêm ngày khi kỳ nguồn chưa rõ hoặc không đồng nhất.'
    return dict(output_name=stem + '.xlsx', naming_basis=basis)
