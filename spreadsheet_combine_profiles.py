"""Opt-in, schema-recognized export adapters; generic append owns no report type."""
import math


def remove_known_footer(keys, records):
    required = {'factory code', 'so#', 'audit date', "defective q'ty(pair)", "total defect q'ty"}
    if not required.issubset(keys) or not records:
        return records, 0, ''
    _, tail, _ = records[-1]
    audit, defect = keys.index('audit date'), keys.index("defective q'ty(pair)")
    if not all(v in (None, '') for i, v in enumerate(tail) if i not in (audit, defect)):
        return records, 0, 'SFC audit'
    try:
        valid = float(str(tail[audit]).replace(',', '')) == len(records) - 1
        expected = sum(float(record[1][defect] or 0) for record in records[:-1])
        valid = valid and math.isclose(float(str(tail[defect]).replace(',', '')), expected, abs_tol=1e-9)
    except (TypeError, ValueError):
        valid = False
    if not valid:
        raise ValueError('Dòng tổng FGWH/SFC không khớp số dòng hoặc số đôi lỗi. Kiểm tra nguồn hoặc chọn giữ dòng tổng.')
    return records[:-1], 1, 'SFC audit'
