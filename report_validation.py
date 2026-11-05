"""Local, exhaustive preflight. Never log cell values or modify source files."""

import csv
from collections import OrderedDict
from copy import deepcopy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import io
import hashlib
import json
from pathlib import Path
import re
import threading
import zipfile

import openpyxl

CONTRACTS = json.loads(
    Path(__file__).with_name("report_contracts.json").read_text(encoding="utf8")
)
OLE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_PREFLIGHT_CACHE = OrderedDict()
_PREFLIGHT_CACHE_LOCK = threading.Lock()
_PREFLIGHT_CACHE_SIZE = 64


class InputFormatError(ValueError):
    """Messages written here are safe to show; parser exception text is not."""


class PreflightError(ValueError):
    def __init__(self, reports):
        self.reports = reports
        count = sum(report["error_count"] for report in reports)
        super().__init__(
            f"Có {count} lỗi đầu vào. Sửa theo vị trí được chỉ ra rồi kiểm tra lại."
        )


def schema_for(tool, group):
    registry = json.loads(
        Path(__file__).with_name("report_registry.json").read_text(encoding="utf8")
    )
    return registry[tool]["schemas"][group]


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError("invalid date")


def number(value):
    raw = str(value).strip()
    if re.fullmatch(r"\d{1,3}(,\d{3})+(\.\d+)?", raw):
        raw = raw.replace(",", "")
    result = Decimal(raw)
    if not result.is_finite() or result < 0:
        raise ValueError("invalid quantity")
    return result


def blank(value):
    return value is None or isinstance(value, str) and not value.strip()


def validate_preflight(path, tool, group, period, *, mime=None, normalize=False):
    """Reuse read-only summaries only when content and every validation rule match.

    Keep summaries in this process, never workbook rows or normalized files.
    A verified engine-ready workbook needs no normalization on its isolated copy.
    """
    path = Path(path)
    schema_id = schema_for(tool, group)
    limits = {**CONTRACTS, **CONTRACTS["schemas"][schema_id].get("limits", {})}
    try:
        if path.stat().st_size > limits["max_file_bytes"]:
            return validate_file(path, tool, group, period, mime=mime, normalize=normalize), False
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError:
        return validate_file(path, tool, group, period, mime=mime, normalize=normalize), False
    key = (
        digest.digest(), path.suffix.lower(), tool, group, period,
        mime.lower() if mime else None, schema_id,
        json.dumps(CONTRACTS, sort_keys=True),
    )
    with _PREFLIGHT_CACHE_LOCK:
        cached = _PREFLIGHT_CACHE.get(key)
        if cached is not None and (not normalize or not cached["valid"] or cached.get("engine_ready")):
            result = deepcopy(_PREFLIGHT_CACHE[key])
            _PREFLIGHT_CACHE.move_to_end(key)
            result["name"] = path.name
            if normalize and result["valid"]:
                result["staged_name"] = path.name
            return result, True
    result = validate_file(path, tool, group, period, mime=mime, normalize=normalize)
    with _PREFLIGHT_CACHE_LOCK:
        _PREFLIGHT_CACHE[key] = deepcopy(result)
        _PREFLIGHT_CACHE[key].pop("staged_name", None)
        _PREFLIGHT_CACHE.move_to_end(key)
        while len(_PREFLIGHT_CACHE) > _PREFLIGHT_CACHE_SIZE:
            _PREFLIGHT_CACHE.popitem(last=False)
    return result, False


def tables(path, limits=None, *, data_only=False):
    """Read actual signatures, including OOXML downloads named .xls by SFC."""
    path = Path(path)
    limits = {**CONTRACTS, **(limits or {})}
    if path.stat().st_size > limits["max_file_bytes"]:
        raise InputFormatError(
            "File vượt giới hạn kích thước. Chia nhỏ kỳ dữ liệu rồi chọn lại."
        )
    with path.open("rb") as stream:
        signature = stream.read(8)
    if signature.startswith(b"PK"):
        with zipfile.ZipFile(path) as package:
            if (
                sum(item.file_size for item in package.infolist())
                > limits["max_expanded_bytes"]
            ):
                raise InputFormatError(
                    "Workbook giải nén quá lớn. Tách workbook nhỏ hơn."
                )
            if "xl/workbook.xml" not in package.namelist():
                raise InputFormatError(
                    "Nội dung không phải workbook Excel. Xuất lại Excel từ nguồn."
                )
        with path.open("rb") as stream:
            book = openpyxl.load_workbook(
                stream, read_only=True, data_only=data_only, keep_links=False
            )
            try:
                for sheet in book:
                    if sheet.max_row and sheet.max_row > limits["max_rows"]:
                        raise InputFormatError(
                            "Sheet vượt giới hạn dòng. Chia nhỏ dữ liệu."
                        )
                    if sheet.max_column and sheet.max_column > limits["max_columns"]:
                        raise InputFormatError(
                            "Sheet quá nhiều cột. Xóa vùng định dạng thừa rồi lưu lại."
                        )
                    yield (
                        sheet.title,
                        [tuple(row) for row in sheet.iter_rows(values_only=True)],
                    )
            finally:
                book.close()
    elif signature == OLE:
        import xlrd

        book = xlrd.open_workbook(str(path), on_demand=True, logfile=io.StringIO())
        try:
            for sheet in book.sheets():
                if (
                    sheet.nrows > limits["max_rows"]
                    or sheet.ncols > limits["max_columns"]
                ):
                    raise InputFormatError(
                        "Sheet vượt giới hạn dòng/cột. Chia nhỏ dữ liệu."
                    )
                yield (
                    sheet.name,
                    [
                        tuple(
                            xlrd.xldate_as_datetime(cell.value, book.datemode)
                            if cell.ctype == xlrd.XL_CELL_DATE
                            else cell.value
                            for cell in sheet.row(r)
                        )
                        for r in range(sheet.nrows)
                    ],
                )
        finally:
            book.release_resources()
    elif path.suffix.lower() == ".csv":
        if b"\x00" in signature or signature.lstrip().startswith(b"<"):
            raise InputFormatError(
                "CSV chứa nội dung nhị phân/HTML. Lưu lại dạng CSV UTF-8."
            )
        with path.open(encoding="utf-8-sig", newline="") as stream:
            rows = []
            for row in csv.reader(stream, strict=True):
                if len(rows) >= limits["max_rows"] or len(row) > limits["max_columns"]:
                    raise InputFormatError(
                        "CSV vượt giới hạn dòng/cột. Chia nhỏ dữ liệu."
                    )
                rows.append(tuple(row))
            yield "CSV", rows
    else:
        raise InputFormatError(
            "Nội dung không phải Excel/CSV hợp lệ; có thể là trang đăng nhập. Tải lại từ nguồn."
        )


def validate_file(path, tool, group, period, *, mime=None, normalize=False):
    path = Path(path)
    schema_id = schema_for(tool, group)
    schema = CONTRACTS["schemas"][schema_id]
    result = dict(
        name=path.name,
        group=group,
        schema=schema_id,
        sheets=[],
        rows=0,
        error_count=0,
        issues=[],
        excluded_summary_rows=[],
        blank_rows=0,
        totals={},
        bc_rows=0,
        bc_quantity="0",
        business_exclusions=dict(rework_rows=0, rework_cost="0"),
        category_counts={},
    )

    def issue(message, sheet=None, row=None, column=None, *, cell_type=None):
        result["error_count"] += 1
        if len(result["issues"]) < 100:
            result["issues"].append(
                dict(sheet=sheet, row=row, column=column, message=message,
                     **({"cell_type": cell_type} if cell_type else {}))
            )

    if mime and mime.lower() not in (
        "application/octet-stream",
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.ms-excel.sheet.macroenabled.12",
        "text/csv",
        "application/csv",
    ):
        issue("MIME không phù hợp. Chọn file Excel/CSV gốc thay vì đổi đuôi file.")
    prepared = []
    try:
        if schema_id == "quarterly_quality":
            from report_tools.quarterly_quality.quarterly_quality import read_quarter

            source_tables = dict(tables(path, schema.get("limits"), data_only=True))
            quarter, errors = read_quarter(source_tables, period, result["sheets"])
            result["coverage"] = dict(months=list(quarter.covered), expected=list(quarter.months))
            for message, sheet, row, column in errors:
                issue(message, sheet, row, column)
            result["rows"] = sum(sheet["rows"] for sheet in result["sheets"])
            result["valid"] = not result["error_count"]
            return result
        source_tables = list(tables(path, schema.get("limits")))
    except Exception as error:
        # Parser exceptions may contain cell contents or absolute paths.
        message = (
            str(error)
            if isinstance(error, InputFormatError)
            else "Không đọc được file. Lưu/đóng Excel, xuất lại file đúng định dạng và thử lại."
        )
        issue(message)
        result["valid"] = False
        return result
    populated = [
        (title, rows)
        for title, rows in source_tables
        if any(any(not blank(v) for v in row) for row in rows)
    ]
    if not populated:
        issue("File rỗng. Cần ít nhất hàng tiêu đề đúng mẫu.")
    if (not schema.get("multiple_sheets") or tool == "reinspection") and len(
        populated
    ) != 1:
        issue(
            "Tool đọc một sheet dữ liệu. Tách từng bảng thành file riêng để không bỏ sót sheet."
        )
    engine_ready = False
    if tool == "reinspection" and group == "ftt" and path.suffix.lower() == ".xlsx" and len(source_tables) == 1:
        with path.open("rb") as stream:
            engine_ready = stream.read(2) == b"PK"
    for title, rows in populated:
        required = schema["required"]
        aliases = schema.get("aliases", {})
        # Only apply documented engine aliases when the canonical column is absent.
        for i, row in enumerate(rows[: schema.get("header_scan", 1)]):
            rows[i] = tuple(
                aliases.get(v, v) if aliases.get(v) not in row else v for v in row
            )
            if rows[i] != row:
                engine_ready = False
        header_index = next(
            (
                i
                for i, row in enumerate(rows[: schema.get("header_scan", 1)])
                if set(required) <= set(row)
            ),
            None,
        )
        if header_index is None:
            best = max(
                rows[: schema.get("header_scan", 1)],
                key=lambda row: len(set(required) & set(row)),
                default=(),
            )
            missing = [name for name in required if name not in best]
            issue(
                "Thiếu cột đúng tên: "
                + ", ".join(missing)
                + ". Xuất lại đúng mẫu; không đổi tên cột.",
                title,
                1,
            )
            continue
        header = list(rows[header_index])
        while header and blank(header[-1]):
            header.pop()
        labels = [str(v) if not blank(v) else "" for v in header]
        engine_ready = engine_ready and header_index == 0 and list(rows[0]) == labels
        duplicates = {
            v
            for v in labels
            if v
            and labels.count(v) > 1
            and not (schema_id == "reinspection" and v == "Defective Rate")
        }
        if duplicates:
            issue(
                "Trùng tên cột; sửa tiêu đề để mỗi cột có tên riêng.",
                title,
                header_index + 1,
            )
        data = []
        numeric = set(schema.get("numbers", [])) & set(labels)
        if schema.get("defects_after") in labels:
            numeric.update(labels[labels.index(schema["defects_after"]) + 1 :])
        date_columns = schema.get("dates", [])
        last_data_row = max(
            (i + 1 for i, row in enumerate(rows) if any(not blank(v) for v in row)),
            default=0,
        )
        for index, raw in enumerate(rows[header_index + 1 :], header_index + 2):
            values = list(raw) + [None] * max(0, len(labels) - len(raw))
            if all(blank(v) for v in values):
                engine_ready = False
                result["blank_rows"] += 1
                continue
            # Explicit DevExpress count footer: no keys, only date count / numeric totals.
            nonempty = {
                labels[i] if i < len(labels) else ""
                for i, v in enumerate(values)
                if not blank(v)
            }
            date_label = date_columns[0] if date_columns else None
            count = values[labels.index(date_label)] if date_label in labels else None
            if schema_id == "copq" and isinstance(count, str):
                matched_count = re.fullmatch(
                    r"Count\s*=\s*([\d,]+)", count.strip(), re.I
                )
                if matched_count:
                    count = matched_count[1]
            try:
                footer_count_matches = number(count) == len(data)
            except (ValueError, InvalidOperation):
                footer_count_matches = False
            is_footer = (
                schema_id in ("hfpa", "reinspection", "copq")
                and index == last_data_row
                and nonempty <= (numeric | {date_label})
                and footer_count_matches
            )
            if is_footer:
                engine_ready = False
                result["excluded_summary_rows"].append(
                    dict(
                        sheet=title,
                        row=index,
                        count_column=openpyxl.utils.get_column_letter(labels.index(date_label) + 1),
                        reason="Dòng tổng DevExpress đã xác nhận; không phải bản ghi.",
                    )
                )
                continue
            if len(raw) > len(labels) and any(not blank(v) for v in raw[len(labels) :]):
                issue(
                    "Có dữ liệu ngoài cột tiêu đề. Bổ sung tiêu đề hoặc xuất lại bảng.",
                    title,
                    index,
                )
            for j, label in enumerate(labels):
                value = values[j]
                column = openpyxl.utils.get_column_letter(j + 1)
                if not label and not blank(value):
                    issue(
                        "Cột có dữ liệu nhưng thiếu tiêu đề. Bổ sung tiêu đề.",
                        title,
                        index,
                        column,
                    )
                cell_type = "number" if label in numeric else "text"
                if isinstance(value, str) and value.startswith("="):
                    issue(
                        "Đầu vào chứa công thức. Dán giá trị (Paste Values) rồi kiểm tra lại.",
                        title,
                        index,
                        column,
                        cell_type=cell_type,
                    )
                if label in schema.get("keys", []) and blank(value):
                    issue(
                        "Thiếu giá trị bắt buộc. Điền dữ liệu hoặc sửa file nguồn.",
                        title,
                        index,
                        column,
                        cell_type=cell_type,
                    )
                if label in date_columns:
                    # Original audit date can precede the reinspection month, or be absent.
                    if blank(value) and schema_id == "reinspection":
                        continue
                    try:
                        day = parse_date(value)
                        if (
                            schema_id != "reinspection"
                            and day.strftime("%Y-%m") != period
                        ):
                            issue(
                                "Ngày nằm ngoài tháng đã chọn. Chọn đúng tháng hoặc xuất lại đúng kỳ.",
                                title,
                                index,
                                column,
                                cell_type=cell_type,
                            )
                    except ValueError:
                        issue(
                            "Ngày không hợp lệ. Dùng ngày Excel hoặc YYYY-MM-DD.",
                            title,
                            index,
                            column,
                            cell_type=cell_type,
                        )
                if engine_ready and label in numeric and value not in (None, "") and type(value) not in (int, float):
                    engine_ready = False
                if label in numeric and not blank(value):
                    try:
                        value_number = number(value)
                        integer_quantity = (
                            label
                            in (
                                "Issues Q'ty",
                                "InspQty",
                                "FailQty",
                                "IssueQty",
                                "Reinspection Q'ty",
                                "2nd Audit Sample Size",
                                "Audit Sample Size",
                                "Defect Q'ty (Total Defect Q'ty)",
                            )
                            or schema.get("defects_after") in labels
                            and j > labels.index(schema["defects_after"])
                        )
                        if (
                            integer_quantity
                            and value_number != value_number.to_integral_value()
                        ):
                            raise ValueError("fractional quantity")
                        values[j] = float(value_number)
                        result["totals"][label] = result["totals"].get(
                            label, 0
                        ) + float(value_number)
                    except (ValueError, InvalidOperation):
                        issue(
                            "Số không hợp lệ/âm hoặc số lượng không nguyên. Sửa giá trị; không tự thay bằng 0.",
                            title,
                            index,
                            column,
                            cell_type=cell_type,
                        )
                if (
                    schema_id == "qastation"
                    and label == "Station"
                    and str(value).strip().upper() != "HFPA"
                ):
                    issue(
                        "Có trạm khác HFPA. Xuất QAStation với duy nhất trạm HFPA.",
                        title,
                        index,
                        column,
                        cell_type=cell_type,
                    )
                if (
                    schema_id == "reinspection"
                    and label == "Category"
                    and value not in ("QC", "MA")
                ):
                    issue(
                        "Category phải là QC hoặc MA. Kiểm tra lại mẫu Re-Inspection.",
                        title,
                        index,
                        column,
                        cell_type=cell_type,
                    )
                if (
                    label in ("Pass/Fail", "2nd Pass/Fail")
                    and not blank(value)
                    and value not in ("Pass", "Fail")
                ):
                    issue(
                        "Trạng thái phải là Pass hoặc Fail. Sửa từ nguồn.",
                        title,
                        index,
                        column,
                        cell_type=cell_type,
                    )
            data.append(values[: len(labels)])
            if (
                schema_id == "copq"
                and str(values[labels.index("Type")] or "").strip().lower() == "rework"
            ):
                result["business_exclusions"]["rework_rows"] += 1
                try:
                    result["business_exclusions"]["rework_cost"] = str(
                        Decimal(result["business_exclusions"]["rework_cost"])
                        + number(values[labels.index("Ttl cost ($)")] or 0)
                    )
                except (ValueError, InvalidOperation):
                    pass  # Reported above; invalid data never reaches an engine.
            if schema_id == "reinspection":
                category = values[labels.index("Category")]
                if category in ("QC", "MA"):
                    result["category_counts"][category] = (
                        result["category_counts"].get(category, 0) + 1
                    )
            if schema_id == "ftt":
                metric = str(values[labels.index("Metric")] or "").strip()
                defect = str(values[labels.index("Defect Issues")] or "").strip()
                if re.search(
                    r"B/Reject|C/Reject|B-Reject|C-Reject|^B$|^C$|Grade B|Grade C",
                    metric,
                    re.I,
                ) or re.search(r"^B/|^C/|^B-|^C-", defect, re.I):
                    result["bc_rows"] += 1
                    try:
                        result["bc_quantity"] = str(
                            Decimal(result["bc_quantity"])
                            + number(values[labels.index("Issues Q'ty")] or 0)
                        )
                    except (ValueError, InvalidOperation):
                        pass
        result["sheets"].append(
            dict(name=title, header_row=header_index + 1, rows=len(data))
        )
        result["rows"] += len(data)
        prepared.append((title, labels, data))
    result["valid"] = result["error_count"] == 0
    if engine_ready and result["valid"]:
        result["engine_ready"] = True
        if normalize:
            result["staged_name"] = path.name
    elif normalize and result["valid"]:
        # Engines use differing readers; canonical copies ensure identical header/sheet treatment.
        destination = path.with_suffix(".xlsx")
        book = openpyxl.Workbook(write_only=True)
        for title, labels, rows in prepared:
            sheet = book.create_sheet(title if title != "CSV" else "Data")
            sheet.append(labels)
            for row in rows:
                sheet.append(row)
        temporary = destination.with_name(destination.name + ".validated")
        book.save(temporary)
        book.close()
        temporary.replace(destination)
        if destination != path:
            path.unlink()  # only the isolated staged copy
        result["staged_name"] = destination.name
    return result
