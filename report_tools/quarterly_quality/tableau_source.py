"""Use the existing VH FTT data and append verified records without replacing history."""

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date
from functools import lru_cache
import hashlib
import json
import math
import os
from pathlib import Path
import posixpath
import re
import shutil
import subprocess
import tempfile
from zipfile import ZipFile

from lxml import etree
from openpyxl.utils.datetime import from_excel, to_excel

from dashboard_paths import APP_ROOT
from report_validation import number, parse_date
from .quarterly_quality import Counts, PLANTS


SHEET = "VH FTT"
HEADER = ("Date", "Factory", "Line1", "Model Name", "Inspection Qty1", "Defect Qty1",
          "Station1", "Defect Issues1", "Issues Qty1")
X = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
CHART_LABELS = {"B/", "C/"}
ROW = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*(?:/>|>.*?</(?:[A-Za-z_][\w.-]*:)?row>)", re.S)
REFERENCE = re.compile(rb'\br="([A-Z]*)(\d+)"')


def default_path():
    path = APP_ROOT / "Report/Quarterly/Tableau data source 2026 HFPA Bonding.xlsx"
    return path if path.is_file() else None


def fingerprint(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def worksheet_part(package):
    book = etree.fromstring(package.read("xl/workbook.xml"))
    sheet = next((item for item in book.find(X + "sheets") if item.get("name") == SHEET), None)
    if sheet is None:
        raise ValueError("Nguồn Tableau thiếu sheet VH FTT.")
    relation = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    rels = etree.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    target = next(item.get("Target") for item in rels if item.get("Id") == relation)
    part = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
    if not part.startswith("xl/worksheets/"):
        raise ValueError("Không xác định được sheet VH FTT trong nguồn Tableau.")
    return part


def shared_strings(package):
    if "xl/sharedStrings.xml" not in package.namelist():
        return []
    values = []
    for _, item in etree.iterparse(package.open("xl/sharedStrings.xml"), events=("end",), tag=X + "si"):
        values.append("".join(node.text or "" for node in item.iter(X + "t")))
        item.clear()
    return values


def cell_value(cell, strings):
    value = cell.find(X + "v")
    value = value.text if value is not None else None
    if cell.get("t") == "s" and value is not None:
        return strings[int(value)]
    if cell.get("t") == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(X + "t"))
    return value


def row_values(row, strings):
    cells = {re.sub(r"\d", "", cell.get("r", "")): cell_value(cell, strings) for cell in row}
    return [cells.get(column) for column in "ABCDEFGHI"]


def ledger_path(path):
    return path.with_suffix(".coverage.json")


@dataclass
class TableauData:
    path: Path
    revision: str
    part: str
    cutoff: str
    values: dict = field(default_factory=dict)
    details: dict = field(default_factory=dict)
    observed: dict = field(default_factory=dict)
    coverage: dict = field(default_factory=dict)
    metrics: dict = field(default_factory=dict)
    last_row: int = 1
    row_count: int = 0

    def covered_to(self, metric, month, plant, last):
        key = metric, month, plant
        known = self.coverage.get("|".join(key), "")
        if key in self.values and month < self.cutoff[:7]:
            return last
        known = max(known, self.observed.get(key, ""))
        return min(known, last) if known else None


def read_source(path, months, cutoff):
    path = Path(path)
    stat = path.stat()
    ledger = ledger_path(path)
    ledger_modified = ledger.stat().st_mtime_ns if ledger.exists() else 0
    return _read_source(str(path.resolve()), stat.st_size, stat.st_mtime_ns, tuple(months), cutoff, ledger_modified)


@lru_cache(maxsize=4)
def _read_source(filename, size, modified, months, cutoff, ledger_modified):
    path = Path(filename)
    revision = fingerprint(path)
    with ZipFile(path) as package:
        part = worksheet_part(package)
        result = TableauData(path, revision, part, cutoff)
        saved = ledger_path(path)
        if saved.exists():
            ledger = json.loads(saved.read_text(encoding="utf-8"))
            if ledger.get("sha256") == revision:
                result.coverage = ledger.get("coverage", {})
                result.metrics = ledger.get("metrics", {})
        strings = shared_strings(package)
        dates, seen = {}, set()
        quantities = defaultdict(lambda: [0.0, 0.0])
        details = defaultdict(lambda: (Counter(), defaultdict(Counter), Counter()))
        for _, row in etree.iterparse(package.open(part), events=("end",), tag=X + "row"):
            index = int(row.get("r"))
            result.last_row = index
            if index == 1:
                if tuple(row_values(row, strings)) != HEADER:
                    raise ValueError("Sheet VH FTT cần giữ đúng chín cột của file Tableau mẫu.")
            elif len(row):
                raw_date = cell_value(row[0], strings)
                if raw_date not in dates:
                    try:
                        day = from_excel(float(raw_date)).date()
                    except (TypeError, ValueError, OverflowError):
                        day = parse_date(str(raw_date).replace("/", "-"))
                    dates[raw_date] = day.isoformat()
                day = dates[raw_date]
                if day[:7] in months and day <= cutoff:
                    values = row_values(row, strings)
                    factory, station, issue = values[1], values[6], str(values[7] or "").strip()
                    if factory in PLANTS and station in ("Assembly", "HFPA") and issue not in CHART_LABELS:
                        result.row_count += 1
                        metric = "HFPA" if station == "HFPA" else "FTT"
                        key = metric, day[:7], factory
                        base, bad, issue_qty = (float(number(values[i] or 0)) for i in (4, 5, 8))
                        if metric == 'HFPA' and bad > base:
                            raise ValueError(f"VH FTT dòng {index} có số đôi lỗi lớn hơn số đôi kiểm/sản xuất.")
                        identity = (*key, day, values[2], values[3], base, bad)
                        model = str(values[3] or "").strip()
                        if identity not in seen:
                            seen.add(identity)
                            quantities[key][0] += base
                            quantities[key][1] += bad
                            if model:
                                details[key][0][model] += bad
                                details[key][2][model] += base
                        if model and issue and issue_qty:
                            details[key][1][model][issue] += issue_qty
                        result.observed[key] = max(result.observed.get(key, ""), day)
            row.clear()
            while row.getprevious() is not None:
                del row.getparent()[0]
        result.values = {key: Counts(*counts) for key, counts in quantities.items()}
        result.details = dict(details)
        for entry in result.coverage:
            metric, month, plant = entry.split("|")
            if month in months:
                result.values.setdefault((metric, month, plant), Counts(0, 0))
        for key, counts in tuple(result.values.items()):
            saved_counts = result.metrics.get('|'.join(key))
            if saved_counts and (saved_counts['inspected'], saved_counts['defects']) == (counts.inspected, counts.defects):
                result.values[key] = Counts(**saved_counts)
    after = path.stat()
    if (size, modified) != (after.st_size, after.st_mtime_ns):
        raise ValueError("File Tableau thay đổi trong khi đọc. Lưu file rồi kiểm tra lại.")
    return result


def normalize_ftt(export, item):
    """Reduce validated FTT lots and metric repetitions to the existing nine-column grain."""
    from .sfc_inputs import FTT_LOT
    from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent

    header, rows = export
    columns = {str(label or "").strip(): index for index, label in enumerate(header)}
    lots, metrics, lot_groups = {}, {}, {}
    issues = defaultdict(Counter)
    for row in rows:
        if str(row[columns["Station"]]).strip().casefold() != "assembly":
            continue
        day = parse_date(row[columns["Date"]]).isoformat()
        lot = (day, *(row[columns[name]] for name in FTT_LOT[1:]))
        metric = str(row[columns["Metric"]]).strip()
        model = str(row[columns["Model 2"]] or row[columns["Model"]] or "").strip()
        plant = str(row[columns["Plant"]] or "").strip()
        parsed_plant, line = ReconciliationEngineAgent.parse_plant_and_line(row[columns["Line"]])
        plant = plant or parsed_plant
        group = day, item["plant"], f"{plant} Line {line}".strip(), model
        lots[lot] = float(number(row[columns["Total Pair Produced"]]))
        metrics[lot, metric] = float(number(row[columns["Number of pairs"]]))
        lot_groups[lot] = group
        if metric != "AGrade" and "Defect Issues" in columns and "Issues Q'ty" in columns:
            label = str(row[columns["Defect Issues"]] or "").strip()
            label = metric if label in CHART_LABELS else label
            quantity = float(number(row[columns["Issues Q'ty"]] or 0))
            if label and quantity:
                issues[group][label] += quantity
    totals = defaultdict(lambda: [0.0, 0.0])
    for lot, produced in lots.items():
        totals[lot_groups[lot]][0] += produced
    for (lot, metric), quantity in metrics.items():
        if metric != "AGrade":
            totals[lot_groups[lot]][1] += quantity
    result = []
    for group, (produced, failed) in sorted(totals.items()):
        for issue, quantity in issues[group].items() or [("", 0)]:
            result.append([*group, produced, failed, "Assembly", issue, quantity])
    return result


def xml_rows(stream):
    """Preserve the original row XML without building the 430,000-row worksheet in memory."""
    pending = b""
    while chunk := stream.read(1024 * 1024):
        pending += chunk
        start = 0
        for match in ROW.finditer(pending):
            yield pending[start:match.start()], match.group()
            start = match.end()
        pending = pending[start:]
    yield pending, None


def renumber(row, index):
    return REFERENCE.sub(lambda match: b'r="' + match[1] + str(index).encode() + b'"', row)


def append_records(source, records, windows, stage, *, metrics=None):
    """Commit one worksheet part atomically and retain all other workbook parts byte-for-byte."""
    if fingerprint(source.path) != source.revision:
        raise ValueError("Nguồn Tableau đã thay đổi từ lúc kiểm tra. Kiểm tra đầu vào lại trước khi bổ sung.")
    if not records:
        save_coverage(source, windows, metrics)
        return dict(path=str(source.path), sheet=SHEET, added_rows=0, backup=None)
    dependency = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node"
    node, modules = dependency / "bin/node.exe", dependency / "node_modules"
    if not node.is_file():
        raise ValueError("Chưa có bộ xử lý Excel để bổ sung nguồn Tableau.")
    stage(f"Đang bổ sung {len(records):,} dòng vào sheet VH FTT...", "processing")
    with tempfile.TemporaryDirectory(prefix="tableau-append-", dir=source.path.parent) as temporary:
        temporary = Path(temporary)
        authored_files = []
        for start in range(0, len(records), 10000):
            batch = records[start:start + 10000]
            rows = [[to_excel(date.fromisoformat(row[0])), *row[1:]] for row in batch]
            payload = temporary / f"rows-{start}.json"
            payload.write_text(json.dumps([list(HEADER), *rows], ensure_ascii=False), encoding="utf-8")
            authored = temporary / f"append-{start}.xlsx"
            stage(f"Đang chuẩn bị dòng {start + 1:,}–{start + len(batch):,}/{len(records):,} cho VH FTT...", "processing")
            result = subprocess.run([str(node), str(Path(__file__).with_name("tableau_append.mjs")),
                                     str(payload), str(authored), str(modules)],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, timeout=120,
                                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if result.returncode:
                raise ValueError(f"Bộ xử lý Excel chưa ghi được lô dữ liệu bắt đầu tại dòng {start + 1:,}.")
            authored_files.append(authored)
            payload.unlink()
        candidate = temporary / source.path.name
        with ZipFile(source.path) as original, ZipFile(candidate, "w") as output:
            strings = shared_strings(original)
            # All new values are authored with the spreadsheet runtime. Reuse the source's styles.
            new_rows = []
            for authored in authored_files:
                with ZipFile(authored) as addition:
                    added_strings = shared_strings(addition)
                    append_part = worksheet_part(addition)
                    for _, row in etree.iterparse(addition.open(append_part), events=("end",), tag=X + "row"):
                        if int(row.get("r")) > 1:
                            for cell in row:
                                value = cell_value(cell, added_strings)
                                if cell.get("t") == "s":
                                    for child in tuple(cell):
                                        cell.remove(child)
                                    cell.set("t", "inlineStr")
                                    etree.SubElement(etree.SubElement(cell, X + "is"), X + "t").text = value or ""
                            new_rows.append(etree.tostring(row, encoding="utf-8"))
                        row.clear()
            next_row, styles = 0, None
            final_count = source.last_row + len(records)
            for entry in original.infolist():
                if entry.filename != source.part:
                    with original.open(entry) as incoming, output.open(entry, "w") as outgoing:
                        shutil.copyfileobj(incoming, outgoing, 1024 * 1024)
                    continue
                with original.open(entry) as incoming, output.open(entry, "w") as outgoing:
                    for prefix, raw in xml_rows(incoming):
                        if prefix:
                            prefix = re.sub(rb'(ref="A1:I)\d+("?)',
                                            lambda match: match[1] + str(final_count).encode() + match[2], prefix)
                        if raw is None:
                            for new_row in new_rows:
                                next_row += 1
                                element = etree.fromstring(new_row)
                                for cell in element:
                                    column = re.sub(r"\d", "", cell.get("r", ""))
                                    if styles and column in styles:
                                        cell.set("s", styles[column])
                                    else:
                                        cell.attrib.pop("s", None)
                                outgoing.write(renumber(etree.tostring(element), next_row))
                            outgoing.write(prefix)
                            break
                        outgoing.write(prefix)
                        element = etree.fromstring(b'<root xmlns="' + X[1:-1].encode() + b'">' + raw + b'</root>')[0]
                        if int(element.get("r")) > 1:
                            styles = {re.sub(r"\d", "", cell.get("r", "")): cell.get("s")
                                      for cell in element if cell.get("s") is not None}
                        next_row += 1
                        outgoing.write(renumber(raw, next_row))
            if next_row != final_count:
                raise ValueError("Số dòng Tableau sau bổ sung không khớp dữ liệu đã kiểm tra.")
        with ZipFile(source.path) as original, ZipFile(candidate) as result:
            if original.namelist() != result.namelist():
                raise ValueError("Cấu trúc file Tableau thay đổi ngoài sheet VH FTT.")
            for entry in original.infolist():
                if entry.filename != source.part and original.read(entry.filename) != result.read(entry.filename):
                    raise ValueError("Có thành phần ngoài sheet VH FTT bị thay đổi.")
        if metrics:
            months = sorted({key[1] for key in metrics})
            verified = read_source(candidate, months, source.cutoff)
            for key, expected in metrics.items():
                if key[2] not in PLANTS:
                    continue
                observed = verified.values.get(key, Counts(0, 0))
                if not (math.isclose(expected.inspected, observed.inspected, rel_tol=0, abs_tol=0.000001)
                        and math.isclose(expected.defects, observed.defects, rel_tol=0, abs_tol=0.000001)):
                    raise ValueError(f"Số {key[0]} · {key[1]} · {key[2]} sau bổ sung không khớp dữ liệu đã kiểm tra.")
        if fingerprint(source.path) != source.revision:
            raise ValueError("File Tableau vừa được cập nhật ở nơi khác. Kiểm tra đầu vào lại.")
        backups = APP_ROOT / "Report/InputBackups/Tableau"
        backups.mkdir(parents=True, exist_ok=True)
        backup = backups / f"{source.path.stem}-{source.revision[:12]}.xlsx"
        if not backup.exists():
            shutil.copyfile(source.path, backup)
        os.replace(candidate, source.path)
    save_coverage(source, windows, metrics)
    return dict(path=str(source.path), sheet=SHEET, added_rows=len(records), backup=str(backup))


def save_coverage(source, windows, metrics=None):
    if not windows:
        return
    coverage = dict(source.coverage)
    saved_metrics = dict(source.metrics)
    for key, counts in (metrics or {}).items():
        if key[2] in PLANTS:
            saved_metrics['|'.join(key)] = asdict(counts)
    for window in windows:
        metric = "HFPA" if window["group"] == "hfpa" else "FTT"
        key = "|".join((metric, window["date_from"][:7], window["plant"]))
        coverage[key] = max(coverage.get(key, ""), window["date_to"])
    from sfc_report_tools import write_json

    write_json(ledger_path(source.path), dict(sha256=fingerprint(source.path), coverage=coverage, metrics=saved_metrics))
