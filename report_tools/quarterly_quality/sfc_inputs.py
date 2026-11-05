"""Validate and aggregate source exports at the audit and FTT production grains."""

from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation

from report_validation import PreflightError, number, parse_date
from sfc_excel import read_export
from sfc_http import GROUPS as HFPA_FACTORY_CODES

from .quarterly_quality import BLOCKS, PLANTS, REGIONS, Counts, Quarter, quarter_months


FTT_LOT = ("Date", "FactoryGroup", "Factory", "Plant", "Line", "Model", "Model 2",
           "SO#", "Season", "Cutting Die", "Style Nbr", "Colorway Cd")
FTT_METRICS = ("AGrade", "Rework1", "Multi Rework", "B/Reject", "C/Reject")
HFPA_FIELDS = ("Audit Date (dd-mmm-yyyy)", "Factory Group Cd_UccId",
               "Internal AUDIT ID (YYMMDDNO+Plant+Line)", "Audit Sample Size", "Defective Q'ty(pair)")


class SourceRowError(ValueError):
    def __init__(self, message, row=None):
        super().__init__(message)
        self.row = row


def add_details(target, source):
    target[0].update(source[0])
    for model, issues in source[1].items():
        target[1][model].update(issues)
    if len(source) > 2:
        target[2].update(source[2])


def source_counts(path, item, details=None, *, approved=None, matched=None, export=None):
    try:
        header, rows = read_export(path) if export is None else export
    except Exception as error:
        raise SourceRowError("Không đọc được bảng Excel xuất từ SFC.") from error
    columns = [str(value or "").strip() for value in header]
    group, plant = item["group"], item["plant"]
    fields = HFPA_FIELDS if group == "hfpa" else (*FTT_LOT, "Total Pair Produced", "Station", "Metric", "Number of pairs")
    if any(columns.count(field) != 1 for field in fields):
        raise SourceRowError("Thiếu hoặc trùng cột bắt buộc của mẫu MES410 / FTT Internal.", 1)
    if group == "hfpa" and not {"SO#", "Country", "Pass/Fail"} <= set(columns):
        raise SourceRowError("Cần mẫu MES410 nội bộ, tắt NIKE khi xuất.", 1)
    indexes = [columns.index(field) for field in fields]
    model_col = next((columns.index(c) for c in ("Model 2", "Model") if c in columns), None)
    issue_cols = []
    if details is not None:
        if group == "hfpa" and "Defect Q'ty (Total Defect Q'ty)" in columns:
            start = columns.index("Defect Q'ty (Total Defect Q'ty)") + 1
            issue_cols = [(i, label) for i, label in enumerate(columns[start:], start) if label]
        elif group == "ftt" and {"Defect Issues", "Issues Q'ty"} <= set(columns):
            issue_cols = [columns.index("Defect Issues"), columns.index("Issues Q'ty")]
    bases, pairs = {}, {}
    daily_counts = defaultdict(lambda: [Decimal(0), Decimal(0)])
    daily_details = defaultdict(lambda: (Counter(), defaultdict(Counter)))
    inspected, defects = Decimal(0), Decimal(0)
    for row_number, raw in enumerate(rows, 2):
        row = tuple(raw[index] for index in indexes)
        try:
            day = parse_date(row[0]).isoformat()
        except ValueError as error:
            raise SourceRowError("Ngày kiểm tra không hợp lệ.", row_number) from error
        if not item["date_from"] <= day <= item["date_to"]:
            raise SourceRowError("Có ngày nằm ngoài khoảng đã yêu cầu.", row_number)
        factory = str(row[1 if group == "hfpa" else 2]).strip().upper()
        expected = HFPA_FACTORY_CODES[plant] if group == "hfpa" else {plant}
        if factory not in expected:
            raise SourceRowError("Mã nhà máy trong file khác nhà máy được yêu cầu.", row_number)
        try:
            if group == "hfpa":
                audit, size, defective = row[2:]
                audit = str(audit or "").strip().upper()
                if not audit or audit in bases:
                    raise SourceRowError("Thiếu hoặc trùng Internal AUDIT ID.", row_number)
                sample, bad = number(size), number(defective)
                if bad > sample:
                    raise SourceRowError("Số đôi lỗi lớn hơn số đôi lấy mẫu.", row_number)
                bases[audit] = sample
                inspected += sample
                defects += bad
                if approved is not None:
                    daily_counts[day][0] += sample
                    daily_counts[day][1] += bad
                if details is not None and model_col is not None:
                    model = str(raw[model_col] or "").strip()
                    if model:
                        target = daily_details[day] if approved is not None else details
                        target[0][model] += bad
                        for i, label in issue_cols:
                            if raw[i]:
                                target[1][model][label] += number(raw[i])
            else:
                produced, station, metric, quantity = row[len(FTT_LOT):]
                if str(station).strip().casefold() != "assembly":
                    continue
                metric = str(metric).strip()
                if metric not in FTT_METRICS:
                    raise SourceRowError("Metric ở trạm Assembly nằm ngoài nhóm AGrade / Rework / B-C Reject.", row_number)
                lot = (day, *row[1:len(FTT_LOT)])
                total, count = number(produced), number(quantity)
                if bases.setdefault(lot, total) != total or pairs.setdefault((lot, metric), count) != count:
                    raise SourceRowError("Số đôi khác nhau giữa các dòng chi tiết của cùng nhóm sản xuất.", row_number)
                if issue_cols and metric != "AGrade":
                    label, quantity = (raw[i] for i in issue_cols)
                    model = str(lot[6] or lot[5] or "").strip()
                    if model and label and quantity:
                        details[1][model][str(label).strip()] += number(quantity)
        except InvalidOperation as error:
            raise SourceRowError("Số lượng cần là số hợp lệ, không âm.", row_number) from error
        except ValueError as error:
            if isinstance(error, SourceRowError):
                raise
            raise SourceRowError("Số lượng cần là số hợp lệ, không âm.", row_number) from error
    if group == "ftt":
        if rows and not bases:
            raise SourceRowError("File FTT Internal không có dữ liệu trạm Assembly.")
        inspected = sum(bases.values(), Decimal(0))
        grades = {metric: Decimal(0) for metric in FTT_METRICS}
        for (_, metric), count in pairs.items():
            grades[metric] += count
        if details is not None:
            for (lot, metric), count in pairs.items():
                model = str(lot[6] or lot[5] or "").strip()
                if model and metric != "AGrade":
                    details[0][model] += count
        rejects = grades["B/Reject"] + grades["C/Reject"]
        if abs(inspected - grades["AGrade"] - rejects) > Decimal("0.000001"):
            raise SourceRowError("Tổng đôi sản xuất không khớp AGrade + B/Reject + C/Reject.")
        defects = grades["Rework1"] + grades["Multi Rework"] + rejects
        if defects > inspected:
            raise SourceRowError("Tổng rework và reject lớn hơn tổng đôi sản xuất.")
        return Counts(float(inspected), float(defects), float(grades["AGrade"]),
                      float(grades["Rework1"] + grades["Multi Rework"])), len(rows)
    if approved is not None and details is not None:
        match_dates = matched if matched is not None else []
        accumulated = [Decimal(0), Decimal(0)]
        for day, (sample, bad) in sorted(daily_counts.items()):
            accumulated[0] += sample
            accumulated[1] += bad
            if tuple(accumulated) == (Decimal(str(approved.inspected)), Decimal(str(approved.defects))):
                match_dates.append(day)
        if match_dates:
            for day, day_details in daily_details.items():
                if day <= match_dates[-1]:
                    add_details(details, day_details)
    return Counts(float(inspected), float(defects)), len(rows)


def corrected_hfpa_counts(mes_path, item, qa_path, details, *, tableau_rows=None):
    """Use the monthly HFPA engine on the same factory and date window for both sources."""
    import pandas as pd
    from report_tools.hfpa.agents.reconciliation_engine_agent import ReconciliationEngineAgent

    try:
        mes_export = read_export(mes_path)
        header, rows = read_export(qa_path)
    except Exception as error:
        raise SourceRowError(f"Không đọc được cặp QAStation {qa_path.name} và MES410.") from error
    mes_total, mes_rows = source_counts(mes_path, item, export=mes_export)
    mes_header = [str(value or "").strip() for value in mes_export[0]]
    if any(mes_header.count(name) != 1 for name in ("Plant", "Line", "Model 2")):
        raise SourceRowError(f"{mes_path.name}: thiếu xưởng, chuyền hoặc model để đối chiếu QAStation.", 1)
    columns = [str(value or "").strip() for value in header]
    required = ("InspDate", "Line", "ShoeName", "InspQty", "FailQty", "Station", "Issues", "IssueQty")
    if any(columns.count(name) != 1 for name in required):
        raise SourceRowError(f"{qa_path.name}: thiếu hoặc trùng cột QAStation HFPA.", 1)
    qa = pd.DataFrame(rows, columns=columns)
    if "FTY" in columns:
        factories = qa["FTY"].astype(str).str.strip().str.upper()
        if not factories.isin({item["plant"], *HFPA_FACTORY_CODES[item["plant"]]}).all():
            raise SourceRowError(f"{qa_path.name}: mã nhà máy trong QAStation không khớp.")
    qa = qa[list(required)].copy()
    try:
        qa["InspDate"] = qa["InspDate"].map(lambda value: parse_date(value).isoformat())
        for col in ("InspQty", "FailQty", "IssueQty"):
            qa[col] = qa[col].map(lambda value: float(number(value)) if pd.notna(value) and value != "" else 0)
    except (InvalidOperation, ValueError) as error:
        raise SourceRowError(f"{qa_path.name}: ngày hoặc số lượng QAStation không hợp lệ.") from error
    if (not qa["InspDate"].between(item["date_from"], item["date_to"]).all()
            or not qa["Station"].astype(str).str.strip().str.upper().eq("HFPA").all()
            or qa["FailQty"].gt(qa["InspQty"]).any()):
        raise SourceRowError(f"{qa_path.name}: sai kỳ, trạm hoặc số đôi lỗi vượt số đôi kiểm.")
    if qa.empty:
        if mes_total.inspected:
            raise SourceRowError(f"{qa_path.name}: QAStation trống trong khi MES410 có số đôi kiểm.")
        return Counts(0, 0), mes_rows, len(rows), []
    qa["Station"] = "HFPA"
    qa["FTY"] = item["plant"]
    mes = pd.DataFrame(mes_export[1], columns=mes_header)
    mes["FTY"] = item["plant"]
    for col in ("Audit Sample Size", "Defective Q'ty(pair)"):
        mes[col] = mes[col].map(lambda value: float(number(value)))
    total_issues = "Defect Q'ty (Total Defect Q'ty)"
    if total_issues in mes:
        mes[total_issues] = pd.to_numeric(mes[total_issues].astype("string").str.replace(",", "", regex=False),
                                         errors="coerce")

    corrected, _, audit = ReconciliationEngineAgent.reconcile_and_validate(qa, mes)
    unresolved = [row for row in audit if row["Status"] != "MES410_Corrected"]
    if unresolved:
        first = unresolved[0]
        model = first["ShoeName_Std"] or "(trống)"
        raise SourceRowError(
            f"{qa_path.name}: còn {len(unresolved)} nhóm HFPA chưa đủ dữ liệu để đối chiếu và hiệu chỉnh theo SOP. "
            f"Nhóm đầu tiên {first['Date_Std']} · {first['FTY']} · xưởng {first['Plant_Std']} · "
            f"chuyền {first['Line_Std']} · model {model} · {first['Status']}."
        )
    if tableau_rows is not None:
        for _, row in corrected.iterrows():
            tableau_rows.append([str(row['Date_Std'])[:10], item['plant'], row['Line'], row['ShoeName_Std'],
                                 float(row['InspQty_Validated']), float(row['FailQty_Validated']), 'HFPA',
                                 row['Issues_Std'], float(row['IssueQty_Validated'])])
    pivot = ReconciliationEngineAgent.inspection_pivot(corrected)
    total = Counts(float(pivot["InspQty"].sum()), float(pivot["FailQty"].sum()))
    if details is None:
        return total, mes_rows, len(rows), audit

    details[0].update(pivot.groupby("Shoename")["FailQty"].sum().to_dict())
    if len(details) > 2:
        details[2].update(pivot.groupby("Shoename")["InspQty"].sum().to_dict())
    issues = corrected[(corrected["Issues_Std"] != "") & (corrected["IssueQty_Validated"] > 0)]
    for model, model_issues in issues.groupby("ShoeName_Std", sort=False):
        details[1][model].update(model_issues.groupby("Issues_Std", sort=False)["IssueQty_Validated"].sum().to_dict())
    return total, mes_rows, len(rows), audit


def load_sfc_quarter(workspace, run, stage, *, baseline=None):
    from sfc_report_tools import TOOLS

    windows = run["acquisition"]
    months = quarter_months(run["period"])
    prior, current = months[:3], months[3:]
    covered = tuple(month for month in current if month <= windows["date_to"][:7])
    reports, values, locations = [], {}, {}
    details_by_region = defaultdict(lambda: (Counter(), defaultdict(Counter), Counter()))
    tableau = None
    tableau_rows = []
    if windows.get('tableau'):
        from .tableau_source import read_source

        reference = windows['tableau']
        tableau = read_source(reference['path'], current, windows['date_to'])
        if tableau.revision != reference['sha256']:
            raise SourceRowError('Nguồn Tableau đã thay đổi từ lúc kiểm tra. Kiểm tra đầu vào lại.')
        reports.append(dict(name=tableau.path.name, group='workbook', rows=tableau.row_count,
                            sheets=[dict(name='VH FTT', header_row=1, rows=tableau.row_count)],
                            valid=True, error_count=0, issues=[]))
        values.update(tableau.values)
        locations.update({key: f'{tableau.path.name} · VH FTT' for key in tableau.values})
        for (metric, month, plant), details in tableau.details.items():
            region = 'VN' if plant in REGIONS['VN'] else 'Indo'
            add_details(details_by_region[metric, region], details)
    run["validation"] = reports
    current_inputs = [item for item in run["inputs"] if item["date_from"][:7] in covered]
    window_key = lambda item: (item["plant"], item["date_from"], item["date_to"])
    qa_inputs = [item for item in current_inputs if item["group"] == "qastation"]
    qa_by_window = {window_key(item): item for item in qa_inputs}
    hfpa_windows = {window_key(item) for item in current_inputs if item["group"] == "hfpa"}
    if len(qa_by_window) != len(qa_inputs) or qa_by_window.keys() != hfpa_windows:
        raise SourceRowError("Cần đúng một file QAStation HFPA cho mỗi file MES410 cùng nhà máy và khoảng ngày.")
    inputs = [item for item in current_inputs if item["group"] != "qastation"]
    run["hfpa_model_review"] = []
    run["hfpa_baseline_differences"] = []
    run["ftt_baseline_differences"] = []
    seen_inputs = set()
    for index, item in enumerate(inputs, 1):
        name = item["name"]
        stage(f'Kiểm tra và cộng số lượng nguồn SFC {index}/{len(inputs)} · {name}', "validation")
        group = item["group"]
        path = workspace / TOOLS["quarterly_quality"]["groups"][group][1] / name
        report = dict(name=name, group=group, sheets=[], rows=0, error_count=0, issues=[], valid=True)
        reports.append(report)
        try:
            key = group.upper(), item["date_from"][:7], item["plant"]
            if key in seen_inputs:
                raise SourceRowError("Trùng chỉ số, tháng và nhà máy trong đầu vào.")
            seen_inputs.add(key)
            region = "VN" if item["plant"] in REGIONS["VN"] else "Indo"
            rank_key = key[0], region
            details = (Counter(), defaultdict(Counter), Counter()) if group == "hfpa" else (Counter(), defaultdict(Counter))
            template_counts = baseline.values.get(key) if baseline else None
            if group == "hfpa":
                qa_name = qa_by_window[window_key(item)]["name"]
                qa_path = workspace / TOOLS["quarterly_quality"]["groups"]["qastation"][1] / qa_name
                arguments = dict(tableau_rows=tableau_rows) if tableau else {}
                raw, report["rows"], qa_rows, audit = corrected_hfpa_counts(path, item, qa_path, details, **arguments)
                reports.append(dict(name=qa_name, group="qastation", sheets=[], rows=qa_rows,
                                    valid=True, error_count=0, issues=[]))
                run["hfpa_model_review"].extend(
                    {key: r[key] for key in ("FTY", "Date_Std", "Plant_Std", "Line_Std", "ShoeName_Std", "Status")}
                    for r in audit if r["Status"] != "MES410_Corrected"
                )
            else:
                if tableau:
                    from .tableau_source import normalize_ftt

                    export = read_export(path)
                    raw, report["rows"] = source_counts(path, item, details, export=export)
                    tableau_rows.extend(normalize_ftt(export, item))
                else:
                    raw, report["rows"] = source_counts(path, item, details)
            stored = values.get(key)
            values[key] = (Counts(
                stored.inspected + raw.inspected, stored.defects + raw.defects,
                stored.grade + raw.grade if stored.grade is not None and raw.grade is not None else None,
                stored.rework + raw.rework if stored.rework is not None and raw.rework is not None else None,
            ) if stored else raw)
            if template_counts is not None and (template_counts.inspected, template_counts.defects) != (raw.inspected, raw.defects):
                run[f"{group}_baseline_differences"].append(dict(
                    month=key[1], plant=key[2], template_insp=template_counts.inspected,
                    template_fail=template_counts.defects, corrected_insp=raw.inspected,
                    corrected_fail=raw.defects,
                ))
            add_details(details_by_region[rank_key], details)
            locations[key] = name
        except SourceRowError as error:
            report.update(valid=False, error_count=1, issues=[dict(row=error.row, message=str(error))])
    if any(not report["valid"] for report in reports):
        raise PreflightError(reports)
    for metric in BLOCKS:
        for month in covered:
            for scope, plants in REGIONS.items():
                records = [values.get((metric, month, plant)) for plant in plants]
                if any(record is None for record in records):
                    raise SourceRowError(f"Thiếu file {metric} của một nhà máy trong tháng {month}.")
                values[metric, month, scope] = Counts(sum(record.inspected for record in records),
                                                      sum(record.defects for record in records))
    quarter = Quarter(run["period"], prior, current, covered, values, locations, {},
                      date_from=windows["date_from"], date_to=windows["date_to"])
    for (metric, region), (models, issues, samples) in details_by_region.items():
        ranked = (sorted(models.items(), key=lambda row: (-row[1], -samples[row[0]], row[0]))
                  if metric == "HFPA" else models.most_common())
        top = [(label, float(quantity), quarter.totals(metric, region).inspected)
               for label, quantity in ranked[:3] if quantity]
        defects = issues[top[0][0]] if top else Counter()
        total = float(defects.total())
        quarter.rankings[metric, region] = (top, [(label, float(quantity), total)
                                                  for label, quantity in defects.most_common(3) if quantity])
    if tableau:
        from .tableau_source import append_records

        run['tableau_update'] = append_records(tableau, tableau_rows, inputs, stage, metrics=values)
    return quarter
