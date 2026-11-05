"""Explain existing report checks in Vietnamese without changing business rules."""

from collections import Counter
from decimal import Decimal, InvalidOperation
import re


def explain_copq_error(error):
    """Translate evidence from the engine; thresholds request review, not correction."""
    if re.fullmatch(r"Ttl cost \(\$\) > \d+(?:\.\d+)?", error):
        limit = error.split(" > ")[1]
        return ("cost", "Chi phí vượt ngưỡng cảnh báo",
                f"Chi phí mỗi dòng vượt {limit} USD theo quy tắc COPQ.",
                "Đối chiếu chứng từ và chi phí trên SFC. Giá trị vượt ngưỡng có thể hợp lệ.")
    hours = re.fullmatch(r"Working hours > ([\d.]+) & Defective Qty\(Pair\) < ([\d.]+)", error)
    if hours:
        return ("hours:" + error, "Giờ làm cao so với số lượng",
                f"Giờ làm lớn hơn {hours[1]} và số lượng nhỏ hơn {hours[2]} đôi.",
                "Xác minh giờ làm, đơn vị tính và số lượng thực tế trước khi kết luận nhập sai.")
    rules = {
        "Abnormal inspection rate (< 25 pairs/h)": (
            "Năng suất cần xác minh", "Năng suất dưới 25 đôi/giờ trong thời gian làm trên 3 giờ.",
            "Đối chiếu giờ làm và số lượng thực tế."),
        "Invalid style nbr => cannot link model": (
            "Chưa liên kết được Model từ Style Nbr", "Tool chưa tìm được Model tương ứng với Style Nbr.",
            "Đối chiếu Style Nbr và dữ liệu liên kết Model; chưa đủ căn cứ kết luận mã bị nhập sai."),
        "Missing style nbr => cannot link model": (
            "Thiếu Style Nbr để liên kết Model", "Không có Style Nbr để tìm Model.",
            "Bổ sung Style Nbr đúng nguồn rồi đối chiếu Model."),
        "Invalid mold code => cannot link model": (
            "Chưa liên kết được Model từ Mold", "Tool chưa tìm được Model tương ứng với mã khuôn.",
            "Đối chiếu mã khuôn và dữ liệu liên kết Model."),
        "Missing mold code => cannot link model": (
            "Thiếu mã khuôn để liên kết Model", "Không có Mold để tìm Model.",
            "Bổ sung mã khuôn đúng nguồn rồi đối chiếu Model."),
        "Missing model => cannot link model": (
            "Thiếu Model", "Dòng Reinspection chưa có Model.", "Bổ sung Model từ nguồn đã xác minh."),
        "Missing data": (
            "Thiếu thông tin nghiệp vụ", "Model và các chỉ tiêu số lượng, giờ làm, chi phí, nhân lực đều trống hoặc bằng 0.",
            "Đối chiếu bản ghi gốc, phân biệt dữ liệu thiếu với giá trị 0 hợp lệ."),
        "Wrong Working hours => Wrong Ttl cost ($)": (
            "Có số lượng nhưng chưa có giờ làm", "Reinspection có số lượng, nhưng giờ làm trống hoặc bằng 0.",
            "Xác minh giờ làm và cách tính chi phí."),
        "Wrong Defective Qty(Pair)": (
            "Có giờ làm nhưng chưa có số lượng", "Reinspection có giờ làm, nhưng số lượng trống hoặc bằng 0.",
            "Xác minh số lượng trên bản ghi gốc."),
        "Working hours present but zero Ttl cost ($)": (
            "Có giờ làm nhưng chi phí bằng 0", "Reinspection có giờ làm, nhưng chi phí trống hoặc bằng 0.",
            "Đối chiếu dữ liệu tính chi phí."),
        "B/C Grade missing Defective Qty(Pair)": (
            "B/C có chi phí nhưng thiếu số lượng", "Dòng B/C có chi phí, nhưng số lượng trống hoặc bằng 0.",
            "Đối chiếu số lượng B/C tại nguồn."),
        "B/C Grade missing Ttl cost ($)": (
            "B/C có số lượng nhưng thiếu chi phí", "Dòng B/C có số lượng, nhưng chi phí trống hoặc bằng 0.",
            "Đối chiếu giá và cách tính chi phí B/C."),
        "SGS Retest Qty recorded without unit price": (
            "SGS có số lượng nhưng thiếu đơn giá", "Số lượng SGS Retest có giá trị nhưng đơn giá trống hoặc bằng 0.",
            "Đối chiếu số lượng và đơn giá SGS."),
        "SGS unit price recorded without Retest Qty": (
            "SGS có đơn giá nhưng thiếu số lượng", "Đơn giá SGS có giá trị nhưng số lượng trống hoặc bằng 0.",
            "Đối chiếu số lượng và đơn giá SGS."),
        "Blank Type => Blank Ttl cost ($)": (
            "Thiếu Type và chi phí", "Type đang trống, chi phí trống hoặc bằng 0.",
            "Bổ sung loại nghiệp vụ và xác minh chi phí từ nguồn."),
        "Type blank": ("Thiếu Type", "Chưa có loại nghiệp vụ COPQ.", "Bổ sung Type từ nguồn."),
        "Category blank": ("Thiếu Category", "Chưa có phân loại COPQ.", "Bổ sung Category từ nguồn."),
    }
    if error in rules:
        return (error, *rules[error])
    if error.startswith("Negative ") and error[9:] in ("Ttl cost ($)", "Defective Qty(Pair)", "Working hours"):
        field = error[9:]
        return (error, "Giá trị âm ở " + field, "Tool phát hiện giá trị âm tại " + field + ".",
                "Đối chiếu bản ghi gốc và quy tắc xử lý giá trị âm.")
    return ("other", "Quy tắc COPQ cần xem lại", "Tool gắn cờ nhưng chưa có diễn giải riêng cho quy tắc này.",
            "Xem cột Validation Errors tại dòng Excel được chỉ ra.")


def _numeric_samples(values, columns):
    result = {}
    for column in columns:
        raw = values.get(column)
        if raw is None or raw == "":
            continue
        try:
            number = Decimal(str(raw))
            if number.is_finite():
                result[column] = str(number)
        except InvalidOperation:
            pass
    return result


def _add_group(groups, key, entry, row, site, kind, values, numeric_columns=()):
    if key not in groups:
        groups[key] = dict(entry, count=0, sites=Counter(), types=Counter(), rows=[], examples=[])
    item = groups[key]
    item["count"] += 1
    item["sites"][str(site or "Chưa rõ nhà máy")[:80]] += 1
    if kind:
        item["types"][str(kind)[:80]] += 1
    if len(item["rows"]) < 12:
        item["rows"].append(row)
    metrics = _numeric_samples(values, numeric_columns)
    if metrics and len(item["examples"]) < 3:
        item["examples"].append(dict(row=row, values=metrics))


def _finish_groups(groups):
    result = []
    for item in groups.values():
        item["sites"] = dict(item["sites"])
        item["types"] = dict(item["types"])
        item["rows_truncated"] = item["count"] > len(item["rows"])
        result.append(item)
    return sorted(result, key=lambda item: -item["count"])


def copq_issues(header, records, filename):
    """records contains (actual Excel row, values), not compressed row offsets."""
    groups = {}
    for row_number, row in records:
        values = dict(zip(header, row))
        if values.get("Data Quality Status") in ("OK", "Auto-Resolved"):
            continue
        missing = tuple(sorted(set(str(values.get("Missing Fields") or "").split(", ")) - {""}))
        raw_errors = str(values.get("Validation Errors") or "").split(", ")
        errors = [explain_copq_error(error) for error in raw_errors if error and not error.startswith("Auto-resolved")]
        # Multiple issues on one row form one group, so counts do not double-count rows.
        errors = list({error[0]: error for error in errors}.values())
        titles, reasons, actions = [], [], []
        if missing:
            titles.append("Thiếu " + ", ".join(missing))
            reasons.append("Các trường bắt buộc đang trống: " + ", ".join(missing) + ".")
            actions.append("Bổ sung thông tin đúng từ nguồn rồi chạy lại báo cáo.")
            if "Factory" in missing:
                reasons.append("Factory là xưởng/bộ phận; cột Site chỉ xác định nhà máy và không thay thế Factory.")
        for _, title, reason, action in errors:
            titles.append(title)
            reasons.append(reason)
            actions.append(action)
        if not titles:
            _, title, reason, action = explain_copq_error("")
            titles, reasons, actions = [title], [reason], [action]
        numeric_columns = []
        if any(error[0] == "cost" for error in errors):
            numeric_columns.append("Ttl cost ($)")
        if any(error[0].startswith("hours:") or error[0].startswith("Abnormal inspection") for error in errors):
            numeric_columns += ["Working hours", "Defective Qty(Pair)"]
        entry = dict(title="; ".join(titles), reason=" ".join(reasons), action=" ".join(dict.fromkeys(actions)),
                     severity="incomplete" if missing else "verify", file=filename, sheet="COPQ_Clean", unit="dòng")
        key = (missing, tuple(sorted(error[0] for error in errors)))
        _add_group(groups, key, entry, row_number, values.get("Site"), values.get("Type"), values, numeric_columns)
    return _finish_groups(groups)


def hfpa_issues(header, records, filename):
    explanations = {
        "HFPA_Only": ("Lot chỉ có trong MES410", "Chưa tìm thấy lot tương ứng trong QAStation.",
                      "Đối chiếu kỳ, nhà máy, plant, chuyền và nguồn QAStation."),
        "QAStation_Only": ("Lot chỉ có trong QAStation", "Chưa tìm thấy lot tương ứng trong MES410.",
                           "Đối chiếu kỳ, nhà máy, plant, chuyền và nguồn MES410."),
        "Different_Count": ("Số lượng lỗi giữa hai nguồn chưa khớp", "Lot có ở hai nguồn nhưng tổng số lượng lỗi khác nhau.",
                            "Đối chiếu số lỗi QAStation với MES410 tại các dòng được chỉ ra."),
    }
    groups = {}
    for row_number, row in records:
        values = dict(zip(header, row))
        status = values.get("Validation_Status")
        if status == "Matched_Exact":
            continue
        title, reason, action = explanations.get(status, (
            "Lot HFPA cần đối chiếu", "Lot chưa được xác nhận khớp chính xác giữa hai nguồn.",
            "Xem Validation_Status và dữ liệu hai nguồn tại dòng được chỉ ra."))
        entry = dict(title=title, reason=reason, action=action, severity="reconcile", file=filename,
                     sheet="Lot_Defect_Variance", unit="lot")
        _add_group(groups, status, entry, row_number, values.get("FTY"), None, values,
                   ("QA_IssueQty_Sum", "HFPA_Defect_Qty"))
    return _finish_groups(groups)


def summarize_checks(checks):
    warnings = [check for check in checks if check["status"] == "warning"]
    issues = []
    for check in warnings:
        if check.get("issues"):
            issues.extend(check["issues"])
            continue
        rule = check["rule"]
        reason = "Kiểm tra chưa đạt; xem chi tiết đối chiếu bên dưới."
        if "expected" in check and "actual" in check:
            reason = f"Giá trị trước xử lý: {check['expected']}; sau xử lý: {check['actual']}."
        elif "mismatches" in check:
            reason = f"Có {check['mismatches']} bản ghi chưa đạt quy tắc. Lượt cũ chưa lưu phân nhóm nguyên nhân."
        issues.append(dict(title=check["message"], reason=reason,
                           action="Đối chiếu dữ liệu nguồn và kết quả của bước này trước khi sử dụng báo cáo.",
                           severity="reconcile", count=check.get("mismatches"), unit="bản ghi"))
    flagged = sum(check.get("mismatches", 0) for check in warnings if check["rule"] in ("copq.quality", "hfpa.lot_variance"))
    headline = f"Có {len(issues)} nhóm vấn đề cần xem lại." if warnings else "Các kiểm tra hiện có đều đạt."
    if not checks:
        headline = "Chưa có kết quả đối chiếu để tóm tắt."
    if flagged:
        headline = f"{flagged:,} bản ghi cần xem lại · {len(issues)} nhóm vấn đề."
    verified = []
    for rules, message in [
        (("copq.row_count", "copq.cost"), "Số bản ghi và tổng chi phí COPQ đối chiếu khớp."),
        (("ftt.bc_rows", "ftt.bc_quantity"), "Số dòng và tổng số lượng FTT B/C đối chiếu khớp."),
    ]:
        if all(any(c["rule"] == rule for c in checks) and
               all(c["status"] == "passed" for c in checks if c["rule"] == rule) for rule in rules):
            verified.append(message)
    return dict(version=1, status="warning" if warnings else "passed", headline=headline,
                flagged_records=flagged, issues=issues, verified=verified)
