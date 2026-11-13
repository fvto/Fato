"""Durable step events, safe cooperative cancellation and value-free diagnostics."""

from pathlib import Path
import traceback
import re


class Cancelled(Exception):
    pass


def diagnostics(error):
    return dict(
        code=type(error).__name__,
        frames=[
            dict(file=Path(frame.filename).name, line=frame.lineno, function=frame.name)
            for frame in traceback.extract_tb(error.__traceback__)[-8:]
        ],
    )


def friendly_error(error):
    from report_validation import PreflightError

    if getattr(error, 'code', None) == 'sfc.dom_export_pending':
        return error.user_message

    if isinstance(error, PreflightError):
        detail = next((issue["message"] for report in error.reports for issue in report.get("issues", [])
                       if issue.get("message")), "")
        return f"{error} {detail}".strip()
    if isinstance(error, PermissionError):
        return (
            "Không ghi/đọc được file. Đóng Excel/PowerPoint đang mở file rồi thử lại."
        )
    if (
        isinstance(error, (TimeoutError, ConnectionError))
        or "timeout" in str(error).lower()
    ):
        duration = re.search(r"\b(\d{1,3}) phút\b", str(error))
        suffix = f" sau {duration[1]} phút" if duration else ""
        return (
            "Timeout: kết nối quá thời gian hoặc bị ngắt"
            + suffix
            + ". Kiểm tra phiên SFC; các file đã hoàn tất được giữ lại."
        )
    return "Bước xử lý chưa hoàn tất. Kiểm tra đầu vào và mục chẩn đoán kỹ thuật; có thể tải các kết quả đã hoàn tất."


class Monitor:
    def __init__(self, folder, run):
        self.folder, self.run = Path(folder), run
        self.active = None
        run.setdefault("steps", [])

    def check_cancel(self):
        if (self.folder / "cancel.request").exists():
            raise Cancelled(
                "Đã hủy tại điểm an toàn. Các file đã hoàn tất được giữ lại."
            )

    def save(self):
        from sfc_report_tools import write_json

        write_json(self.folder / "run.json", self.run)

    def stage(self, message, key=None):
        from sfc_report_tools import now

        self.check_cancel()
        # Messages without a key (e.g. acquisition details) update the active step.
        if key and key != self.active:
            if self.active:
                self.finish_step("done")
            self.active = key
            step = next((s for s in self.run["steps"] if s["id"] == key), None)
            if step is None:
                step = dict(id=key, title=message, status="pending")
                self.run["steps"].append(step)
            step.update(status="running", started_at=now())
        self.run["stage"] = message
        self.save()

    def finish_step(self, status):
        from sfc_report_tools import now

        if self.active:
            step = next(s for s in self.run["steps"] if s["id"] == self.active)
            step.update(status=status, finished_at=now())
            self.active = None


class QuietEngineOutput:
    """Legacy engines may print rows, paths or identifiers. Persist no raw output."""

    encoding = "utf-8"

    def write(self, value):
        return len(value)

    def flush(self):
        pass

    def reconfigure(self, **kwargs):
        pass
