"""List (default) or remove expired terminal runs created by the new monitor only."""

import argparse
from datetime import datetime, timedelta
import shutil

import sfc_report_tools as reports


def expired_runs(days=30):
    if days < 1:
        raise ValueError("Retention phải ít nhất một ngày.")
    cutoff = datetime.fromisoformat(reports.now()) - timedelta(days=days)
    for path in reports.RUN_ROOT.glob("*/run.json"):
        try:
            run = reports.read_json(path)
            if (
                run.get("monitor_version") == 1
                and run.get("status") in ("done", "error", "cancelled")
                and run.get("finished_at")
                and datetime.fromisoformat(run["finished_at"]) < cutoff
                and not reports.process_alive(run.get("pid"))
            ):
                yield run["id"]
        except (ValueError, KeyError, OSError):
            continue


def cleanup(days=30, apply=False):
    ids = list(expired_runs(days))
    if apply:
        with reports.LOCK:
            for run_id in ids:
                folder = reports.run_folder(run_id).resolve()
                # Verify both the strict id and final resolved target before recursive removal.
                if folder.parent != reports.RUN_ROOT.resolve():
                    raise ValueError("Thư mục retention nằm ngoài ToolRuns.")
                current = reports.read_json(folder / "run.json")
                if current.get("status") in (
                    "done",
                    "error",
                    "cancelled",
                ) and not reports.process_alive(current.get("pid")):
                    shutil.rmtree(folder)
    return ids


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=30)
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Xóa cả đầu vào, workspace và kết quả hết hạn; mặc định chỉ liệt kê.",
    )
    args = parser.parse_args()
    print({"apply": args.apply, "expired_ids": cleanup(args.days, args.apply)})
