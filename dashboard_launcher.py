"""Start the local dashboard, optionally with the SFC DOM connection."""
import importlib.util
import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser

ROOT = Path(__file__).resolve().parent
DASHBOARD = "http://localhost:5050"
CDP = "http://localhost:9222"
SFC = "https://sfc.chingluh.com/vhmes/QA/QCHFPA.aspx?tree_id=Q504"


def read_json(url):
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    with opener.open(url, timeout=2) as response:
        return json.load(response)


def dashboard_ready():
    try:
        data = read_json(DASHBOARD + "/api/health")
        return (data.get("service") == "sfc-dashboard" and data.get("ready") is True
                and bool(data.get('app_root')) and Path(data['app_root']).resolve() == ROOT)
    except Exception:
        return False


def dom_ready():
    try:
        return bool(read_json(CDP + "/json/version").get("webSocketDebuggerUrl"))
    except Exception:
        return False


def wait_until(check, seconds):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if check():
            return True
        time.sleep(0.4)
    return False


def port_open(port):
    try:
        with socket.create_connection(("localhost", port), timeout=1):
            return True
    except OSError:
        return False


def start_server():
    if dashboard_ready():
        print("[OK] Dashboard Server da chay tai " + DASHBOARD, flush=True)
        return
    if port_open(5050):
        raise RuntimeError("Cong 5050 dang duoc su dung. Hay dong server cu roi mo lai MO_SERVER.bat.")
    missing = [name for name in ("flask", "playwright", "openpyxl", "xlrd")
               if importlib.util.find_spec(name) is None]
    if missing:
        raise RuntimeError("Python thieu thu vien: " + ", ".join(missing)
                           + ". Cai bang: python -m pip install flask playwright openpyxl xlrd")
    log_path = ROOT / "scratch" / "dashboard_server.log"
    log_path.parent.mkdir(exist_ok=True)
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    with log_path.open("ab") as log:
        process = subprocess.Popen(
            [sys.executable, "-u", str(ROOT / "sfc_dashboard_server.py")],
            cwd=ROOT, stdout=log, stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL, creationflags=flags,
        )
    if not wait_until(dashboard_ready, 25):
        raise RuntimeError(f"Khong khoi dong duoc server. Xem {log_path} (PID {process.pid}).")
    print("[OK] Dashboard Server da san sang: " + DASHBOARD, flush=True)


def start_dom():
    if dom_ready():
        print("[OK] Da co ket noi DOM cong 9222; giu nguyen trinh duyet hien tai.", flush=True)
        return
    if port_open(9222):
        raise RuntimeError("Cong 9222 dang ban nhung khong phai ket noi DOM hop le.")
    candidates = [
        Path(os.environ.get("PROGRAMFILES", "C:/Program Files")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("PROGRAMFILES", "C:/Program Files")) / "Microsoft/Edge/Application/msedge.exe",
    ]
    executable = next((path for path in candidates if path.is_file()), None)
    if executable is None:
        raise RuntimeError("Khong tim thay Google Chrome hoac Microsoft Edge.")
    profile = (Path(os.environ.get("LOCALAPPDATA", str(ROOT))) / "Google/Chrome/MQAA_Automation"
               if executable.name == "chrome.exe" else ROOT / "edge_auto_profile")
    subprocess.Popen([
        str(executable), "--remote-debugging-port=9222", "--remote-debugging-address=127.0.0.1",
        f"--user-data-dir={profile}", SFC,
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not wait_until(dom_ready, 25):
        raise RuntimeError("Chua ket noi duoc DOM. Hay dong cua so MQAA_Automation roi chon lai option 2.")
    print("[OK] Da mo SFC va ket noi DOM. Dang nhap SFC neu duoc yeu cau.", flush=True)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "server"
    if mode not in {"server", "dashboard", "dom"}:
        raise RuntimeError("Lua chon khong hop le: " + mode)
    start_server()
    if mode == "dom":
        start_dom()
    if mode != "server":
        if not webbrowser.open(DASHBOARD):
            raise RuntimeError("Hay mo trinh duyet tai " + DASHBOARD)
        print("[OK] Da mo Dashboard.", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print("[LOI] " + str(error), flush=True)
        sys.exit(1)
