"""Resolve Dashboard files independently of the legacy MQAA program."""
import json
import os
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent


def mqaa_workspace():
    config_path = APP_ROOT / 'dashboard_config.json'
    config = json.loads(config_path.read_text(encoding='utf-8')) if config_path.exists() else {}
    configured = os.environ.get('CHINGLUH_MQAA_ROOT') or config.get('legacy_mqaa_root', '../MQAA-Record')
    path = Path(configured).expanduser()
    return (path if path.is_absolute() else APP_ROOT / path).resolve()


def mqaa_master_file():
    root = mqaa_workspace()
    for name in ('MQAA_Master_DoiChieu_21_Forms_MoiNhat.xlsx', 'MQAA_Master_DoiChieu_21_Forms.xlsx'):
        path = root / name
        if path.is_file():
            return path
    return None
