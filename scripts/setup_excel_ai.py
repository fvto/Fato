"""Install the pinned Excel MCP runtime without changing Dashboard dependencies."""
from pathlib import Path
import shutil
import subprocess
import sys

root = Path(__file__).resolve().parents[1]
runtime = root / '.runtime' / 'excel-ai'
python = runtime / ('Scripts/python.exe' if sys.platform == 'win32' else 'bin/python')
uv = shutil.which('uv') or str(Path.home() / 'AppData/Local/hermes/bin/uv.exe')
if Path(uv).exists() or shutil.which(uv):
    if not python.exists():
        subprocess.run([uv, 'venv', '--python', sys.executable, str(runtime)], check=True)
    subprocess.run([uv, 'pip', 'install', '--python', str(python), '-r', str(root / 'requirements-excel-ai.txt')], check=True)
else:
    if not python.exists():
        subprocess.run([sys.executable, '-m', 'venv', str(runtime)], check=True)
    subprocess.run([str(python), '-m', 'pip', 'install', '-r', str(root / 'requirements-excel-ai.txt')], check=True)
print('Excel AI runtime ready. Codex CLI must also be installed and signed in.')
