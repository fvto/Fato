from pathlib import Path
source=(Path(__file__).parent/'test_http_single_jan.py').read_text(encoding='utf8')
source=source.replace("else '2026-01-01'","else '2026-01-07'").replace("'2026-01-01','2026-01-01'","'2026-01-01','2026-01-07'").replace('Jan1_direct','Jan_week_direct').replace('Jan 1 VH','Jan week VH')
exec(compile(source,'scratch/test_http_week.py','exec'))
