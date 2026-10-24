from pathlib import Path
import sys,logging
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from sfc_http import Session
logging.basicConfig(level=logging.INFO)
for plant in ['VH2','JV','JV2']:
 s=Session(plant);s.prepare()
 print('PREPARED',plant,flush=True)
 item=s.export_part('2026-10-24','2026-10-24',True,Path('Report/HFPAExports/http_pilot')/('today_'+plant+'_NIKE'))
 print('VERIFIED',plant,item['rows'],flush=True)
