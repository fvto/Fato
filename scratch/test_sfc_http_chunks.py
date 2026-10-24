from pathlib import Path
import sys,logging,time
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from sfc_http import Session
logging.basicConfig(level=logging.INFO)
for plant,first,last in [('JV','2026-01-02','2026-01-02'),('VH','2026-01-02','2026-01-04')]:
 s=Session(plant);s.prepare();print('PREPARED',plant,flush=True)
 item=s.export_part(first,last,True,Path('Report/HFPAExports/http_pilot')/(plant+'_'+first+'_'+last))
 print('VERIFIED',plant,item['rows'],flush=True)
