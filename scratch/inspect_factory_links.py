from pathlib import Path
exec((Path(__file__).parent/'inspect_sfc_factory.py').read_text(encoding='utf8').split('import re')[0])
import re
for m in re.finditer(r'<(?:li|a|span)[^>]*(?:menuRightTop_DXI0(?:i\d+)?_|dxm-content)[^>]*>',body):
 part=body[m.start():m.start()+500]
 if 'VH' in part or 'JV' in part:print(part[:450].replace('\n',' '))
for name in ['VH','VH2','JV','JV2']:
 m=re.search(r'>'+name+r'<',body)
 if m:print(name,body[m.start()-300:m.start()+100])
print('scripts:',re.findall(r'<script[^>]+src="([^"]+)"',body))
