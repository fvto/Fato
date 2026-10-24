from pathlib import Path
exec((Path(__file__).parent/'inspect_menu_script.py').read_text(encoding='utf8').split("for term in")[0])
for term in ['RaiseItemClick','OnItemClick','ItemClick:', 'DXSaveCurrentFactoryToCookies','SendPostBack("CLICK']:
 print('TERM',term)
 for m in list(re.finditer(re.escape(term),script))[:6]:print(script[max(0,m.start()-150):m.start()+650])
