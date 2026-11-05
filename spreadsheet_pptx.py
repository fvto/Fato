"""Local PowerPoint export using the bundled Artifact Tool runtime."""
import json
import base64
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile

import spreadsheet_compare as compare
from spreadsheet_report import presentation_model

APP_ROOT=Path(__file__).resolve().parent


def build_presentation(result,build_dir):
    runtime=Path(os.environ.get('CHINGLUH_ARTIFACT_RUNTIME',str(Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies')))
    node=runtime/'node/bin/node.exe'
    python=runtime/'python/python.exe'
    modules=runtime/'node/node_modules'
    skills=list((Path.home()/'.codex/plugins/cache/openai-primary-runtime/presentations').glob('*/skills/presentations'))
    if not node.is_file() or not python.is_file() or not modules.is_dir() or not skills:
        raise ValueError('Chưa tìm thấy runtime xuất PowerPoint trên máy.')
    skill=sorted(skills)[-1]
    link=APP_ROOT/'report_export/node_modules'
    if not link.exists():
        if os.name=='nt':
            script="New-Item -ItemType Junction -Path '"+str(link).replace("'","''")+"' -Target '"+str(modules).replace("'","''")+"' | Out-Null"
            command=['powershell','-NoProfile','-EncodedCommand',base64.b64encode(script.encode('utf-16-le')).decode('ascii')]
            completed=subprocess.run(command,capture_output=True,creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0))
            if completed.returncode and not link.exists(): raise ValueError('Không chuẩn bị được thư viện xuất PowerPoint.')
        else:
            link.symlink_to(modules,target_is_directory=True)
    build_dir=Path(build_dir)
    build_dir.mkdir(parents=True,exist_ok=True)
    payload=build_dir/'report.json'
    payload.write_text(json.dumps(presentation_model(result),ensure_ascii=False),encoding='utf-8')
    target=build_dir/'output/report.pptx'
    try:
        environment=os.environ.copy()
        environment.update(RUNTIME_NODE_MODULES=str(modules),RUNTIME_NODE=str(node),RUNTIME_PYTHON=str(python))
        completed=subprocess.run([str(node),str(APP_ROOT/'report_export/build_pptx.mjs'),str(payload),str(target),str(python),str(skill)],
                                 cwd=APP_ROOT,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=180,
                                 creationflags=getattr(subprocess,'CREATE_NO_WINDOW',0),env=environment)
    except subprocess.TimeoutExpired as error:
        raise ValueError('Xuất PowerPoint quá thời gian. Hãy chia nhỏ vùng dữ liệu rồi thử lại.') from error
    verified=False
    try:
        receipt=json.loads((build_dir/'validation.json').read_text(encoding='utf-8'))
        verified=(target.is_file() and receipt['finalSha256']==hashlib.sha256(target.read_bytes()).hexdigest()
                  and receipt['packageIntegrity']['status']=='pass'
                  and receipt['presentationLayout']['findingCount']==0
                  and receipt['firstPartyImport']['passed'] is True)
    except (OSError,ValueError,KeyError):
        pass
    # The Windows native renderer can fail during process teardown after the verified file is complete.
    # Accept that exit only with an intact, fully validated final artifact; never return a partial draft.
    if not verified or completed.returncode not in (0,3221226505):
        (build_dir/'error.log').write_text(f'Exit code: {completed.returncode}\n'+completed.stdout+'\n'+completed.stderr,encoding='utf-8')
        raise ValueError('Không xuất được PowerPoint. Xem nhật ký xuất cục bộ hoặc thử lại với ít vùng dữ liệu hơn.')
    if completed.returncode:
        (build_dir/'runtime-exit.json').write_text(json.dumps({'exit_code':completed.returncode,'artifact_verified':True}),encoding='utf-8')
    return target


def pptx_bytes(result):
    folder=compare.ROOT/'pptx_builds'
    folder.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='report_',dir=folder) as temporary:
        return build_presentation(result,Path(temporary)).read_bytes()
