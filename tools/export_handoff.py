"""Build an explicit, credential-free handoff ZIP from the current working files."""
from pathlib import Path
import argparse
import hashlib
import json
import sys
import zipfile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from quarter_pipeline import read_config

FILES=['README.md','AGENTS.md','.env.example','.gitignore','.gitattributes',
       'quarter_pipeline.py','quarter_picker.py','dart_api.py','parse_annual.py',
       'ir_parse.py','ir_pdf.py','backtest_2512.py','requirements.txt','requirements-dev.txt',
       'pytest.ini','recalculate_excel.ps1','convert_legacy_factsheet.ps1','처음설정.cmd','분기입력.cmd']
DOCS=['2506_최종점검.md','개발_운영참고.md','2609_실행지침.md','quarterly_lessons.json','review_navigation_2025.json',
      'company_field_sources.json','회사별_칼럼별_출처안내.md','회사별_칼럼별_출처목록.csv',
      '2025_검증현황.md','2025_오류원인메모.md','2025_오류원인메모.csv',
      '2025_company_field_status.csv','2512_불일치_빈칸_분석.csv']


def export(output,workbook=None):
    output=Path(output).resolve()
    config=read_config(ROOT/'quarterly.json')
    workbook=Path(workbook or config['workbook'])
    if not workbook.is_absolute():workbook=ROOT/workbook
    reference=ROOT/config['reference_workbook']
    if not workbook.is_file() or not reference.is_file():raise FileNotFoundError('작업 Excel 또는 정답본이 없습니다.')
    if workbook.resolve()==reference.resolve():raise ValueError('정답본은 작업 양식으로 지정할 수 없습니다.')
    paths=[ROOT/name for name in FILES]+[ROOT/'docs'/name for name in DOCS]
    for folder in ['quarterly','tools','tests']:
        paths.extend(sorted((ROOT/folder).glob('*.py')))
    if any(not p.is_file() for p in paths):raise FileNotFoundError('Required handoff source missing')
    if output.exists():raise FileExistsError('기존 전달용 ZIP을 덮어쓰지 않습니다. 다른 출력 경로를 지정하세요.')
    sources=[*paths,workbook,reference]
    if output in [p.resolve() for p in sources]:raise ValueError('Output overlaps an input file')
    defaults=json.loads((ROOT/'quarterly.json').read_text(encoding='utf-8'))
    packed_config={**defaults,'workbook':'작업양식.xlsx','reference_workbook':reference.name,'use_fisis':False}
    hashes={}
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'x',compression=zipfile.ZIP_DEFLATED) as archive:
        for p in paths:
            name=p.relative_to(ROOT).as_posix();data=p.read_bytes()
            archive.writestr('Dartapi/'+name,data);hashes[name]=hashlib.sha256(data).hexdigest()
        for p,name in [(workbook,'작업양식.xlsx'),(reference,reference.name)]:
            data=p.read_bytes();archive.writestr('Dartapi/'+name,data);hashes[name]=hashlib.sha256(data).hexdigest()
        archive.writestr('Dartapi/quarterly.json',json.dumps(packed_config,ensure_ascii=False,indent=2))
        archive.writestr('Dartapi/HANDOFF_MANIFEST.json',json.dumps(dict(
            files=hashes,excluded=['.env','quarterly.local.json','.venv','.git','runs','data','login credentials'],
            setup='README.md → Codex에서 폴더 열기 → 채팅으로 실행 환경 준비 → 담당자 본인의 DART 키·작업 파일 설정'),ensure_ascii=False,indent=2))
    return output


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,default=ROOT/'runs/handoff/Dartapi_인수인계.zip')
    parser.add_argument('--workbook',type=Path);args=parser.parse_args()
    print(export(args.output,args.workbook))
