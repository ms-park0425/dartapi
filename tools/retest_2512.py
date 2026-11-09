"""New annual reconstruction using prior source snapshots and updated guidance."""
from pathlib import Path
import shutil
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,Period,digest
from backtest_2512 import initialize,run_prepare

def main():
    for p in ('2503','2506'):
        b=Path((ROOT/'runs'/f'active_backtest_{p}.txt').read_text(encoding='utf-8').strip())
        if not (b/'accuracy_summary.json').exists():
            raise ValueError('Complete earlier quarter scoring and notes before the annual retest')
    old=Path((ROOT/'runs/active_backtest.txt').read_text(encoding='utf-8').strip())
    old_summary=load(old/'accuracy_summary.json')
    if not (old/'frozen_outputs.json').exists():raise ValueError('Annual baseline must be frozen')
    if (old/'retest_baseline.json').exists():raise ValueError('Current annual run is already a retest; do not silently replace it')
    base=initialize(Period(2025,4))
    dump(base/'retest_baseline.json',dict(path=str(old),scores=old_summary,final_sha256=digest(old/'run/final.xlsx'),
         policy='New extraction. Historical source snapshots reused, no reference values or previous findings reused as inputs.'))
    config=load(base/'run_config.json');root=base/'sandbox'
    manifest=load(old/'run/source_manifest.json');dump(base/'seed_source_manifest.json',manifest)
    for s in manifest:
        source=Path(s['path']);kind=s['kind'];name=s['company']
        if kind in ('pdf','disclosure_pdf','dart_pdf','factsheet'):
            label='경영공시_' if kind=='disclosure_pdf' else kind+'_'
            dest=root/'data'/name/'IR'/'2512'/(label+s['sha256']+source.suffix)
        elif kind in ('xbrl','taxonomy'):
            dest=root/'data'/name/'사업보고서'/'2025'/source.name
        elif kind=='document_xml':dest=root/'data'/name/'doc_2512.xml'
        else:continue
        dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(source,dest)
    config.update(cached_sources_only=True,use_fisis=False,ai_source_discovery=False,
        review_keywords=['보험부채 현황','회계모형별','보험계약마진','CSM','지급여력','일반회계와 감독회계','보험금 예실차','가정변경효과'])
    dump(base/'run_config.json',config)
    run_prepare(base)

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8');main()
