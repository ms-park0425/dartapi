"""Recalculate execution copies, freeze results, then open target references for scoring."""
from pathlib import Path
import argparse
import shutil
import subprocess
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,digest
from backtest_2512 import freeze_outputs,score
from tools.report_quarter_backtest import report

def recalc(path):
    result=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-File',
        str(ROOT/'recalculate_excel.ps1'),'-Path',str(path.resolve())],capture_output=True,text=True,timeout=240)
    if result.returncode or 'recalculated' not in result.stdout:
        raise RuntimeError('실행 사본 Excel 재계산 실패; 원본·정답본은 수정하지 않음')

def evaluate(base):
    run=base/'run';protocol=load(base/'blind_protocol.json')
    if (base/'frozen_outputs.json').exists():
        score(base);report(base);return
    stage=base/'stage1_calculated.xlsx';shutil.copy2(run/'stage1.xlsx',stage)
    recalc(stage);recalc(run/'final.xlsx')
    frozen=freeze_outputs(base,run/'final.xlsx')
    dump(base/'pre_score_freeze.json',frozen)
    # Only after all outputs are frozen is a copy of the reference opened by Excel.
    answer=Path(protocol['answer_file']);assert digest(answer)==protocol['answer_sha256']
    target=base/'answer_for_scoring.xlsx';shutil.copy2(answer,target);recalc(target)
    assert digest(run/'final.xlsx')==frozen['final_sha256']
    assert digest(answer)==protocol['answer_sha256']
    summary=load(run/'summary.json');summary['recalculation']='recalculated';dump(run/'summary.json',summary)
    state=load(run/'state.json');state['summary']=summary;dump(run/'state.json',state)
    score(base);report(base)

if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    parser=argparse.ArgumentParser();parser.add_argument('periods',nargs='+');args=parser.parse_args()
    for code in args.periods:
        pointer=ROOT/'runs'/('active_backtest.txt' if code=='2512' else f'active_backtest_{code}.txt')
        evaluate(Path(pointer.read_text(encoding='utf-8').strip()))
