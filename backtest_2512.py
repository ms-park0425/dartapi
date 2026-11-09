"""Blind reconstruction of 2512 from a workbook truncated after 2509.

The target answers are read only by `score`, after the reviewed workbook is frozen.
The answer workbook itself is never modified.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import sys

import openpyxl
from openpyxl.formula.translate import Translator

from quarterly.core import COMPANIES, Period, digest, dump, load, korea_now, numeric
from quarterly.workbook import DataBook, resolve

ROOT = Path(__file__).resolve().parent
FORMULAS = {9,10,11,14,24,31,38,45,46,47,48,49,50,51,63,64,68,71,72,82,88,89,90,91,103,104,159}
KEYWORDS = ["보험부채 현황", "일반회계와 감독회계", "가정변경효과", "보험금 예실차", "예상손해율", "보험계약마진",
            "CSM", "지급여력", "운용자산", "해약환급금준비금", "재무상태표", "손익계산서", "보험서비스", "보험부채 변동"]


def initialize(period=None):
    period = period or Period(2025, 4)
    code = period.code
    previous = Period(period.year, period.quarter-1) if period.quarter>1 else Period(period.year-1,4)
    previous_code = previous.code
    config = load(ROOT / "quarterly.json")
    answer = ROOT / config["reference_workbook"]
    answer_sha = digest(answer)
    base = ROOT / "runs" / "backtests" / (code + "_" + korea_now().strftime("%Y%m%d_%H%M%S"))
    base.mkdir(parents=True)
    sandbox = base / "sandbox"
    sandbox.mkdir()
    book = DataBook(answer, protect_through=previous_code)
    for sheet in list(book.wb.worksheets):
        if sheet.title != "DATA":
            book.wb.remove(sheet)
    for _, short in COMPANIES:
        source, target = book.rows[(previous_code, short)], book.rows[(code, short)]
        for field, ec in book.columns.items():
            if field <= 3:
                continue
            old = book.ws.cell(source, ec)
            cell = book.ws.cell(target, ec)
            cell.value = Translator(old.value, origin=old.coordinate).translate_formula(cell.coordinate) if field in FORMULAS and old.data_type == "f" else None
            cell.comment = None
    for (p,short), row in list(book.rows.items()):
        if int(p) > int(code):
            for ec in book.columns.values():
                book.ws.cell(row,ec).value=None
    masked = sandbox / f"input_through_{previous_code}.xlsx"
    book.save(masked)
    target_numeric = sum(numeric(v) for _,s in COMPANIES for v in book.values(code,s).values())
    if target_numeric:
        raise ValueError("Blind input still has target numeric answers")
    for name,_ in COMPANIES:
        for relative in (Path("IR")/code, Path(period.report_type)/str(period.year)):
            src, dst = ROOT / "data" / name / relative, sandbox / "data" / name / relative
            if src.exists():
                shutil.copytree(src,dst)
        src = ROOT / "data" / name / (f"doc_annual_{period.year}.xml" if period.quarter==4 else f"doc_{code}.xml")
        if src.exists():
            dst = sandbox / "data" / name
            dst.mkdir(parents=True,exist_ok=True)
            shutil.copy2(src,dst / src.name)
            if src.name != f"doc_{code}.xml":shutil.copy2(src,dst / f"doc_{code}.xml")
    # Group factbooks are supplied to the relevant subsidiary, with entity headers checked in PDF review.
    for group, children in (("KB금융지주",["KB생명","KB손해보험"]),("신한지주",["신한라이프생명"])):
        src=ROOT/"data"/group/"IR"/code
        for child in children:
            for file in src.glob("*"):
                dst=sandbox/"data"/child/"IR"/code/file.name
                dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(file,dst)
    run_config = {**config,"workbook":str(masked),"protect_through":previous_code,"use_fisis":False,"ai_source_discovery":False,
                  "cached_sources_only":code=='2512',"review_keywords":KEYWORDS}
    run_config.pop("reference_workbook",None)
    dump(base/"run_config.json",run_config)
    dump(base/"blind_protocol.json",dict(answer_file=str(answer),answer_sha256=answer_sha,input_file=str(masked),input_sha256=digest(masked),
         previous_period=previous_code,target_period=code,numeric_target_answers=target_numeric,other_worksheets_removed=True,
         fisis_used=False,source_policy="target-period DART/IR sources only; future answer rows removed; answers opened only for final scoring"))
    pointer='active_backtest.txt' if code=='2512' else f'active_backtest_{code}.txt'
    (ROOT/"runs"/pointer).write_text(str(base.resolve()),encoding="utf-8")
    print("검증 사본 준비:",base,flush=True)
    return base


def run_prepare(base):
    from quarterly.pipeline import prepare
    config=load(base/"run_config.json")
    period=Period.parse(load(base/'blind_protocol.json')['target_period'])
    run=prepare(base/"sandbox",period,config,offline=config.get('cached_sources_only',False),run=base/"run")
    dump(base/"progress.json",dict(stage="awaiting_direct_pdf_review",run=str(run)))
    print("검증 실행 폴더:",run,flush=True)


def freeze_outputs(base, final):
    frozen=dict(stage1_sha256=digest(base/'run/stage1.xlsx'),final_sha256=digest(final))
    path=base/'frozen_outputs.json'
    existing=load(path,{})
    if existing and any(existing.get(k)!=v for k,v in frozen.items()):
        raise ValueError('확정된 결과가 변경됐습니다. 채점 후 튜닝은 새 검증 실행으로 구분하세요.')
    if not existing:
        dump(path,{**frozen,'frozen_at':korea_now().isoformat()})
    return frozen


def score(base, final=None):
    protocol=load(base/"blind_protocol.json")
    target=protocol['target_period'];previous=protocol['previous_period']
    run=base/"run"
    final=Path(final) if final else run/"final.xlsx"
    if not final.exists():raise ValueError("2차 결과를 확정한 뒤에 채점합니다.")
    freeze_outputs(base,final)
    answer=Path(protocol["answer_file"])
    if digest(answer)!=protocol["answer_sha256"]:raise ValueError("정답 파일 변경 감지")
    from quarterly.history import read_reference,compare
    from quarterly.availability import formula_available
    from quarterly.review import CHECKS
    scoring_answer=base/'answer_for_scoring.xlsx'
    if scoring_answer.exists():
        truths,uncached=read_reference(scoring_answer)
    else:
        truths,uncached=read_reference(answer)
    truths={k:v for k,v in truths.items() if k[0]==target}
    truth_book=DataBook(scoring_answer if scoring_answer.exists() else answer)
    truths={k:v for k,v in truths.items() if formula_available(truth_book.wb,'DATA',truth_book.ws.cell(truth_book.rows[k[:2]],truth_book.columns[k[2]]).coordinate)}
    # User excludes confirmation/calculation cells: score direct inputs only.
    input_book=DataBook(run/'stage1.xlsx',protect_through=previous)
    excluded=[]
    for key in list(truths):
        p,s,c=key
        reference_cell=truth_book.ws.cell(truth_book.rows[(p,s)],truth_book.columns[c])
        input_cell=input_book.ws.cell(input_book.rows[(p,s)],input_book.columns[c])
        if reference_cell.data_type=='f' or input_cell.data_type=='f':
            excluded.append(dict(period=p,company=s,field=c,reason='calculation/formula cell excluded by user'))
            del truths[key]
    dump(base/'excluded_calc_cells.json',excluded)
    calculated_stage1=base/'stage1_calculated.xlsx'
    stages={"stage1":calculated_stage1 if calculated_stage1.exists() else run/"stage1.xlsx","stage2":final}
    summaries={}
    for label,path in stages.items():
        current,missing_cache=read_reference(path)
        current_book=DataBook(path,protect_through=previous)
        unavailable=set()
        for key in list(current):
            if key[0]==target and not formula_available(current_book.wb,'DATA',current_book.ws.cell(current_book.rows[key[:2]],current_book.columns[key[2]]).coordinate):
                unavailable.add(key);current.pop(key)
        cache_keys={(r['period'],r['company'],r['field']) for r in missing_cache}
        comparisons=[]
        for key,expected in truths.items():
            value=current.get(key);status,delta,tolerance=compare(value,expected,key[2])
            if status=='NO_VALUE' and key in cache_keys and key not in unavailable:status='UNCALCULATED'
            comparisons.append(dict(period=key[0],company=key[1],field=key[2],answer=expected,value=value,status=status,difference=delta,tolerance=tolerance))
        summary={}
        for short in [s for _,s in COMPANIES]+['ALL']:
            rows=[r for r in comparisons if short=='ALL' or r['company']==short]
            n=len(rows);ok=sum(r['status']=='OK' for r in rows);fail=sum(r['status']=='DIFF' for r in rows);miss=sum(r['status']=='NO_VALUE' for r in rows);calc=sum(r['status']=='UNCALCULATED' for r in rows)
            summary[short]=dict(total=n,ok=ok,wrong=fail,missing=miss,uncalculated=calc,accuracy_filled=ok/(ok+fail) if ok+fail else None,accuracy_all=ok/n if n else None)
        dump(base/f'{label}_scores.json',comparisons);summaries[label]=summary
    dump(base/'accuracy_summary.json',summaries)
    print(json.dumps(summaries,ensure_ascii=False,indent=2))
    return summaries


if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    parser=argparse.ArgumentParser();parser.add_argument('action',choices=['prepare','score']);parser.add_argument('--base',type=Path);parser.add_argument('--period',default='2512')
    args=parser.parse_args()
    if args.action=='prepare':run_prepare(args.base or initialize(Period.parse(args.period)))
    else:
        pointer='active_backtest.txt' if args.period=='2512' else f'active_backtest_{args.period}.txt'
        score(args.base or Path((ROOT/'runs'/pointer).read_text(encoding='utf-8')))
