"""Readable quarter backtest report. Reads frozen results; does not tune them."""
from pathlib import Path
import argparse
import csv
import shutil
import sys
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,digest,COMPANIES
from quarterly.review import RATIOS
from quarterly.workbook import DataBook


def report(base):
    protocol=load(base/'blind_protocol.json');target=protocol['target_period'];previous=protocol['previous_period'];run=base/'run'
    summary=load(base/'accuracy_summary.json');catalog=load(run/'field_catalog.json');run_summary=load(run/'summary.json')
    before={(r['company'],r['field']):r for r in load(base/'stage1_scores.json')};after=load(base/'stage2_scores.json')
    changed={(r['company'],r['field']):r for r in load(run/'changes.json')}
    rows=[]
    for r in after:
        a=before[(r['company'],r['field'])];change=changed.get((r['company'],r['field']),{})
        rows.append(dict(company=r['company'],field=r['field'],label=catalog.get(str(r['field']),{}).get('label',''),answer=r['answer'],stage1=a['value'],stage1_status=a['status'],stage2=r['value'],stage2_status=r['status'],difference=r['difference'],tolerance=r['tolerance'],ratio_rounding_only=r['status']=='DIFF' and r['field'] in RATIOS and round(r['value'],2)==round(r['answer'],2),source='; '.join(f"{e['source_id']} p{e['page']}" for e in change.get('evidence',[]))))
    with (base/'accuracy_detail.csv').open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=list(rows[0]));writer.writeheader();writer.writerows(rows)
    assert digest(Path(protocol['answer_file']))==protocol['answer_sha256']
    DataBook(run/'final.xlsx',protect_through=previous).assert_history_unchanged(load(run/'protected_history.json'))
    assert digest(run/'final.xlsx')==load(base/'frozen_outputs.json')['final_sha256']
    output=base/f'{target}_추출결과.xlsx';shutil.copy2(run/'final.xlsx',output)
    regressions=[r for r in rows if r['stage1_status']=='OK' and r['stage2_status']=='DIFF']
    failures=[c for c in load(run/'validation_after.json') if c['status']=='fail']
    dump(base/'verification.json',dict(answer_unchanged=True,history_preserved_through=previous,fisis_used=False,calc_excluded=len(load(base/'excluded_calc_cells.json')),regressions=len(regressions),remaining_failed_checks=failures,output=str(output)))
    lines=[f'# {target} 추출·정확도 검증','',f'{previous}까지 남긴 DATA 사본에서 {target} 직접 입력값을 제거하고 해당 분기 DART/IR 자료로 추출했습니다. 2512에서 정리한 위치·정의 안내를 사용했습니다. 원본·정답본은 보존했습니다.','',
       '결과를 확정한 다음 정답본과 비교했으며, 채점 후 정답에 맞춰 결과값을 수정하지 않았습니다. FISIS와 확인용 CALC·수식 셀은 제외했습니다.','',
       '| 지표 | 1차 코드 | 2차 PDF |','|---|---:|---:|']
    one,two=summary['stage1']['ALL'],summary['stage2']['ALL']
    for label,key in [('비교 대상','total'),('일치','ok'),('불일치','wrong'),('빈칸','missing'),('미계산','uncalculated')]:lines.append(f'| {label} | {one[key]} | {two[key]} |')
    for label,key in [('채워진 값 중 일치율','accuracy_filled'),('전체 대상 중 일치율','accuracy_all')]:lines.append(f'| {label} | {one[key]:.2%} | {two[key]:.2%} |')
    lines += ['',f"허용오차: 금액 max(1백만원, 정답 절댓값의 0.1%), 비율 max(0.0001, 정답 절댓값의 0.1%). 수식/계산 셀 {len(load(base/'excluded_calc_cells.json'))}개 제외.",'',
       f"선택한 PDF {run_summary['reviewed_pages']}페이지에서 {run_summary['pdf_verified_fields']}개 항목을 확인하고 {run_summary['stage2_changes']}셀을 보충·수정했습니다. 모든 PDF 페이지·셀을 독립 검증한 결과는 아닙니다.",
       f"1차 일치 → 2차 불일치: {len(regressions)}개. 비율 반올림 차이로 분류되는 2차 불일치: {sum(r['ratio_rounding_only'] for r in rows)}개. 엄격 채점에서는 불일치로 유지했습니다.",'',
       '## 회사별 결과','','| 회사 | 대상 | 1차 일치 | 2차 일치 | 2차 불일치 | 빈칸 | 채워진 값 일치율 |','|---|---:|---:|---:|---:|---:|---:|']
    for _,s in COMPANIES:
        a,b=summary['stage1'][s],summary['stage2'][s];lines.append(f"| {s} | {b['total']} | {a['ok']} | {b['ok']} | {b['wrong']} | {b['missing']} | {b['accuracy_filled']:.1%} |")
    lines += ['', '## 남은 검산 문제','']+[f"- {c['company']} {c['check']}: 차이 {c.get('difference')}백만원. 범위/기초잔액 등의 원인을 확정하지 못해 기존 값을 보존했습니다." for c in failures]
    diagnosis=load(base/'post_score_diagnosis.json',{})
    if diagnosis:
        lines += ['', '## 채점 후 원문 대조에서 확인한 차이','',
          '- 교보 CSM 기간별 5개: 정답본은 연결 기준과 1~3년/3~5년 구간이고, 이번 입력은 별도 기준과 1~2년/2~5년 구간입니다. 회사·기간별 정의를 추가로 명시해야 합니다.',
          '- 삼성생명 CSM 조정: 원문 IR은 음수이고 정답본은 양수입니다. 이번에는 원문 부호로 보정했으며, 정답본의 표시 규약 또는 오류 여부는 별도 확인이 필요합니다.',
          '- K-ICS 비율은 정답본 표시 정밀도와 원문 정밀도의 차이가 대부분입니다. 경계값인 DB손보를 포함해 엄격 채점의 불일치는 그대로 유지합니다.',
          '- 채점 후 분석은 결과를 설명하기 위한 것이며, 확정한 엑셀과 채점 숫자를 소급 변경하지 않았습니다.']
    lines += ['',f'[추출 엑셀]({output.name}) · [셀별 정답 비교](accuracy_detail.csv)','', '이번 검증은 출처 안내를 다른 분기에 적용한 개발 검증입니다. 기존 추출 코드가 과거 2509 개발에 사용된 이력이 있어 완전히 독립된 미학습 평가로 해석하지 않습니다.']
    (base/'검증결과.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print('report:',base/'검증결과.md');print('overall:',one,two);print('regressions:',len(regressions))


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('base',type=Path);args=parser.parse_args();report(args.base)
