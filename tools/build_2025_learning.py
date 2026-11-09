"""Publish numerical-score summaries and value-free PDF navigation separately."""
from pathlib import Path
import sys
import csv

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from quarterly.core import load,dump,digest


def build():
    examples=[];runs=[];details=[];diagnoses=[]
    prior_analysis={}
    with (ROOT/'docs/2512_불일치_빈칸_분석.csv').open(encoding='utf-8-sig') as stream:
        prior_analysis={(r['company'],int(r['field'])):r for r in csv.DictReader(stream)}
    groups=load(ROOT/'docs/company_field_sources.json').get('field_groups',[])
    for period in ('2503','2506','2509','2512'):
        pointer=ROOT/'runs'/('active_backtest.txt' if period=='2512' else f'active_backtest_{period}.txt')
        if not pointer.exists():continue
        base=Path(pointer.read_text(encoding='utf-8').strip());run=base/'run'
        summary=load(base/'accuracy_summary.json',{});state=load(run/'state.json',{})
        runs.append(dict(period=period,path=str(base),status=state.get('status'),scores=summary,
            reviewed=load(run/'summary.json',{}),frozen=load(base/'frozen_outputs.json',{})))
        results=load(run/'review_results.json',{});sources={s['id']:s for s in load(run/'source_manifest.json',[])}
        viewed={(p['source_id'],p['page'],p['short']) for p in results.get('reviewed_pages',[])}
        rejected=load(run/'rejected_reviews.json',[])
        unresolved=set()
        if (run/'unresolved.csv').exists():
            with (run/'unresolved.csv').open(encoding='utf-8-sig') as stream:
                unresolved={(r['company'],int(r['field'])) for r in csv.DictReader(stream)}
        for f in results.get('findings',[]):
            if (f.get('source_id'),f.get('page'),f.get('short')) not in viewed:continue
            if f.get('decision') not in ('KEEP','FILL','CORRECT'):continue
            if any(r.get('finding')==f for r in rejected):continue
            source=sources.get(f['source_id'],{})
            examples.append(dict(period=period,company=source.get('company'),field=f['field'],
                source_sha256=source.get('sha256'),page=f['page'],source_kind=source.get('kind'),
                table=f['table'],row=f['row'],column=f['column'],raw_unit=f['raw_unit'],
                basis=f['basis'],value_status=f['value_status'],period_type=f['period_type'],
                evidence_level='image_reviewed_navigation_only',
                measure_status='unresolved_definition_or_conflict' if (f['short'],f['field']) in unresolved else 'pdf_verified_in_that_run',
                caution='과거 정의를 다음 분기에 자동 적용하지 않는다. 2509 교보 83~87은 별도/연결 기준 불일치 사례이므로 lessons 우선.' if f.get('short')=='교보' and 83<=f['field']<=87 else '현재 분기 표·범위·단위를 재확인'))
        for stage in ('stage1','stage2'):
            for r in load(base/f'{stage}_scores.json',[]):
                details.append(dict(period=period,stage=stage,company=r['company'],field=r['field'],status=r['status']))
                if stage!='stage2' or r['status'] not in ('DIFF','NO_VALUE'):continue
                field=r['field'];company=r['company'];old=prior_analysis.get((company,field),{}) if period=='2512' else {}
                group=next((g for g in groups if field in g['fields']),{})
                status='원인 미확정';cause='원문·계정 매핑·범위·기간·단위의 추가 대조 필요'
                if r['status']=='NO_VALUE':
                    cause='현재 수집·추출·선택 PDF 검토에서 해당 정의의 값을 확보하지 못함. 미공시로 단정하지 않음'
                elif period=='2509' and company=='교보' and field in range(83,88):
                    status='원인 확인';cause='별도 표를 사용했지만 정답은 연결 표와 대응. 84·85 기간구간도 1~2/2~5년과 1~3/3~5년으로 다름'
                elif period=='2509' and company=='삼성' and field==78:
                    status='부호 차이 확인·정답 규약 미확정';cause='IR 원문 음수와 정답본 양수가 다름. 동일 지표 원문 부호로 보정했으나 정답본 오류 여부는 단정하지 않음'
                elif period=='2512' and company=='DB손보' and field==97:
                    status='원인 확인·채점 후 코드 수정';cause='연간 사업보고서가 외부검증 전 잠정치임을 명시했으나 파서가 최종 칸에도 복사. 최종 경영공시와 차이. 이후 코드에서 잠정치의 최종 칸 복사 차단, 확정 실행에는 미소급'
                elif period=='2512' and company=='현대해상' and field==92 and r['status']=='NO_VALUE':
                    status='보류 이유 확인';cause='IR 이미지에 잠정 상태를 명시한 근거를 확인하지 못해 보류. 게시 IR이라는 이유만으로 잠정 칸을 채우지 않음'
                elif field in (23,92,95) and r['value'] is not None and abs(r['value']-r['answer'])<=.01:
                    status='정밀도 차이 후보';cause='표시 반올림 차이 가능성. 원문 정밀값·공시시점·경과조치 기준을 재확인하며 엄격 채점 불일치는 유지'
                if old:
                    cause += '; 기존 분석 분류: '+old['category']
                diagnoses.append(dict(period=period,company=company,field=field,result=r['status'],diagnosis_status=status,
                    cause=cause,next_table=old.get('next_table') or group.get('table','공시 재무제표/주석의 해당 계정'),
                    next_scope_check=old.get('scope') or group.get('scope','회사 본체·기간·별도/연결·총액/순액 확인'),
                    reference=r['answer'],observed=r['value'],action='기존 확정 결과 보존. 원문 진단 후 새 실행에서 수정') )
    dump(ROOT/'docs/review_navigation_2025.json',dict(version=1,
        policy='No answer/value/raw_values/quote included; historical navigation is not a current-period instruction.',examples=examples))
    dump(ROOT/'runs/learning/2025_validation_summary.json',dict(runs=runs))
    with (ROOT/'docs/2025_company_field_status.csv').open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=['period','stage','company','field','status']);writer.writeheader();writer.writerows(details)
    if diagnoses:
        with (ROOT/'docs/2025_오류원인메모.csv').open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=list(diagnoses[0]));writer.writeheader();writer.writerows(diagnoses)
    (ROOT/'docs/2025_오류원인메모.md').write_text('\n'.join([
        '# 불일치·빈칸 원인 메모','',
        '각 셀의 상태·원인 확인 수준·다음 확인 표·범위 점검·처리 방침은 `2025_오류원인메모.csv`에 기록한다. 원인 미확정과 확인된 원인을 구분한다. 정답과 다르다고 공시값이 틀렸다고 단정하지 않는다.','',
        '| 사례 | 확인 내용 | 다음 실행의 처리 |','|---|---|---|',
        '| 2509 교보 83~87 | 연결/별도 및 기간구간 불일치 확인 | 회사·칼럼별 기준을 명시한 새 실행에서 재확인 |',
        '| 2509 삼성생명 78 | IR 음수와 정답본 양수의 차이 확인 | 표시 규약·변동식 재확인, 정답 오류로 단정 금지 |',
        '| K-ICS 비율 일부 | 정밀도 차이 후보 | 원문 정밀값 유지, 기준·공시시점도 확인 |',
        '| 2506 한화 K-ICS | 본문 경과조치 후/각주 경과조치 전이 다름 | 기준을 맞춘 자본·요구자본·비율을 함께 확인 |',
        '| 2503/2506 회계모형별 CSM | 원수·수재 표와 순보유 표의 범위 차이 | 총액과 순액을 혼용하지 않음 |',
        '| 메리츠/DB CSM 변동 검산 | 일부 분기의 기초·변동 범위 미해결 | 역산으로 조정을 만들지 않고 원문 대조 |','',
        '숫자가 포함된 채점 메모는 진단용이며 자동 입력값의 출처로 쓰지 않는다. 검토 프롬프트에는 별도의 수치 없는 `quarterly_lessons.json`과 `review_navigation_2025.json`만 전달한다.','']),encoding='utf-8')
    lines=['# 2025년 분기 검증 및 다음 분기 준비','','정답본을 수정하지 않고 각 분기 입력을 지운 사본에서 추출했다. 아래는 개발 검증이며 독립 미학습 성능이 아니다.','',
        '| 분기 | 대상 | 1차 일치 | 2차 일치 | 2차 불일치 | 빈칸 | 전체 일치율 | PDF 검증 항목 |','|---|---:|---:|---:|---:|---:|---:|---:|']
    for run in runs:
        scores=run['scores'];a=scores.get('stage1',{}).get('ALL');b=scores.get('stage2',{}).get('ALL')
        prior=load(Path(run['path'])/'retest_baseline.json',{})
        if prior:
            x=prior['scores']['stage1']['ALL'];y=prior['scores']['stage2']['ALL']
            lines.append(f"| {run['period']} 기존 | {y['total']} | {x['ok']} | {y['ok']} | {y['wrong']} | {y['missing']} | {y['accuracy_all']:.1%} | 기존 실행 보존 |")
        if not a or not b:
            lines.append(f"| {run['period']} | 검증 진행 중 | | | | | | |")
        else:lines.append(f"| {run['period']} | {b['total']} | {a['ok']} | {b['ok']} | {b['wrong']} | {b['missing']} | {b['accuracy_all']:.1%} | {run['reviewed'].get('pdf_verified_fields',0)} |")
    lines += ['','CALC·수식은 채점에서 제외한다. 전체 일치율과 채워진 값의 일치율은 별개다. PDF 직접 검토는 인용한 페이지만 해당하며 전 셀 독립 검증을 의미하지 않는다.','',
        '## 재사용 파일','',
        '- `quarterly_lessons.json`: 의미·범위·기간·부호·수집 실수 방지 규칙. 미해결 가설도 명시.',
        '- `review_navigation_2025.json`: 실제 이미지 확인한 회사·칼럼별 표/행/열/단위 예시. 정답 수치는 포함하지 않는다.',
        '- `company_field_sources.json`: 기존 2512 출처·팩트시트 규칙 안내.',
        '- `2025_company_field_status.csv`: 분기·회사·칼럼별 일치/불일치/빈칸 상태.',
        '- `2025_오류원인메모.csv`: 불일치·빈칸별 원인 확인 수준·다음 확인 위치·처리 방침.',
        '- `2609_실행지침.md`: 다음 요청의 수집→PDF 직접 검토→최종 저장 절차.','',
        '## 반복 방법','',
        '기존 확정 실행은 덮어쓰지 않는다. 규칙 변경 후 재평가는 새 실행으로 구분하며 변경 전후 점수를 따로 보존한다. `tools/backtest_year.py --reuse`는 준비된 실행을 재사용한다. `tools/build_2025_learning.py`는 검증 요약과 수치 없는 위치 안내만 재생성한다.']
    (ROOT/'docs/2025_검증현황.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print('Quarter runs:',len(runs),'reviewed navigation examples:',len(examples),'cell status rows:',len(details))


if __name__=='__main__':
    sys.stdout.reconfigure(encoding='utf-8');build()
