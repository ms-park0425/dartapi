"""Build navigation metadata from an already scored development run.

Answer values stay in the audit CSV. Runtime rules contain locations, never answers.
"""
from pathlib import Path
import argparse
import csv
import json
import sys
from collections import Counter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from quarterly.core import COMPANIES, load, dump
from quarterly.review import RATIOS
from ir_parse import SPECS

NOTES = {
 '미래에셋생명':['팩트시트 SAP-BS의 운용자산·OCI 세부, EFF의 위험보험료·보험금, PROFIT의 예실차를 우선 탐색.', 'IR CSM 변동표와 K-ICS의 (e) 잠정치를 구분. 초기 IR 잠정 자본과 후속 공시 자본 차이를 오류로 단정하지 않음.'],
 '삼성생명':['IR CSM Movement는 별도 기준 확인 후 이용. 다른 손익·BS 슬라이드의 연결 기준을 혼용하지 않음.', '2512 IR의 예상 K-ICS 옆 자본 표는 9월 기준이었음. 비율과 자본의 기준일을 각각 확인.'],
 '교보생명':['경영공시 일반회계와 감독회계의 차이에서 일반회계(A)를 선택.', '회계모형별 CSM은 GMM+VFA 합산. 계리적 가정 표의 신계약 CSM을 탐색.', '보험계약마진 상각 일정은 이자부리 차감 순상각일 수 있어 기대수익 인식 버킷으로 자동 매핑 금지.'],
 '한화생명':['팩트시트/IR의 별도 보험손익·CSM 신계약·상각·예실차 표를 탐색.', 'IR 투자손익에 관계기업 지분법 손익이 포함되는지 확인. 일반회계 투자손익과 바로 교체하지 않음.'],
 '신한라이프생명':['경영공시 일반회계(A), 회계모형별 GMM+VFA CSM, 계리적 가정의 신계약/가정변경을 탐색.', '경영공시 지급여력 총괄의 해당 분기와 경과조치 여부를 확인.', 'K-ICS 시가평가 운용자산과 자산운용률의 운용자산은 혼용 금지. 순상각 일정은 기대수익 버킷으로 사용 금지.'],
 '동양생명':['IR CSM 변동표의 연초·연말 잔액과 연간 누계 신계약·이자부리·조정·상각을 탐색.', '그래프 숫자가 PDF 텍스트에 누락될 수 있어 이미지로 재확인. 잠정과 확정 K-ICS 발표 시점을 분리.'],
 'KB생명':['KB금융 IR의 KB라이프 개별 재무제표 각주와 CSM 변동표를 탐색. 그룹 전체/연결 순이익을 별도 일반회계와 혼용 금지.', '그룹 IR의 이자+기타 결합값으로 Col78/79를 임의 분리하지 않음.', '경영공시 계리적 가정의 신계약 CSM, 지급여력 총괄의 확정 자본을 탐색. 글꼴 인코딩으로 55가 53으로 추출된 사례가 있어 이미지 확인.'],
 '삼성화재':['IR 장기보험 CSM Movement의 연간 합계를 탐색. 분기별 숫자를 연간으로 오인하지 않음.', '장기보험 보험손익·예실차와 회사 전체 보험손익의 범위를 확인. 그래프 숫자는 이미지 확인.'],
 '현대해상':['IR CSM 변동표의 보유/출재 차감 범위를 확인. 순보유 CSM과 XBRL 발행계약 CSM을 교체하지 않음.', '신계약 원수+수재, K-ICS 잠정 비율을 각각 확인. 조원 단위의 거친 자본 그래프로 정밀값을 교체하지 않음.'],
 '메리츠화재':['IR 별도 손익표와 CSM 연간(YTD) 변동표를 탐색. 같은 슬라이드 우측 4Q 단독값을 사용하지 않음.', 'CSM 원수 신계약과 변동표의 신계약 범위, 잠정/후속 K-ICS 차이를 보류·기록.'],
 'KB손해보험':['KB금융 IR의 보험사별 CSM 변동표와 운용자산을 탐색.', '신계약 CSM은 경영공시 계리적 가정(원수·수재)과 그룹 변동표가 다를 수 있음. 선택 기준 확정 전 둘을 동일 값으로 취급하지 않음.', '이자+기타 결합값은 Col78/79로 임의 분리 금지. 확정 K-ICS는 경영공시 지급여력 총괄에서 확인.'],
 'DB손해보험':['경영공시 재무정보 요약·보험계약부채/가정·지급여력 총괄을 탐색.', '운용자산은 재무정보의 일반적인 자산운용 표를 사용. K-ICS 건전성감독기준 시가평가 운용자산과 혼용 금지.', 'GMM만 있는 CSM 표는 BEL/RA/CSM/PAA 4열. BEL/RA는 잔여보장 범위이므로 순부채 총액으로 쓰지 않음.'],
}

GROUPS = [
 dict(id='invested_assets',fields=[21,22,23],table='자산운용 현황 / 자산운용률 / SAP-BS',rows='총자산, 운용자산, 자산운용률',unit='표 단위 확인',scope='K-ICS 시가평가 자산과 혼용 금지'),
 dict(id='oci',fields=list(range(25,31)),table='기타포괄손익 변동 / SAP-BS',rows='보험·재보험 순금융손익, 현금흐름위험회피 등 해당 계정',unit='원/백만원/억원 별도 확인',scope='세전/세후, 당기변동/잔액, 보험/재보험을 구분'),
 dict(id='pnl',fields=[32,35,40,41,42,43,44],table='일반회계와 감독회계의 차이 / 별도 손익',rows='일반회계(A) 해당 손익',unit='통상 억원, XBRL 정밀값 우선',scope='별도·당기 누계. 연결 손익 및 관계기업 포함 범위 구분'),
 dict(id='net_liabilities',fields=[73,74],table='보험계약부채 변동 주석 / 발행·출재 보험계약 주석',rows='BEL, RA 기말 및 출재 대응 계정',unit='주석 단위 확인',scope='LRC+LIC 및 원수·출재 순액 정의를 확인. 잔여보장 포트폴리오 표만으로 총액 산출 금지'),
 dict(id='csm',fields=list(range(75,81)),table='CSM Movement / 회계모형별·포트폴리오별 보험부채 현황 / 계리적 가정',rows='기말, 신계약, 상각, 조정, 이자, 전년말',unit='억원×100 / 십억원×1000 / 조원×1000000',scope='stock은 분기말, flow는 연초 누계, 전기CSM은 전년말. 원수·수재·출재 및 결합 항목 구분'),
 dict(id='csm_maturity',fields=list(range(83,88)),table='보험계약마진 기대수익 인식 시기 / 보험수익 인식 시기',rows='1년, 1~2년, 2~5년, 5~10년, 10년초과',unit='주석 단위 확인',scope='이자부리 차감 순상각 일정은 대체 출처로 사용 금지. 원문 버킷 경계와 출재 제외 확인'),
 dict(id='kics_provisional',fields=[92,93,94],table='IR K-ICS / 지급여력 잠정치',rows='해당 분기 예상 비율, 지급여력금액, 지급여력기준금액',unit='%→소수비율, 자본은 원문 단위',scope='(e)·잠정 여부와 각 셀 기준일 확인. 확정값을 잠정값으로 복제하지 않음'),
 dict(id='kics_final',fields=[95,96,97],table='지급여력비율 총괄 / 확정 지급여력',rows='당분기 비율, 지급여력금액, 지급여력기준금액',unit='%→소수비율, 통상 억원×100',scope='경과조치 전후·연결 범위·공시 확정 상태 확인. 원문 정밀 비율 유지'),
 dict(id='loss_ratio',fields=[101,102],table='보험금 예실차비율 / 최적가정',rows='예상손해율(A), 실제손해율(B)',unit='%→소수비율',scope='연간 누계. 정답본 표시 자릿수 차이는 숫자 오류와 별도로 기록'),
 dict(id='experience',fields=list(range(111,118)),table='예실차 / 손익 세부 현황',rows='보험금차, 예상·실제 보험금, 사업비차, 예상·실제 사업비',unit='원문 단위 확인',scope='회사전체/장기보험 범위와 누계 확인'),
 dict(id='assumptions',fields=list(range(146,151))+list(range(153,159)),table='보험부채 변동내역 > 계리적 가정',rows='해지율, 위험률, 사업비율, 기타, 물량차이, 손실요소',unit='통상 억원×100, 괄호 음수',scope='BEL+RA는 앞 두 열 합, CSM은 마지막 열. 손실요소 행과 물량차이 행 분리'),
]


def build(base):
    run=base/'run';catalog=load(run/'field_catalog.json');manual=load(base/'manual_findings.json')
    scores=load(base/'stage2_scores.json');score={(r['company'],r['field']):r for r in scores}
    sources={s['id']:s for s in load(run/'source_manifest.json')}
    candidates=load(run/'candidates.json')
    book=dict(version=1,training_period='2512',policy=['동일 지표는 정밀 데이터 우선. PDF 차이만으로 기존값 교체 금지.', '페이지 번호는 학습 예시이며 다음 분기에는 표·행·열·각주를 재확인.', '정답 일치는 의미 일치의 증명이 아님. FISIS 제외.'],field_groups=GROUPS,companies={})
    inventory=[];ledger=[]
    for company,short in COMPANIES:
        examples=[]
        for f in manual:
            if f['short']!=short:continue
            source=sources.get(f['source_id'],{})
            examples.append(dict(field=f['field'],table=f['table'],row=f['row'],column=f['column'],raw_unit=f['raw_unit'],scale=f['scale'],operation=f['operation'],basis=f['basis'],period_type=f['period_type'],value_status=f['value_status'],reference_period='2512',reference_page=f['page'],source_id=f['source_id'],source_url=source.get('url'),source_file=Path(source.get('original_path',source.get('path',''))).name))
        specs=SPECS.get(company,{})
        locations=[dict(field=i.col,sheet=i.sheet,row_label=i.label,section=i.section,mode=i.mode,sign=i.sign,scale=i.unit if i.unit is not None else specs.get('unit')) for i in specs.get('items',[]) if str(i.col) in catalog]
        book['companies'][company]=dict(short=short,notes=NOTES[company],examples=examples,factsheet_locations=locations)
        for field,meta in catalog.items():
            f=int(field);s=score.get((short,f));ex=[e for e in examples if e['field']==f];loc=[x for x in locations if x['field']==f]
            own=[c for c in candidates if c['short']==short and c['field']==f]
            group=next((g for g in GROUPS if f in g['fields']),None)
            extracted=[e for c in own if c.get('method')=='factsheet_rules' for e in c.get('evidence',[])]
            row=dict(company=short,field=f,label=meta['label'],status_2512=s['status'] if s else 'NOT_SCORED',
                 pdf_tables='; '.join(dict.fromkeys(e['table'] for e in ex)),
                 pdf_rows='; '.join(dict.fromkeys(e['row'] for e in ex)),
                 pdf_units='; '.join(dict.fromkeys(e['raw_unit'] for e in ex)),
                 pdf_reference_pages='; '.join(f"{e['source_file']} p{e['reference_page']}" for e in ex),
                 factsheet_navigation='; '.join(f"{e['sheet']} / {e['row_label']} / {e['mode']}" for e in loc),
                 pdf_location_status='IMAGE_REVIEWED_2512' if ex else 'NOT_LOCATED',pdf_examples=json.dumps(ex,ensure_ascii=False),
                 factsheet_rule_status='CODED_RULE' if loc else 'NO_CODED_RULE',factsheet_locations=json.dumps(loc,ensure_ascii=False),
                 extracted_cells_2512=json.dumps(extracted,ensure_ascii=False),legacy_methods=';'.join(sorted({c['method'] for c in own})),
                 next_table=group['table'] if group else '별도 XBRL 계정 및 해당 주석 확인',scope=group['scope'] if group else '별도·연결/기간/단위/계정 정의 확인')
            inventory.append(row)
            if s and s['status']!='OK':
                category=('ratio_display_precision' if s['status']=='DIFF' and f in RATIOS and round(s['value'],2)==round(s['answer'],2) else group['id'] if group else 'account_mapping_unresolved')
                ledger.append(dict(company=short,field=f,label=meta['label'],status=s['status'],category=category,answer=s['answer'],observed=s['value'],next_table=row['next_table'],scope=row['scope'],resolution='개발 분석용. 원문 정의 확인 전 미해결; 정답에 맞춘 자동 교체 금지'))
    docs=ROOT/'docs';docs.mkdir(exist_ok=True)
    dump(docs/'company_field_sources.json',book)
    for filename,rows in [('회사별_칼럼별_출처목록.csv',inventory),('2512_불일치_빈칸_분석.csv',ledger)]:
        with (docs/filename).open('w',encoding='utf-8-sig',newline='') as stream:
            writer=csv.DictWriter(stream,fieldnames=list(rows[0]));writer.writeheader();writer.writerows(rows)
    summary=dict(companies=len(book['companies']),company_field_entries=len(inventory),pdf_examples=len(manual),pdf_fields=len({(f['short'],f['field']) for f in manual}),factsheet_rules=sum(len(c['factsheet_locations']) for c in book['companies'].values()),scored_fields=len(scores),matched=sum(r['status']=='OK' for r in scores),issues=len(ledger),categories=dict(Counter(r['category'] for r in ledger)))
    lines=['# 회사별·칼럼별 출처 안내','', '2512 개발 검증에서 일치한 값과 불일치·빈칸을 모두 포함한 탐색 목록입니다. 정답 수치를 재사용하는 규칙이 아닙니다.', '',
       f"12개사 {summary['company_field_entries']}개 회사·칼럼 조합, PDF 직접 확인 {summary['pdf_fields']}개 조합({summary['pdf_examples']}개 근거), 기존 팩트시트 규칙 {summary['factsheet_rules']}개를 기록했습니다.", '',
       '정답 대조 대상 854개 중 일치 733개, 불일치 55개, 빈칸 66개를 포함합니다. 일치 여부와 출처 검증 수준은 별도입니다. XBRL 추출 묶음만 있는 항목은 정확한 태그·context 위치까지 확인한 것으로 표시하지 않았습니다.', '',
       '[전체 회사·칼럼 출처 목록](회사별_칼럼별_출처목록.csv) · [121개 불일치·빈칸 분석](2512_불일치_빈칸_분석.csv)', '',
       '## 사용 기준','', '페이지는 2512에서 찾았던 위치 예시입니다. 다음 분기에는 표 제목·행 이름·연도·분기 열·각주·단위를 다시 확인합니다. PDF 텍스트와 이미지가 다르면 이미지를 확인합니다. 숫자 반올림이나 발표 시점 차이만으로 기존 정밀값을 교체하지 않습니다.', '',
       '## 회사별 탐색 위치','']
    for company,short in COMPANIES:
        e=book['companies'][company];lines += [f'### {company}','']+['- '+n for n in e['notes']]+['',f"PDF 근거 {len(e['examples'])}개, 팩트시트 코드 규칙 {len(e['factsheet_locations'])}개.",'']
    lines += ['## 칼럼별 공통 탐색 기준','','| 칼럼 | 표·행 | 확인할 범위 |','|---|---|---|']
    lines += [f"| {','.join(map(str,g['fields']))} | {g['table']} / {g['rows']} | {g['scope']} |" for g in GROUPS]
    lines += ['', '## 검증 범위','', '2512 정답을 보고 보완하는 개발 단계입니다. 기존 블라인드 결과는 수정하지 않았으며, 다른 분기 성능 향상을 입증한 결과가 아닙니다. 위치 미확인·정의 미확정 항목은 계속 미해결로 남깁니다. PDF 코드 보완은 별도·단위·기간·표 구조 확인을 대체하지 않습니다.']
    (docs/'회사별_칼럼별_출처안내.md').write_text('\n'.join(lines)+'\n',encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--backtest',type=Path)
    args=parser.parse_args();build(args.backtest or Path((ROOT/'runs/active_backtest.txt').read_text(encoding='utf-8')))
