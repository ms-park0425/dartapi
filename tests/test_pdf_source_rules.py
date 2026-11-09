import json
from pathlib import Path

import ir_pdf
from quarterly.core import COMPANIES
from quarterly.source_rules import reviewer_guide


def test_parentheses_and_spaced_korean_labels():
    assert ir_pdf._nums('물량차이 13,027 (711) (12,318)') == [13027,-711,-12318]
    assert ir_pdf._row_after(['보 험 손 익 7 ,042 7 ,042 -'],0,'보험손익') == [7042,7042]


def test_gmm_only_and_previous_year_are_distinct():
    text='회계모형별 보험부채 현황\n(단위 : 억원)\n<2025년>\n일반모형\n합 계 81,469 19,697 122,053 30,876\n<2024년>\n일반모형\n합계 91,083 18,665 122,318 30,235'
    assert ir_pdf.disclosure_csm_stocks(text,2025)=={75:12205300,80:12231800}
    # Do not populate a future period from a prior-year table.
    assert ir_pdf.disclosure_csm_stocks(text,2026)=={80:12205300}


def test_vfa_is_summed_without_paa_or_lrc_bel_ra():
    text='회계모형별 보험부채 현황\n단위 : 억원\n2025년\n일반모형 변동수수료접근법\n합계 327,903 13,601 74,205 43,834 215 1,332 9'
    assert ir_pdf.disclosure_csm_stocks(text,2025)=={75:7553700}
    assert ir_pdf.disclosure_csm_stocks(text.replace('2025년','기준연도 미상'),2025)=={}


def test_disclosure_continuation_and_new_contract_csm(monkeypatch):
    texts=[
        '회계모형별 보험부채 현황\n단위 : 억원\n<2025년>\n일반모형',
        '합계 81,469 19,697 122,053 30,876\n<2024년>\n일반모형\n합계 91,083 18,665 122,318 30,235',
        '계리적 가정\n<2025년>\n가정변경효과\n신계약 (32,250) 3,334 29,327\n해지율 가정 변경 4,429 (4,429)',
    ]
    class Page:
        def __init__(self,text):self.text=text
        def extract_text(self):return self.text
    class PDF:
        pages=[Page(t) for t in texts]
        def __enter__(self):return self
        def __exit__(self,*args):pass
    monkeypatch.setattr(ir_pdf.pdfplumber,'open',lambda path:PDF())
    values,evidence=ir_pdf.parse_disclosure('fixture.pdf',2025)
    assert values[75]==12205300 and values[80]==12231800
    assert values[76]==2932700
    assert values[153]==-442900


def test_all_companies_have_guides_and_no_answer_values():
    for company,_ in COMPANIES:
        guide=reviewer_guide(company,{'75':{},'76':{},'95':{}})
        assert guide['company_notes']
        assert guide['field_groups']
        assert all(e['field'] in {75,76,95} for e in guide['historical_navigation_examples'])
        for e in guide['historical_navigation_examples']:
            assert 'value' not in e and 'raw_values' not in e and 'answer' not in e


def test_proven_sign_error_preserves_existing_precision():
    from quarterly.review import check_finding
    from quarterly.core import Period
    f=dict(field=78,decision='CORRECT',value=-390000,source_id='s',page=1,entity='삼성생명',basis='separate',value_status='not_applicable',period_type='ytd',period_start='2025-01-01',period_end='2025-09-30',raw_values=[-.39],raw_unit='KRW_trillion',operation='direct',scale=1000000,display_decimals=2,table='CSM',row='조정',column='YTD',quote='조정 -0.39 조원',reason='동일 별도 CSM 조정 부호 확인',correction_proof=dict(same_measure=True,existing_value=393576,error_type='sign',detail='원문 조정 음수, 현재값 양수, 동일 범위 검산 불일치'))
    catalog={'78':dict(basis='separate',value_status='not_applicable',period_type='ytd',unit='KRW_million')}
    assets=[dict(source_id='s',page=1,company='삼성생명')]
    result=check_finding(f,Period(2025,3),catalog,assets,393576,'삼성생명','삼성')
    assert result['decision']=='CORRECT' and result['value']==-393576


def test_prompt_contains_only_requested_company_navigation(tmp_path):
    from quarterly.review import prompt_for
    from quarterly.core import Period
    text=tmp_path/'page.txt';text.write_text('CSM',encoding='utf-8')
    prompt=prompt_for([dict(company='KB생명',short='KB라이프',source_id='source',page=1,pdf_path='source.pdf',text=str(text))],Period(2026,3),{'75':{}},{'75':None})
    context=json.loads(prompt[prompt.index('{"company"'):])
    assert context['source_navigation']['company_notes']
    assert '이자+기타' in '\n'.join(context['source_navigation']['company_notes'])
