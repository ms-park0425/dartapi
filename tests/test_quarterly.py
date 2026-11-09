from pathlib import Path
import shutil

import openpyxl
import pytest

from quarterly.core import Period, digest, dump, load
from quarterly.workbook import DataBook, validate
from quarterly.review import check_finding, field_catalog
from quarterly.sources import Collector, links_in, period_hint
from quarterly.pipeline import finalize, prepare, review_run


@pytest.fixture
def workbook(tmp_path):
    p = tmp_path / "work.xlsx"
    w = openpyxl.Workbook()
    s = w.active
    s.title = "DATA"
    for c in range(1, 118):
        s.cell(1, c).value = c
    for c, label in {4:"총자산", 9:"FVPL 비중", 12:"부채", 15:"자본", 32:"보험손익", 75:"CSM", 76:"신계약CSM", 77:"CSM상각", 78:"CSM조정", 79:"이자부리", 80:"전기CSM", 111:"예실차", 112:"보험금예실차", 113:"예상보험금", 114:"실제보험금", 115:"사업비예실차", 116:"예상사업비", 117:"실제사업비"}.items():
        s.cell(5,c).value = label
    for row, period, name in [(6,2603,"미래"),(7,2606,"미래"),(10,2603,"삼성"),(11,2606,"=C10")]:
        s.cell(row,1).value = f"{period}key"
        s.cell(row,2).value = period
        s.cell(row,3).value = name
        s.cell(row,4).value = 123
        s.cell(row,9).value = f"=E{row}/D{row}"
        s.cell(row,75).value = 10
    w.create_sheet("보고")["A1"] = "=DATA!D7"
    w.save(p)
    return p


def finding(**changes):
    result = dict(field=75, decision="CORRECT", value=2000, source_id="pdf1", page=1,
                  value_status="not_applicable",
                  entity="미래에셋생명", basis="separate", period_type="instant", period_start=None,
                  period_end="2026-09-30", raw_values=[2], raw_unit="KRW_billion", operation="direct", scale=1000,
                  display_decimals=0, table="CSM", row="기말", column="2026.09", quote="CSM 기말 2 십억원", reason="동일 회사·기준일의 별도 CSM 잔액",
                  correction_proof=dict(same_measure=True,existing_value=10,error_type="unit",detail="단위 오류를 확인한 테스트 근거"))
    result.update(changes)
    return result


def catalog_and_assets(book):
    return field_catalog(book), [dict(source_id="pdf1",page=1,company="미래에셋생명")]


def test_period_and_exact_report():
    p = Period(2026,3)
    assert p.code == "2609" and p.end.isoformat() == "2026-09-30" and p.report_code == "11014"
    assert p.matches_report("[기재정정]분기보고서 (2026.09)")
    assert not p.matches_report("분기보고서 (2026.03)")
    assert not p.matches_report("분기보고서 (2025.09)")
    with pytest.raises(ValueError): Period.parse("2608")


def test_new_period_keeps_history_and_translates_formula(workbook,tmp_path):
    original = digest(workbook)
    book = DataBook(workbook)
    book.ensure_period(Period(2026,3),[("미래에셋생명","미래"),("삼성생명","삼성")])
    assert book.rows[("2609","미래")] == 8
    assert book.rows[("2609","삼성")] == 12
    assert book.values("2609","미래")[4] is None
    assert book.values("2609","미래")[9] == "=E8/D8"
    assert book.values("2606","미래")[4] == 123
    book.apply("2609","미래",4,0)
    out = tmp_path / "out.xlsx"
    book.save(out)
    assert DataBook(out).values("2609","미래")[4] == 0
    assert digest(workbook) == original
    assert openpyxl.load_workbook(out)["보고"]["A1"].value == "=DATA!D7"


def test_occupied_spacer_appends_without_shifting(workbook):
    b = DataBook(workbook)
    b.ws.cell(8,4).value = "keep this note"
    b.ensure_period(Period(2026,3),[("미래에셋생명","미래")])
    assert b.rows[("2606","삼성")] == 11
    assert b.rows[("2609","미래")] > 11
    assert b.ws.cell(8,4).value == "keep this note"


def test_duplicate_and_formula_protection(workbook):
    b=DataBook(workbook)
    with pytest.raises(ValueError,match="수식 보호"): b.apply("2606","미래",9,1)
    with pytest.raises(ValueError,match="기존값 변경"): b.apply("2606","미래",4,500,expected=0)
    with pytest.raises(ValueError): b.apply("2606","미래",4,float("nan"))
    b.ws.cell(8,2).value=2606
    b.ws.cell(8,3).value="미래"
    with pytest.raises(ValueError,match="중복"): b.refresh()


@pytest.mark.parametrize("changes,error",[
    ({"basis":"consolidated"},"회계 기준"),
    ({"period_end":"2026-06-30"},"기준일"),
    ({"period_type":"quarter"},"기간 성격"),
    ({"scale":100},"단위 환산"),
    ({"value":3000},"변환값"),
    ({"source_id":"invented"},"PDF 근거"),
    ({"entity":"삼성생명"},"회사"),
    ({"quote":""},"근거 누락"),
])
def test_review_rejects_wrong_definition_period_unit_source(workbook,changes,error):
    c,a=catalog_and_assets(DataBook(workbook))
    with pytest.raises(ValueError,match=error): check_finding(finding(**changes),Period(2026,3),c,a,10,"미래에셋생명","미래")


def test_ytd_and_opening_contracts(workbook):
    c,a=catalog_and_assets(DataBook(workbook))
    with pytest.raises(ValueError,match="연초 누계"):
        check_finding(finding(field=76,period_type="ytd",period_start="2026-07-01"),Period(2026,3),c,a,10,"미래에셋생명","미래")
    f=finding(field=80,period_type="opening",period_end="2025-12-31")
    assert check_finding(f,Period(2026,3),c,a,10,"미래에셋생명","미래")["value"]==2000


def test_zero_is_real_value_and_rounding_keeps_precision(workbook):
    c,a=catalog_and_assets(DataBook(workbook))
    zero=finding(raw_values=[0],value=0,quote="CSM 기말 0 십억원")
    assert check_finding(zero,Period(2026,3),c,a,None,"미래에셋생명","미래")["value"]==0
    assert check_finding(finding(decision="KEEP"),Period(2026,3),c,a,2010,"미래에셋생명","미래")["value"]==2010
    with pytest.raises(ValueError,match="수식"):
        check_finding(finding(),Period(2026,3),c,a,"=A1","미래에셋생명","미래")


def test_ratio_conversion(workbook):
    c,a=catalog_and_assets(DataBook(workbook))
    f=finding(field=9,value=0.25,raw_values=[25],raw_unit="percent",scale=0.01,quote="비중 25%")
    f['correction_proof']['existing_value']=0.1
    assert check_finding(f,Period(2026,3),c,a,0.1,"미래에셋생명","미래")["value"]==0.25


def test_public_links_and_quarter_filter():
    html='<a href="/files/a.pdf">FY2026 3Q Results</a><a onclick="dl(\'/files/b.xlsx\')">2026년 3분기</a>'
    links=links_in(html,"https://example.com/ir")
    assert ("https://example.com/files/a.pdf","FY2026 3Q Results") in links
    assert ("https://example.com/files/b.xlsx","2026년 3분기") in links
    assert all(period_hint(label,Period(2026,3)) for _,label in links)
    assert not period_hint("FY2026 1Q",Period(2026,3))
    assert not period_hint("2026",Period(2026,3))


def test_dart_pagination_selects_corrected_target_not_other_quarter(tmp_path):
    c=Collector(tmp_path,tmp_path,Period(2026,3))
    seen=[]
    class Response:
        content=b"not zip"
        text=""
        def __init__(self,data): self.data=data
        def json(self): return self.data
    def get(url,**kw):
        seen.append((url,kw))
        if url.endswith("list.json"):
            page=kw["params"]["page_no"]
            return Response(dict(status="000",total_page=2,list=[dict(report_nm="분기보고서 (2026.03)",rcept_no="bad",rcept_dt="20261130")]
                if page==1 else [dict(report_nm="[기재정정]분기보고서 (2026.09)",rcept_no="good",rcept_dt="20261120")]))
        return Response({})
    c.get=get
    c.dart("미래에셋생명","code","secret")
    assert len([u for u,k in seen if u.endswith("list.json")])==2
    calls=[k["params"] for u,k in seen if u.endswith("fnlttXbrl.xml")]
    assert calls[0]["rcept_no"]=="good" and calls[0]["reprt_code"]=="11014"


def prepared_run(tmp_path,workbook,monkeypatch):
    import quarterly.pipeline as p
    monkeypatch.setattr(p.Collector,"cached",lambda *a:None)
    monkeypatch.setattr(p.Collector,"dart",lambda *a:None)
    monkeypatch.setattr(p.Collector,"ir",lambda *a:None)
    monkeypatch.setattr(p,"extract_rules",lambda *a,**kw:[dict(company="미래에셋생명",short="미래",field=75,value=1000)])
    def assets(run,sources,companies,limit):
        pdf=run/"source.pdf";pdf.write_bytes(b"%PDF test evidence")
        image=run/"page.png";image.write_bytes(b"rendered evidence fixture")
        text=run/"page.txt";text.write_text("CSM 기말 2 십억원",encoding="utf-8")
        dump(run/"source_manifest.json",[dict(path=str(pdf),sha256=digest(pdf))])
        data=[dict(source_id="pdf1",page=1,company="미래에셋생명",short="미래",image=str(image),image_sha256=digest(image),
                   pdf_path=str(pdf),pdf_sha256=digest(pdf),text=str(text))]
        dump(run/"review_assets.json",data)
        dump(run/"review_asset_issues.json",[])
    monkeypatch.setattr(p,"prepare_assets",assets)
    monkeypatch.setattr(p,"recalculate",lambda *a:"pending_excel_recalculation")
    return prepare(tmp_path,Period(2026,3),dict(workbook=str(workbook)),offline=True,company_filter=["미래"],log=lambda *a:None)


def test_end_to_end_model_result_changes_existing_value_and_is_idempotent(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    class Reviewer:
        def review(self,*args): return [finding(correction_proof=dict(same_measure=True,existing_value=1000,error_type="unit",detail="원본 표 단위와 기존 후보의 단위 환산을 비교해 1000배 오류 확인"))]
    output=review_run(run,Reviewer(),log=lambda *a:None)
    assert DataBook(run/"stage1.xlsx").values("2609","미래")[75]==1000
    assert DataBook(output).values("2609","미래")[75]==2000
    assert DataBook(output).values("2606","미래")[75]==10
    assert load(run/"changes.json")[0]["old"]==1000
    before=digest(output)
    finalize(run,log=lambda *a:None)
    assert digest(output)==before
    assert load(run/"summary.json")["reviewed_pages"]==1
    assert load(run/"state.json")["status"]=="completed_with_unresolved"


def test_conflicting_pdf_candidates_not_applied(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    dump(run/"review_results.json",dict(findings=[dict(**finding(),short="미래"),dict(**finding(value=5000,raw_values=[5],quote="기말 5 십억원"),short="미래")]))
    output=finalize(run,log=lambda *a:None)
    assert DataBook(output).values("2609","미래")[75]==1000
    assert not load(run/"changes.json")


def test_compatible_pdf_sources_choose_more_precise_value(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    proof=dict(same_measure=True,existing_value=1000,error_type='unit',detail='같은 항목 원문 단위 오류 확인')
    coarse=finding(correction_proof=proof)
    fine=finding(value=1999,raw_values=[1.999],display_decimals=3,quote='CSM 기말 1.999 십억원',correction_proof=proof)
    dump(run/'review_results.json',dict(findings=[dict(**coarse,short='미래'),dict(**fine,short='미래')]))
    output=finalize(run,log=lambda *a:None)
    assert DataBook(output).values('2609','미래')[75]==1999


def test_changed_source_or_workbook_rejects_application(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    dump(run/"review_results.json",dict(findings=[dict(**finding(),short="미래")]))
    (run/"source.pdf").write_bytes(b"changed")
    with pytest.raises(ValueError,match="근거가 변경"):
        finalize(run,log=lambda *a:None)


def test_cumulative_summary_not_added_to_q1_again(tmp_path):
    from ir_parse import Item, extract, CUM, POINT
    w=openpyxl.Workbook();s=w.active;s.title="EFF"
    s.cell(3,4).value="YTD"
    for c,y,q in [(4,2026,"Q2"),(5,2025,"Q2"),(9,2026,"Q2"),(10,None,"Q1"),(11,2025,"FY"),(12,None,"Q4"),(13,None,"Q3"),(14,None,"Q2"),(15,None,"Q1")]:
        s.cell(4,c).value=y;s.cell(5,c).value=q
    s.cell(7,2).value="보험금"
    for c,v in [(4,275),(5,220),(9,139),(10,136)]:s.cell(7,c).value=v
    path=tmp_path/"factsheet.xlsx";w.save(path)
    evidence=[]
    values,missing=extract(path,"2606",[Item("EFF","보험금",109,CUM),Item("EFF","보험금",110,POINT)],unit=1000,evidence=evidence)
    assert values[109]==275000  # Q1 136 + Q2 139, not YTD 275 + Q1 136.
    assert values[110]==275000  # POINT still selects the summary column.
    assert [x["cell"] for x in evidence[0]["cells"]]==["J7","I7"]


def test_provisional_kics_cannot_fill_final_column(workbook):
    book=DataBook(workbook)
    book.ws.cell(5,95).value="K-ICS 최종"
    catalog=field_catalog(book)
    assets=[dict(source_id="pdf1",page=1,company="미래에셋생명")]
    f=finding(field=95,basis="supervisory",value_status="provisional",raw_values=[155.3],raw_unit="percent",scale=0.01,value=1.553,quote="K-ICS 155.3% (e)")
    with pytest.raises(ValueError,match="잠정/확정"):
        check_finding(f,Period(2026,3),catalog,assets,None,"미래에셋생명","미래")


def test_quote_must_contain_observed_number(workbook):
    c,a=catalog_and_assets(DataBook(workbook))
    with pytest.raises(ValueError,match="원문 수치 없음"):
        check_finding(finding(quote="그냥 맞다고 판단함"),Period(2026,3),c,a,10,"미래에셋생명","미래")


def test_pdf_text_mismatch_requires_image_bound_transcription(workbook,tmp_path):
    c,a=catalog_and_assets(DataBook(workbook))
    text=tmp_path/'page.txt';text.write_text('123456789012345678901234567890',encoding='utf-8')
    image=tmp_path/'page.png';image.write_bytes(b'fixed rendered page')
    a[0].update(text=str(text),image=str(image),image_sha256=digest(image))
    f=finding()
    with pytest.raises(ValueError,match='이미지 재확인'):
        check_finding(f,Period(2026,3),c,a,10,'미래에셋생명','미래')
    a[0]['visual_transcriptions']=[dict(raw_values=[2],quote=f['quote'],image_sha256=digest(image),reason='full page visually checked')]
    assert check_finding(f,Period(2026,3),c,a,10,'미래에셋생명','미래')['value']==2000
    image.write_bytes(b'changed page')
    with pytest.raises(ValueError,match='이미지 재확인'):
        check_finding(f,Period(2026,3),c,a,10,'미래에셋생명','미래')


def test_fisis_keeps_disclosed_zero(monkeypatch):
    import dart_api
    class Response:
        def json(self):return dict(result=dict(list=[dict(account_cd="A113",a="0"),dict(account_cd="A112",a="0"),dict(account_cd="A122",a="0")]))
    monkeypatch.setattr(dart_api,"_get",lambda *a,**kw:Response())
    values=dart_api.fetch_fisis_balance("미래에셋생명","202609")
    assert values[75]==0 and values[74]==0


def test_catalog_treats_paa_as_amount_and_loss_ratio_as_ytd(workbook):
    b=DataBook(workbook);b.ws.cell(5,81).value="보험료배분접근법";b.ws.cell(5,107).value="손해율(연간누계)"
    c=field_catalog(b)
    assert c['81']['unit']=='KRW_million'
    assert c['107']['unit']=='ratio' and c['107']['period_type']=='ytd'


def test_hold_correction_that_breaks_previously_closed_csm(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    b=DataBook(run/'stage1.xlsx')
    for field,value in {75:1000,80:1000,76:0,77:0,78:0,79:0}.items():b.apply('2609','미래',field,value)
    b.save(run/'stage1.xlsx')
    state=load(run/'state.json');state['stage1_sha256']=digest(run/'stage1.xlsx');dump(run/'state.json',state)
    f=finding(correction_proof=dict(same_measure=True,existing_value=1000,error_type='unit',detail='원본 표와 단위 환산 오류 대조'))
    dump(run/'review_results.json',dict(findings=[dict(**f,short='미래')]))
    output=finalize(run,log=lambda *a:None)
    assert DataBook(output).values('2609','미래')[75]==1000
    assert load(run/'summary.json')['held_changes']==1
    assert not load(run/'changes.json')


def test_precision_kept_and_unexplained_pdf_difference_held(workbook):
    c,a=catalog_and_assets(DataBook(workbook))
    f=finding(correction_proof=None)
    kept=check_finding(f,Period(2026,3),c,a,2010,'미래에셋생명','미래')
    assert kept['decision']=='KEEP' and kept['value']==2010
    held=check_finding(f,Period(2026,3),c,a,1000,'미래에셋생명','미래')
    assert held['decision']=='CONFLICT' and held['value']==1000
    assert check_finding(f,Period(2026,3),c,a,None,'미래에셋생명','미래')['decision']=='FILL'
    stale=finding(correction_proof=dict(same_measure=True,existing_value=999,error_type='unit',detail='구체적 오류 근거'))
    assert check_finding(stale,Period(2026,3),c,a,1000,'미래에셋생명','미래')['decision']=='CONFLICT'


def test_unproven_existing_value_is_preserved_and_not_marked_verified(tmp_path,workbook,monkeypatch):
    run=prepared_run(tmp_path,workbook,monkeypatch)
    dump(run/'review_results.json',dict(findings=[dict(**finding(correction_proof=None),short='미래')]))
    output=finalize(run,log=lambda *a:None)
    assert DataBook(output).values('2609','미래')[75]==1000
    assert not load(run/'changes.json')
    assert load(run/'summary.json')['pdf_verified_fields']==0
    assert '기존값 유지' in (run/'unresolved.csv').read_text(encoding='utf-8-sig')


def test_historical_values_and_formulas_are_locked(workbook,tmp_path):
    w=openpyxl.load_workbook(workbook);s=w['DATA']
    s.cell(13,2).value=2512;s.cell(13,3).value='미래';s.cell(13,4).value=100;s.cell(13,9).value='=E13/D13';w.save(workbook)
    b=DataBook(workbook)
    with pytest.raises(ValueError,match='과거 정답 기간 보호'):b.apply('2512','미래',4,999)
    b.ws.cell(13,4).value=999
    with pytest.raises(ValueError,match='과거 값'):b.save(tmp_path/'bad.xlsx')


def test_history_guard_accepts_only_machine_float_serialization_noise(workbook,tmp_path):
    b=DataBook(workbook,protect_through='2606')
    b.ws.cell(7,4).value=286199.16105600004
    snapshot=b.history_snapshot()
    b.protected_snapshot=snapshot
    path=tmp_path/'float.xlsx';b.save(path)
    saved=DataBook(path,protect_through='2606')
    saved.assert_history_unchanged(snapshot)
    saved.ws.cell(7,4).value+=0.001
    with pytest.raises(ValueError,match='과거'):
        saved.assert_history_unchanged(snapshot)


def test_baseline_comparison_joins_keys_not_row_numbers(workbook,tmp_path):
    from quarterly.history import read_reference,compare
    w=openpyxl.load_workbook(workbook);s=w['DATA'];s.cell(13,2).value=2512;s.cell(13,3).value='미래';s.cell(13,4).value=100;w.save(workbook)
    first,_=read_reference(workbook)
    s.cell(14,2).value=2512;s.cell(14,3).value='미래';s.cell(14,4).value=100
    for c in range(1,118):s.cell(13,c).value=None
    shifted=tmp_path/'shifted.xlsx';w.save(shifted)
    second,_=read_reference(shifted)
    assert first==second
    assert compare(second[('2512','미래',4)],first[('2512','미래',4)],4)[0]=='OK'


def test_api_error_does_not_expose_request_credentials(monkeypatch):
    import dart_api
    import requests
    class Session:
        def get(self,*args,**kwargs):
            raise requests.ConnectionError('https://example.test?crtfc_key=private-test-value')
    monkeypatch.delenv('DART_OFFLINE',raising=False)
    monkeypatch.setattr(dart_api.requests,'Session',Session)
    with pytest.raises(RuntimeError) as error:
        dart_api._get('https://opendart.fss.or.kr/api/list.json',params={'crtfc_key':'private-test-value'})
    assert 'private-test-value' not in str(error.value)
    assert error.value.__suppress_context__ is True
