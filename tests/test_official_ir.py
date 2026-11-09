from quarterly.core import Period
from quarterly.official_ir import discover, matching
from quarterly.source_rules import company_catalog
from quarterly.sources import Collector


def test_exact_quarter_does_not_select_latest_or_other_annual():
    assert matching('FY2025 상반기 실적발표 자료', Period(2025,2))
    assert not matching('FY2025 3분기 실적발표 자료', Period(2025,4))
    assert not matching('FY2024 연간 실적발표 자료', Period(2025,1))
    assert not matching('2026년 2분기 경영실적', Period(2026,3))


def test_samsung_file_response_array_and_form_encoding(tmp_path, monkeypatch):
    from quarterly import official_ir
    c=Collector(tmp_path,tmp_path,Period(2025,1))
    calls=[]
    def post(collector,url,**kwargs):
        calls.append(kwargs)
        if url.endswith('/list'):
            return {'response':{'list':[
                {'bltnTitlCntnt':'FY2025 3분기 실적발표 자료','bltnItmCntnt01':'wrong'},
                {'bltnTitlCntnt':'FY2025 1분기 실적발표 자료','bltnItmCntnt01':'correct'}]}}
        assert kwargs['data']['cntsApndFileId']=='correct'
        return {'response':[{'fileDownUrl':'/dcms/q1.pdf'}]}
    monkeypatch.setattr(official_ir,'public_post',post)
    assert discover(c,'삼성생명')==[('https://www.samsunglife.com/dcms/q1.pdf','FY2025 1분기 실적발표 자료')]
    assert all('data' in call and 'json' not in call for call in calls)


def test_meritz_ignores_latest_top_level_files(tmp_path,monkeypatch):
    from quarterly import official_ir
    monkeypatch.setattr(official_ir,'public_post',lambda *a,**k:{'fileMap':[{'file_url':'wrong.pdf'}],
        'result':[{'title_ko':'2025년 2분기 경영실적','fileMap':[{'file_url':'/hld/q2.pdf','lang_div_cd':'P'}]},
                  {'title_ko':'2025년 1분기 경영실적','fileMap':[{'file_url':'/hld/q1.pdf','lang_div_cd':'P'}]}]})
    assert discover(Collector(tmp_path,tmp_path,Period(2025,1)),'메리츠화재')==[
        ('https://www.meritzgroup.com/commfiles/hld/q1.pdf','2025년 1분기 경영실적')]


def test_company_definition_does_not_change_other_fields_or_company():
    catalog={'75':{'basis':'separate'},'84':{'basis':'separate'}}
    config={'company_fields':{'교보생명':{'84':{'basis':'consolidated','maturity_bounds_years':[1,3]}}}}
    assert company_catalog(catalog,'교보생명',config)['84']['basis']=='consolidated'
    assert company_catalog(catalog,'교보생명',config)['75']['basis']=='separate'
    assert company_catalog(catalog,'삼성생명',config)['84']['basis']=='separate'
    assert catalog['84']=={'basis':'separate'}
