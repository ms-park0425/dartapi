from pathlib import Path
from quarterly.core import Period
from quarterly.extract import extract_rules


def test_annual_parser_alias_uses_registered_xml_snapshot(tmp_path,monkeypatch):
    import parse_annual
    root=tmp_path/'sandbox';root.mkdir();run=tmp_path/'run';run.mkdir()
    snapshot=tmp_path/'source.xml';snapshot.write_text('<DOCUMENT>registered current-year snapshot</DOCUMENT>',encoding='utf-8')
    def extract(*args,**kwargs):
        alias=Path('data/미래에셋생명/doc_annual_2025.xml')
        assert alias.read_bytes()==snapshot.read_bytes()
        return {}
    monkeypatch.setattr(parse_annual,'extract_company_data',extract)
    sources=[dict(company='미래에셋생명',kind='document_xml',id='xml',path=str(snapshot)),
             dict(company='미래에셋생명',kind='xbrl',id='xbrl',path='unused.xbrl')]
    assert extract_rules(root,run,Period(2025,4),[('미래에셋생명','미래')],sources,offline=True)==[]
