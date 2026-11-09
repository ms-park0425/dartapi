from quarterly.core import dump
import quarter_pipeline


def test_local_workbook_preference_does_not_override_explicit_config(tmp_path,monkeypatch):
    monkeypatch.setattr(quarter_pipeline,'ROOT',tmp_path)
    dump(tmp_path/'quarterly.json',{'workbook':'default.xlsx','use_fisis':False})
    dump(tmp_path/'quarterly.local.json',{'workbook':'user_latest.xlsx'})
    dump(tmp_path/'backtest.json',{'workbook':'masked.xlsx'})
    assert quarter_pipeline.read_config(tmp_path/'quarterly.json')=={'workbook':'user_latest.xlsx','use_fisis':False}
    assert quarter_pipeline.read_config(tmp_path/'backtest.json')=={'workbook':'masked.xlsx'}
