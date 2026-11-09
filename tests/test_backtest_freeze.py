import pytest
from backtest_2512 import freeze_outputs
from quarterly.core import load


def test_scoring_cannot_replace_frozen_results(tmp_path):
    (tmp_path/'run').mkdir()
    (tmp_path/'run/stage1.xlsx').write_bytes(b'first output')
    final=tmp_path/'run/final.xlsx';final.write_bytes(b'reviewed output')
    freeze_outputs(tmp_path,final)
    original=load(tmp_path/'frozen_outputs.json')
    freeze_outputs(tmp_path,final)
    assert load(tmp_path/'frozen_outputs.json')==original
    final.write_bytes(b'changed after answers were seen')
    with pytest.raises(ValueError,match='새 검증 실행'):
        freeze_outputs(tmp_path,final)
    assert load(tmp_path/'frozen_outputs.json')==original
