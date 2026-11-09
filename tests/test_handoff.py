import hashlib
import json
import zipfile
import pytest
from tools import export_handoff


def test_handoff_uses_working_files_and_excludes_private_state(tmp_path,monkeypatch):
    root=tmp_path/'source';root.mkdir()
    monkeypatch.setattr(export_handoff,'ROOT',root)
    monkeypatch.setattr(export_handoff,'FILES',['README.md','.env.example'])
    monkeypatch.setattr(export_handoff,'DOCS',['quarterly_lessons.json'])
    (root/'docs').mkdir();(root/'docs/quarterly_lessons.json').write_text('{}')
    for folder in ['tools','tests','quarterly']:(root/folder).mkdir()
    (root/'README.md').write_text('current uncommitted operator guide')
    (root/'.env.example').write_text('DART_API_KEY=')
    (root/'.env').write_text('DART_API_KEY=private-test-value')
    (root/'quarterly.local.json').write_text('{"workbook":"personal/path.xlsx"}')
    (root/'work.xlsx').write_bytes(b'original template bytes')
    (root/'answer.xlsx').write_bytes(b'original reference bytes')
    defaults={'workbook':'work.xlsx','reference_workbook':'answer.xlsx','use_fisis':False}
    (root/'quarterly.json').write_text(json.dumps(defaults))
    monkeypatch.setattr(export_handoff,'read_config',lambda _:defaults)
    output=export_handoff.export(tmp_path/'handoff.zip')
    with zipfile.ZipFile(output) as z:
        assert z.read('Dartapi/README.md')==b'current uncommitted operator guide'
        assert not any(name.endswith('/.env') or 'quarterly.local.json' in name for name in z.namelist())
        assert b'private-test-value' not in b''.join(z.read(name) for name in z.namelist())
        config=json.loads(z.read('Dartapi/quarterly.json'))
        assert config['workbook']=='작업양식.xlsx' and config['use_fisis'] is False
        assert z.read('Dartapi/작업양식.xlsx')==(root/'work.xlsx').read_bytes()
        manifest=json.loads(z.read('Dartapi/HANDOFF_MANIFEST.json'))
        assert manifest['files']['answer.xlsx']==hashlib.sha256((root/'answer.xlsx').read_bytes()).hexdigest()
    with pytest.raises(FileExistsError):export_handoff.export(output)
