# -*- coding: utf-8 -*-
"""LibreOffice 로 수식을 실제 계산시킨 사본을 만든다.

openpyxl 로 저장한 파일에는 수식의 계산값이 남지 않아, 그대로 읽으면
수식 칸이 전부 빈칸으로 보인다. 채움률을 정직하게 세려면 한 번 계산시켜야 한다.
"""
import os, sys, subprocess, tempfile, shutil
sys.stdout.reconfigure(encoding='utf-8')


def recalc(src, dst=None, timeout=600):
    src = os.path.abspath(src)
    tmp = tempfile.mkdtemp(prefix='recalc_')
    env = dict(os.environ, HOME=tmp)
    subprocess.run(['soffice', '--headless', '--norestore',
                    '--convert-to', 'xlsx:Calc MS Excel 2007 XML',
                    '--outdir', tmp, src],
                   check=True, timeout=timeout, env=env,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out = os.path.join(tmp, os.path.basename(src))
    if not os.path.exists(out):
        raise RuntimeError('LibreOffice 변환 실패: ' + src)
    dst = dst or (os.path.splitext(src)[0] + '_calc.xlsx')
    shutil.copy(out, dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return dst


if __name__ == '__main__':
    print(recalc(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None))
