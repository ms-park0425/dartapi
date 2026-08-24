# -*- coding: utf-8 -*-
"""IR 팩트시트 추출값을 정답 엑셀과 대조한다."""
import sys, openpyxl
sys.stdout.reconfigure(encoding='utf-8')
from ir_parse import extract_company, SPECS

ANSWER = {'2503': '202512_동업사 공시비교_DATA_작업_260320.xlsx',
          '2506': '202512_동업사 공시비교_DATA_작업_260320.xlsx',
          '2509': '202512_동업사 공시비교_DATA_작업_260320.xlsx',
          '2512': '202512_동업사 공시비교_DATA_작업_260320.xlsx'}
SHORT = {'미래에셋생명': '미래', '삼성생명': '삼성', '한화생명': '한화',
         '동양생명': '동양', '삼성화재': '삼성화재', '현대해상': '현대해상',
         '메리츠화재': '메리츠화재', 'KB손해보험': 'KB손보', 'KB생명': 'KB라이프',
         '교보생명': '교보', '신한라이프생명': '신한라이프', 'DB손해보험': 'DB손보'}

_cache = {}
def answers(path, period, short):
    key = (path, period, short)
    if key in _cache: return _cache[key]
    wb = _cache.get(path) or openpyxl.load_workbook(path, data_only=True)
    _cache[path] = wb
    ws = wb['DATA']
    col2ec = {}
    for ec in range(1, ws.max_column + 1):
        try: col2ec[int(ws.cell(1, ec).value)] = ec
        except (TypeError, ValueError): pass
    out = {}
    for r in range(6, ws.max_row + 1):
        if str(ws.cell(r, 2).value) == period and str(ws.cell(r, 3).value).strip() == short:
            for col, ec in col2ec.items():
                v = ws.cell(r, ec).value
                if isinstance(v, (int, float)): out[col] = v
            break
    _cache[key] = out
    return out


def run(company, factsheet, periods):
    short = SHORT[company]
    tot_ok = tot_fail = tot_na = 0
    for period in periods:
        vals, miss = extract_company(company, factsheet, period)
        ans = answers(ANSWER[period], period, short)
        print(f'--- {company} {period} ---')
        for col in sorted(vals):
            got, exp = vals[col], ans.get(col)
            if exp is None:
                print(f'  Col{col:<4} {got:>18,.3f}   (정답 없음)'); tot_na += 1; continue
            err = abs(got - exp) / max(abs(exp), 1e-9)
            mark = 'OK ' if err < 0.001 else 'FAIL'
            if err < 0.001: tot_ok += 1
            else: tot_fail += 1
            print(f'  Col{col:<4} {got:>18,.3f} vs {exp:>18,.3f}  {mark} ({err*100:.3f}%)')
        for c, l, why in miss:
            print(f'  Col{c:<4} -- 추출실패: {l} ({why})')
    print(f'\n>>> {company}: OK={tot_ok} FAIL={tot_fail} 정답없음={tot_na}')
    return tot_ok, tot_fail


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('company'); ap.add_argument('factsheet')
    ap.add_argument('--periods', default='2503,2506,2509,2512')
    a = ap.parse_args()
    run(a.company, a.factsheet, a.periods.split(','))
