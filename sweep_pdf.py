# -*- coding: utf-8 -*-
"""PDF(경영공시·IR 발표자료)를 정답값으로 역탐색한다."""
import sys, argparse, os
sys.stdout.reconfigure(encoding='utf-8')
from ir_discover import answer_row, discover_pdf, SHORT, LOOSE, TOL
from sweep_miss import miss_cols

ap = argparse.ArgumentParser()
ap.add_argument('pdf'); ap.add_argument('short'); ap.add_argument('period')
ap.add_argument('--loose', action='store_true')
ap.add_argument('--all', action='store_true', help='MISS 가 아닌 컬럼도 포함')
a = ap.parse_args()

ans, names = answer_row(a.period, a.short)
if not a.all:
    miss = miss_cols(a.period, a.short)
    ans = {c: v for c, v in ans.items() if c in miss}
ans = {c: v for c, v in ans.items() if v}
print(f'=== {a.short} {a.period} 대상 {len(ans)}개 컬럼  ({os.path.basename(a.pdf)})')
hits = discover_pdf(a.pdf, ans, tol=LOOSE if a.loose else TOL, names=names)
for col in sorted(hits):
    print(f'\nCol{col} {names.get(col,"")[:55]} = {ans[col]:,.3f}')
    seen = set()
    for pno, ln, line, un, sg in hits[col]:
        if line in seen:
            continue
        seen.add(line)
        print(f'    p{pno} L{ln} {un}{"" if sg==1 else " 부호반전"}  {line}')
        if len(seen) >= 4:
            break
print('\n못 찾은 컬럼:', ', '.join(f'Col{c}' for c in sorted(set(ans)-set(hits))))
