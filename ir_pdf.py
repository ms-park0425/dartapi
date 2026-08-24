# -*- coding: utf-8 -*-
"""IR/경영공시 PDF 에서 DATA 시트 컬럼을 추출한다.

팩트시트(xlsx)를 내지 않는 회사용. 현재 교보생명 정기경영공시를 지원한다.

PDF 텍스트에는 숫자 중간에 공백이 끼어 나오는 경우가 많다("1 42,185" = 142,185).
`num()` 이 그 공백을 걷어낸다.
"""
import re
import pdfplumber

EOK = 100.0   # 억원 → 백만원


def num(tok):
    t = str(tok).replace(',', '').replace('△', '-').replace('▲', '-').strip()
    try:
        return float(t)
    except ValueError:
        return None


_TOKEN = re.compile(r'-?[\d,]+(?:\.\d+)?')


def _nums(line):
    """한 줄에서 숫자를 순서대로 뽑는다.

    PDF 텍스트는 자릿수 구분점 앞에서 줄이 쪼개져 '1 42,185'(=142,185),
    '1 ,083'(=1,083) 처럼 나오는 일이 잦다. 앞 토큰이 콤마 없는 1~2자리이고
    뒤 토큰이 콤마로 시작하거나 콤마 앞이 1~2자리일 때만 붙인다
    (콤마가 이미 있는 '3,758 3,397' 같은 정상 배열은 건드리지 않는다).
    """
    toks = _TOKEN.findall(line)
    merged, i = [], 0
    while i < len(toks):
        a = toks[i]
        b = toks[i + 1] if i + 1 < len(toks) else None
        # 뒤 토큰이 부호로 시작하면 별개의 값이다('5,200 8 -5,208' 을 붙이면 안 된다)
        if (b and not b.startswith('-') and a.isdigit() and len(a) <= 2 and
                (b.startswith(',') or (',' in b and len(b.split(',')[0]) <= 2))
                and num(a + b) is not None):
            merged.append(a + b)
            i += 2
        else:
            merged.append(a)
            i += 1
    return [v for v in (num(t) for t in merged) if v is not None]


def find_pages(pdf, *needles):
    """조건을 만족하는 페이지 (index, text) 를 전부 돌려준다."""
    out = []
    for i in range(len(pdf.pages)):
        t = pdf.pages[i].extract_text() or ''
        if all(n in t for n in needles):
            out.append((i, t))
    return out


# ---------------------------------------------------------------- 교보생명
# CSM 상각 만기 버킷 → DATA 컬럼 (표는 1년,2년,…,10년,11~15,16~20,21~25,26~30,30년이후,계)
KYOBO_CSM_BUCKETS = {
    83: [0],                 # 1년 이하
    84: [1],                 # 1~2년
    85: [2, 3, 4],           # 2~5년
    86: [5, 6, 7, 8, 9],     # 5~10년
    87: [10, 11, 12, 13, 14],  # 10년 초과
    88: [15],                # 합계
}


def parse_kyobo(path):
    """교보생명 정기경영공시 PDF → ({col: 백만원}, 참고메모)"""
    out, notes = {}, []
    with pdfplumber.open(path) as pdf:
        # ① CSM 상각 만기별 → Col83~88 (표는 여러 페이지에 걸치고 마지막에 '합계' 행)
        total = None
        for i, _ in find_pages(pdf, '보험계약마진 상각'):
            for pno in range(i, min(i + 3, len(pdf.pages))):
                for line in (pdf.pages[pno].extract_text() or '').split('\n'):
                    if line.strip().startswith('합계'):
                        v = _nums(line)
                        if len(v) >= 16:
                            total = v[:16]
                            break
                if total:
                    break
            if total:
                break
        if total:
            for col, idx in KYOBO_CSM_BUCKETS.items():
                out[col] = sum(total[k] for k in idx) * EOK
        else:
            notes.append('CSM 상각 만기별 합계 행을 찾지 못함')

        # ② K-ICS (경과조치 적용 전) → Col95~97
        avail = req = None
        for i, t in find_pages(pdf, '경과조치 적용 전 지급여력비율 세부'):
            lines = t.split('\n')
            for n, s2 in enumerate(lines):
                st = s2.strip()
                if avail is None and st.startswith('가. 지급여력금액'):
                    for k in range(n, min(n + 3, len(lines))):
                        v = _nums(lines[k])
                        if len(v) >= 3:
                            avail = v[0]
                            break
                if req is None and st.startswith('나. 지급여력기준금액'):
                    v = _nums(st)
                    if len(v) >= 3:
                        req = v[-3]
            if avail and req:
                break
        if avail:
            out[96] = avail * EOK
        if req:
            out[97] = req * EOK
        if avail and req:
            out[95] = avail / req * 100
        else:
            notes.append('K-ICS 가용/요구자본 파싱 실패')

        # ③ 해약환급금준비금 → Col99
        for i, t in find_pages(pdf, '해약환급금준비금 등의 적립'):
            for line in t.split('\n'):
                if line.strip().startswith('해약환급금준비금'):
                    v = _nums(line)
                    if v:
                        out[99] = v[0] * EOK
                    break
            if 99 in out:
                break
        if 99 not in out:
            notes.append('해약환급금준비금 파싱 실패')
    return out, notes


PARSERS = {'교보생명': parse_kyobo}




# ------------------------------------------------------------------
# 금감원 정기경영공시 PDF (모든 보험사가 같은 목차·같은 표 제목을 쓴다)
# ------------------------------------------------------------------
DISC_UNIT = 100.0          # 경영공시 본문은 억원 → DATA(백만원)는 ×100


def _row_after(lines, start, label, n=None):
    """start 이후 label 로 시작하는 줄의 숫자들

    항목번호('(3) …', '가. …')가 숫자로 잡히지 않도록 라벨 부분은 잘라내고 읽는다.
    """
    for i in range(start, len(lines)):
        s = lines[i].strip()
        if s.startswith(label):
            v = _nums(s[len(label):])
            return v[:n] if n else v
    return []


def _block(lines, title, end_titles=(), span=80):
    """title 이 나오는 줄부터 다음 제목 전까지"""
    for i, ln in enumerate(lines):
        if title in ln:
            j = len(lines)
            for k in range(i + 1, min(i + span, len(lines))):
                if any(e in lines[k] for e in end_titles):
                    j = k
                    break
            return lines[i:j]
    return []


def parse_disclosure(path, year=None):
    """정기경영공시 PDF → {DATA 컬럼번호: 백만원}

    표 제목이 표준화돼 있어 회사가 달라도 같은 코드로 읽힌다.
    각 표에서 '해당 연도' 열(첫 숫자 열)만 쓴다.
    """
    out, src = {}, {}

    def put(col, val, where):
        if val is None:
            return
        out[col] = val * DISC_UNIT
        src[col] = where

    with pdfplumber.open(path) as pdf:
        pages = [(i + 1, p.extract_text() or '') for i, p in enumerate(pdf.pages)]

    # ---- ① 회계모형별·포트폴리오별 보험부채 현황 → CSM 잔액(Col75)
    for pno, txt in pages:
        if '회계모형별' in txt and '보험부채 현황' in txt:
            lines = txt.split('\n')
            # <2025년> 블록이 먼저 나온다. 그 안의 '합계' 행.
            blk = _block(lines, '<%d년>' % year, ('<%d년>' % (year - 1),)) if year else lines
            v = _row_after(blk or lines, 0, '합계')
            if len(v) >= 6:
                # 일반모형(최선추정·위험조정·보험계약마진) + VFA(동일 3열)
                put(75, v[2] + v[5], f'p{pno} 보험부채현황 합계')
            break

    # ---- ② 계리적 가정 → 가정변경(BEL기준 Col146~150 / CSM기준 Col153~157)
    for pno, txt in pages:
        if '계리적 가정' in txt and '가정변경효과' in txt:
            lines = txt.split('\n')
            blk = _block(lines, '<%d년>' % year, ('<%d년>' % (year - 1),)) if year else lines
            blk = blk or lines
            pairs = [('해지율 가정 변경', 146, 153), ('위험률 가정 변경', 147, 154),
                     ('사업비율 가정 변경', 148, 155), ('기타 가정 변경', 149, 156)]
            # 표는 [최선추정부채, 위험조정, 보험계약마진] 3열.
            # DATA 의 'BEL+RA기준'은 앞 두 열의 합, 'CSM기준'은 마지막 열이다.
            # 값이 0 인 열은 PDF 에서 아예 빠지기도 해 숫자가 2개만 잡힌다.
            for lab, cbel, ccsm in pairs:
                v = _row_after(blk, 0, lab)
                if len(v) >= 2:
                    put(cbel, sum(v[:-1]), f'p{pno} {lab}')
                    put(ccsm, v[-1], f'p{pno} {lab}')
            # '물량차이 및 / 투자요소예실차 등' 은 두 줄로 쪼개져 숫자가 가운데 줄에 온다
            for i, ln in enumerate(blk):
                if ln.strip().startswith('물량차이'):
                    for k in range(i, min(i + 3, len(blk))):
                        v = _nums(blk[k])
                        if len(v) >= 3:
                            put(150, sum(v[:-1]), f'p{pno} 물량차이')
                            put(157, v[-1], f'p{pno} 물량차이')
                            break
                    break
            break

    # ---- ③ 보험금 예실차비율 → 예상손해율(101)·실제손해율(102)·예실차비율(103)
    for pno, txt in pages:
        if '보험금 예실차비율' in txt and '예상손해율' in txt:
            lines = txt.split('\n')
            for ln in lines:
                v = _nums(ln)
                if ln.strip().startswith(str(year or '')) and len(v) >= 4:
                    out[101], out[102] = v[1] / 100, v[2] / 100
                    out[103] = v[3] / 100
                    src[101] = src[102] = src[103] = f'p{pno} 예실차비율'
                    break
            break

    # 지급여력비율 총괄(경과조치 전/후)은 경영공시가 잠정치·반올림이라
    # DATA(FISIS 최종치)와 최대 1%p 어긋나 쓰지 않는다.

    # ---- ⑤ 일반회계와 감독회계의 차이 → 보험손익(48)·투자손익(49)·영업외손익(50)
    for pno, txt in pages:
        if '일반회계와 감독회계의 차이' in txt:
            lines = txt.split('\n')
            for lab, col in (('보험손익', 48), ('투자손익', 49), ('영업외손익', 50)):
                v = _row_after(lines, 0, lab)
                if v:
                    put(col, v[0], f'p{pno} 일반회계 {lab}')
            break

    # ---- ⑥ 해약환급금준비금 등의 적립 → 이익잉여금(18/71)·해약환급금준비금(99)
    for pno, txt in pages:
        if '해약환급금준비금 등의 적립' in txt:
            lines = txt.split('\n')
            for lab, cols in (('이익잉여금', (18, 71)), ('해약환급금준비금', (99,))):
                v = _row_after(lines, 0, lab)
                if v:
                    for c in cols:
                        put(c, v[0], f'p{pno} {lab}')
            break

    # ---- ⑦ 보험수익 주석 → CSM 상각(77)
    # DATA 의 CSM상각은 '(3) 당기손익에 인식한 보험계약마진 금액' 이다.
    # 그 아래 '가. CSM 상각액' 은 취소계약분이 빠져 있어(2024년 153억 차이) 쓰면 안 된다.
    for pno, txt in pages:
        if 'CSM 상각액' in txt:
            lines = txt.split('\n')
            v = _row_after(lines, 0, '(3) 당기손익에 인식한 보험계약마진')
            if not v:
                v = _row_after(lines, 0, '가. CSM 상각액')
            if v:
                out[77] = v[0]           # 이 주석만 백만원 단위
                src[77] = f'p{pno} 당기손익 인식 CSM'
            break

    return out, src


if __name__ == '__main__':
    import argparse, sys as _s
    _s.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf'); ap.add_argument('--year', type=int, default=2025)
    ap.add_argument('--check', default='', help='정답 대조용 회사약칭')
    ap.add_argument('--period', default='2512')
    a = ap.parse_args()
    vals, src = parse_disclosure(a.pdf, a.year)
    ans = {}
    if a.check:
        from ir_discover import answer_row
        ans, _ = answer_row(a.period, a.check)
    for c in sorted(vals):
        av = ans.get(c)
        mark = ''
        if av is not None:
            d = abs(vals[c] - av) / max(abs(av), 1e-9)
            mark = f'  정답={av:,.3f}  {"OK" if d < 0.002 else "FAIL(%.3f%%)" % (d*100)}'
        print(f'Col{c:<4} {vals[c]:>18,.3f}  [{src[c]}]{mark}')
