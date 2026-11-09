# -*- coding: utf-8 -*-
"""IR/경영공시 PDF 에서 DATA 시트 컬럼을 추출한다.

연간 경영공시의 공통 표를 읽는 규칙 추출 모듈이다.
분기별 판단과 최종 반영은 quarterly 파이프라인의 PDF 직접 검토에서 수행한다.

PDF 텍스트에는 숫자 중간에 공백이 끼어 나오는 경우가 많다("1 42,185" = 142,185).
`_nums()`가 열 사이 공백과 숫자 안의 공백을 구분한다.
"""
import re
import pdfplumber


def num(tok):
    t = str(tok).replace(',', '').replace('△', '-').replace('▲', '-').strip()
    if t.startswith('(') and t.endswith(')'):
        t = '-' + t[1:-1].strip()
    try:
        return float(t)
    except ValueError:
        return None


_TOKEN = re.compile(r'\(\s*-?[\d,]+(?:\.\d+)?\s*\)|-?[\d,]+(?:\.\d+)?')


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


# ---------------------------------------------------------------- 교보생명
# CSM 상각 만기 버킷 → DATA 컬럼 (표는 1년,2년,…,10년,11~15,16~20,21~25,26~30,30년이후,계)


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
        match = re.match(''.join(re.escape(c) + r'\s*' for c in label if not c.isspace()), s)
        if match:
            v = _nums(s[match.end():])
            return v[:n] if n else v
    return []


def _explicit_year_block(text, year):
    """Require an explicit table-year heading; never silently read a prior year."""
    lines = text.splitlines()
    heading = re.compile(r'^\s*<?\s*(\d{4})\s*년(?:도)?\s*>?\s*$')
    start = None
    for i, line in enumerate(lines):
        match = heading.match(line)
        if match:
            if start is not None:
                return lines[start:i]
            if int(match[1]) == year:
                start = i + 1
    return lines[start:] if start is not None else []


def disclosure_csm_stocks(text, year):
    """CSM only: BEL/RA here are LRC-only and cannot fill full net liabilities."""
    compact = re.sub(r'\s+', '', text)
    if not all(t in compact for t in ('회계모형별', '보험부채현황', '억원')):
        return {}
    result = {}
    for table_year, field in ((year, 75), (year-1, 80)):
        block = _explicit_year_block(text, table_year)
        if not block:
            continue
        values = _row_after(block, 0, '합계')
        has_vfa = '변동수수료' in re.sub(r'\s+', '', '\n'.join(block))
        # GMM: BEL, RA, CSM, [PAA]. GMM+VFA: two groups of three, [PAA].
        if has_vfa and len(values) in (6, 7):
            result[field] = (values[2] + values[5]) * DISC_UNIT
        elif not has_vfa and len(values) in (3, 4):
            result[field] = values[2] * DISC_UNIT
    return result


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
        pages = []
        for i, p in enumerate(pdf.pages):
            try:
                pages.append((i + 1, p.extract_text() or ''))
            except Exception:
                pages.append((i + 1, ''))

    # ---- ① Explicit current/prior year, GMM and optional VFA CSM only.
    for index, (pno, txt) in enumerate(pages):
        compact = re.sub(r'\s+', '', txt)
        if year and '회계모형별' in compact and '보험부채현황' in compact:
            # Totals and the prior-year table often continue on the next page.
            table_text = '\n'.join(t for _, t in pages[index:index+3])
            for field, value in disclosure_csm_stocks(table_text, year).items():
                out[field], src[field] = value, f'p{pno} 명시 연도 보험부채현황 CSM 합계'

    # ---- ② 계리적 가정 → 가정변경(BEL기준 Col146~150 / CSM기준 Col153~157)
    for pno, txt in pages:
        if '계리적 가정' in txt and '가정변경효과' in txt:
            lines = txt.split('\n')
            blk = _explicit_year_block(txt, year) if year else []
            if not blk:
                continue
            new_csm = _row_after(blk, 0, '신계약')
            if len(new_csm) == 3:
                put(76, new_csm[2], f'p{pno} 계리적 가정 신계약 보험계약마진; 원수·수재 범위 확인 필요')
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

    # K-ICS의 잠정/최종과 경과조치 전후는 PDF 직접 검토에서 확인한다.

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
