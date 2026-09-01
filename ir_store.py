# -*- coding: utf-8 -*-
"""받아온 IR 자료를 data/{회사명}/IR/{기간}/ 로 정리한다.

파일명이 회사마다 제각각이라(현대해상·메리츠는 해시 이름) **파일 내용을 보고**
회사와 기간을 판별한다. 시트 구성이 회사마다 뚜렷이 다르다는 점을 이용한다.

사용:
  python ir_store.py                      # 기본 소스 폴더에서 정리
  python ir_store.py --from ~/Downloads   # 소스 지정
  python ir_store.py --dry                # 어디로 갈지만 보여주기
"""
import os, re, shutil, argparse, sys
sys.stdout.reconfigure(encoding='utf-8')

from ir_parse import load, period_columns

# 시트 구성으로 회사를 판별한다. 값은 "이 시트들이 다 있으면 그 회사" 라는 뜻.
SHEET_SIGNATURES = {
    '미래에셋생명': {'INDEX', 'PROFIT', 'CSM', 'SAP-BS'},
    '삼성생명':     {'Ⅰ-2', 'Ⅰ-3', 'Ⅰ-5'},
    '한화생명':     {'보험부채 Movement', '신계약 CSM 및 수익성'},
    '삼성화재':     {'Profit & Loss Breakdown', 'Liability structure'},
    '동양생명':     {'FH', 'FS', 'CSM', 'Stability'},
    '현대해상':     {'cover_Q', '(3) UW Income'},
    '메리츠화재':   {'Group highlight', 'Insurance_Efficiency', 'CSM'},
    'KB금융지주':   {'G_BS', 'I_Key', 'L_Key'},
    '신한지주':     {'Shinhan Life'},
}

# PDF 는 파일명/본문 키워드로 판별한다.
# 폴더명은 ir_fill.SHORT 의 키와 같아야 한다 (KB라이프 → 'KB생명', 신한라이프 → '신한라이프생명').
PDF_HINTS = [
    ('교보생명',   ['교보생명']),
    ('DB손해보험', ['DB Insurance', 'DB손해보험', '경영실적 및']),
    ('KB손해보험', ['KB손해보험', 'KB손보']),
    ('KB생명',     ['KB라이프', 'KB생명']),
    ('신한라이프생명', ['신한라이프', '신한생명']),
    ('KB금융지주', ['PT_KOR', 'Factbook', 'KB금융']),
    ('삼성생명',   ['SLI ']),
    ('삼성화재',   ['SFMI', '삼성화재']),
    ('한화생명',   ['HLI ']),
    ('동양생명',   ['Tongyang', 'TYL']),
    ('미래에셋생명', ['Q4+Results', 'Q2+Results', 'Mirae']),
]

QUARTER_MONTH = {'Q1': '03', 'Q2': '06', 'Q3': '09', 'Q4': '12'}


def identify_xlsx(path):
    """시트 구성으로 회사를 판별한다. 기간은 보지 않는다.

    기간을 본문에서 추정하려 해봤으나, 시트 안의 숫자를 헤더로 오인해
    엉뚱한 값(7609, 8606 …)이 나오는 일이 잦았다. 기간은 받을 때 아는 값이므로
    --period 인자나 파일명에서 받는 편이 정확하다.
    """
    try:
        wb = load(path)
    except Exception:
        return None, None
    names = set(wb.sheetnames)
    wb.close()
    for c, sig in SHEET_SIGNATURES.items():
        if sig <= names:
            return c
    return None


def identify_pdf(path):
    base = os.path.basename(path)
    for comp, keys in PDF_HINTS:
        if any(k.lower() in base.lower() for k in keys):
            return comp
    # 신한라이프 경영공시는 파일명이 '2025년 년결산.pdf' 처럼 회사명이 없다.
    # 경영공시로 보이는 이름이면 첫 장 본문에서 회사명을 찾는다.
    if re.search(r'현황|결산|경영공시', base):
        try:
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                head = ' '.join((p.extract_text() or '') for p in pdf.pages[:2])
        except Exception:
            return None
        for comp, keys in PDF_HINTS:
            if any(k in head for k in keys):
                return comp
    return None


def period_from_name(name):
    """파일명에서 기간코드를 추정한다. 못 찾으면 None."""
    pats = [
        # 한국어 분기/반기/결산 패턴 (경영공시 PDF 파일명용)
        (r'(20\d\d)\D{0,3}([1-4])분기',  lambda m: (int(m.group(1)), int(m.group(2)))),
        (r'(20\d\d)\D{0,3}(?:상반기|반기)', lambda m: (int(m.group(1)), 2)),
        (r'(20\d\d)\D{0,3}(?:년?결산)',  lambda m: (int(m.group(1)), 4)),
        # 영문 패턴
        (r'(20\d\d)[^\d]{0,3}([1-4])Q', lambda m: (int(m.group(1)), int(m.group(2)))),
        (r'(\d\d)\.([1-4])Q',           lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
        (r'(20\d\d)\D{0,3}Q([1-4])',    lambda m: (int(m.group(1)), int(m.group(2)))),
        (r'(\d\d)_?([1-4])Q',           lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
        (r'([1-4])Q(\d\d)',             lambda m: (2000 + int(m.group(2)), int(m.group(1)))),
        (r'1H\s*FY(\d\d)',             lambda m: (2000 + int(m.group(1)), 2)),
        (r'(20\d\d)[._-]?1H',           lambda m: (int(m.group(1)), 2)),
        (r'1H(\d\d)',                   lambda m: (2000 + int(m.group(1)), 2)),
        (r'FY(20\d\d)[._-]?1H',         lambda m: (int(m.group(1)), 2)),
        (r'FY(20\d\d)',                 lambda m: (int(m.group(1)), 4)),
        (r'(20\d\d)년',                 lambda m: (int(m.group(1)), 4)),
        (r'FY(\d\d)(?!\d)',            lambda m: (2000 + int(m.group(1)), 4)),
    ]
    for pat, conv in pats:
        m = re.search(pat, name, re.I)
        if m:
            y, q = conv(m)
            return f'{y % 100:02d}{QUARTER_MONTH["Q%d" % q]}'
    return None


KIND = {'.xlsx': '팩트시트', '.pdf': '발표자료'}


def kind_of(comp, ext, name):
    if ext == '.pdf':
        if '공시' in name or re.search(r'현황|결산', name):
            return '경영공시'
        if 'factbook' in name.lower():
            return '팩트북'
        return '발표자료'
    if comp in ('KB금융지주', '신한지주') or 'factbook' in name.lower():
        return '팩트북'
    return '팩트시트'


def plan(src, period=None):
    """[(원본, 대상, 회사, 기간)]

    저장할 때 `{회사}_{기간}_{종류}.{확장자}` 로 이름을 바꾼다.
    현대해상·메리츠는 원래 파일명이 해시라 그대로 두면 나중에 알아볼 수 없다.
    """
    out, unknown = [], []
    for name in sorted(os.listdir(src)):
        path = os.path.join(src, name)
        if not os.path.isfile(path):
            continue
        ext = os.path.splitext(name)[1].lower()
        if ext == '.xlsx':
            comp = identify_xlsx(path)
        elif ext == '.pdf':
            comp = identify_pdf(path)
        else:
            continue
        if not comp:
            continue
        pd = period or period_from_name(name)
        if not pd:
            unknown.append((path, comp))
            continue
        kind = kind_of(comp, ext, name)
        newname = f'{comp}_{pd}_{kind}{ext}'
        out.append((path, os.path.join('data', comp, 'IR', pd, newname), comp, pd))
    return out, unknown


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--from', dest='src', default=None)
    ap.add_argument('--period', default=None,
                    help='이 배치의 기간코드(예: 2606). 주면 파일명 추정보다 우선한다')
    ap.add_argument('--dry', action='store_true')
    ap.add_argument('--copy', action='store_true', help='원본을 남긴다(기본은 이동)')
    a = ap.parse_args()

    from ir_fill import DL
    src = a.src or DL
    src = os.path.expanduser(src)
    if not os.path.isdir(src):
        print(f'소스 폴더가 없습니다: {src}')
        return

    items, unknown = plan(src, a.period)
    if not items:
        print(f'{src} 에서 IR 자료를 찾지 못했습니다.')
        return

    print(f'소스: {src}')
    for _s, dst, comp, period in items:
        print(f'  {comp:<10} {period:<6} → {dst}')
    for _s, comp in unknown:
        print(f'  {comp:<10} ?      → 기간을 알 수 없습니다. --period 로 지정하세요: '
              f'{os.path.basename(_s)}')
    if a.dry:
        print('\n(--dry 라 실제로 옮기지 않았습니다)')
        return

    moved = 0
    for s, dst, _c, _p in items:
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        if os.path.exists(dst) and os.path.getsize(dst) == os.path.getsize(s):
            continue                      # 같은 파일이 이미 있다
        (shutil.copy2 if a.copy else shutil.move)(s, dst)
        moved += 1
    print(f'\n{moved}개 파일 정리 완료 ({"복사" if a.copy else "이동"})')


if __name__ == '__main__':
    main()
