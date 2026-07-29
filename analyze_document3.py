# -*- coding: utf-8 -*-
"""
미래에셋생명 사업보고서 document.xml 심층 분석 - Part 3
위험보험료/사고보험금/예실차 테이블 및 CSM 기간별 테이블 집중 탐색
"""
import sys
import os
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8')

from dart_api import _fetch_document_content

rcept_no = "20250814003144"
content = _fetch_document_content(rcept_no)

def extract_tables_near(content, positions, context=3000, label=""):
    """positions 리스트의 각 위치 전후 context 내 TABLE 추출"""
    for pos in positions:
        # 해당 위치가 속한 TABLE 찾기
        table_start = content.rfind("<TABLE", 0, pos)
        # TABLE이 너무 멀면 skip
        if pos - table_start > 5000:
            continue
        table_end = content.find("</TABLE>", pos)
        if table_end < 0:
            continue
        chunk = content[table_start:table_end + 8]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        if len(tr_matches) < 2:
            continue
        print(f"\n  [{label}] 위치 {pos}: TABLE 크기={len(chunk)}, TR={len(tr_matches)}")
        for i, tr in enumerate(tr_matches[:25]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
            if clean:
                print(f"    TR[{i:2d}]: {clean}")
        return  # 첫 번째 유효한 것만

# =========================================================
# 1. 위험보험료 관련 실제 위치 탐색 - 보험수리지표 테이블
# =========================================================
print("=" * 70)
print("1. 위험보험료/사고보험금 보험수리지표 테이블 탐색")
print("=" * 70)

# 위치 1681111, 1683700, 1686219 주변
positions_risk = [m.start() for m in re.finditer("위험보험료", content)]
print(f"위험보험료 전체 위치: {positions_risk}")
for pos in positions_risk:
    table_start = content.rfind("<TABLE", 0, pos)
    table_end = content.find("</TABLE>", pos)
    if table_start < 0 or table_end < 0:
        continue
    if pos - table_start > 8000:
        continue
    chunk = content[table_start:table_end + 8]
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
    if len(tr_matches) < 3:
        continue
    print(f"\n위치 {pos}: TABLE TR={len(tr_matches)}")
    for i, tr in enumerate(tr_matches[:25]):
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
        if clean:
            print(f"  TR[{i:2d}]: {clean}")
    print()

# =========================================================
# 2. 보험수리 관련 위치 770608, 771310 탐색
# =========================================================
print("=" * 70)
print("2. 보험수리 관련 위치(770608, 771310) 탐색")
print("=" * 70)

for pos in [770608, 771310, 1812148, 2525915]:
    if pos >= len(content):
        continue
    # 주변 텍스트 확인
    snippet = content[max(0,pos-200):pos+500]
    # 텍스트만 추출
    clean_snippet = re.sub(r"<[^>]+>", "", snippet).strip()
    print(f"\n위치 {pos} 주변:")
    print(f"  {clean_snippet[:300]}")

    # TABLE 탐색
    table_start = content.rfind("<TABLE", 0, pos)
    table_end = content.find("</TABLE>", pos)
    if table_start >= 0 and table_end >= 0 and (pos - table_start) < 8000:
        chunk = content[table_start:table_end + 8]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"  TABLE: TR={len(tr_matches)}")
        for i, tr in enumerate(tr_matches[:20]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
            if clean:
                print(f"    TR[{i:2d}]: {clean}")

# =========================================================
# 3. CSM 기대수익인식 - 더 다양한 키워드 탐색
# =========================================================
print("\n" + "=" * 70)
print("3. CSM 기대수익인식/잔여보장기간 - 더 다양한 키워드")
print("=" * 70)

csm_kws = [
    "기대수익인식", "잔여보장기간", "잔여 보장기간", "기대 수익인식",
    "보장기간", "expected recognition", "remaining coverage",
    "미래에 인식될", "미래 인식", "향후 인식", "인식 예정"
]
for kw in csm_kws:
    idx = content.lower().find(kw.lower())
    if idx >= 0:
        snippet = re.sub(r"<[^>]+>", "", content[max(0,idx-100):idx+300]).strip()
        print(f"키워드 '{kw}' 위치 {idx}: ...{snippet[:200]}...")

# =========================================================
# 4. 손해율 테이블 실제 위치 탐색
# =========================================================
print("\n" + "=" * 70)
print("4. 손해율 테이블 상세 탐색")
print("=" * 70)

loss_ratio_positions = [m.start() for m in re.finditer("손해율", content)]
print(f"손해율 전체 위치: {loss_ratio_positions[:10]}")

for pos in loss_ratio_positions:
    table_start = content.rfind("<TABLE", 0, pos)
    table_end = content.find("</TABLE>", pos)
    if table_start < 0 or table_end < 0:
        continue
    if pos - table_start > 10000:
        continue
    chunk = content[table_start:table_end + 8]
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
    if len(tr_matches) < 3:
        continue
    # 실제 숫자가 있는 TR인지 확인
    all_cells = []
    for tr in tr_matches:
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        all_cells.extend([re.sub(r"<[^>]+>", "", c).strip() for c in cells])
    nums = [c for c in all_cells if re.match(r"[\d,\.]+$", c)]
    if len(nums) < 5:
        continue
    print(f"\n손해율 위치 {pos}: TABLE TR={len(tr_matches)}, 숫자셀={len(nums)}")
    for i, tr in enumerate(tr_matches[:20]):
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
        if clean:
            print(f"  TR[{i:2d}]: {clean}")

# =========================================================
# 5. 예실차 테이블 상세 (모든 위치)
# =========================================================
print("\n" + "=" * 70)
print("5. 예실차 테이블 모든 위치 탐색")
print("=" * 70)

yesil_positions = [m.start() for m in re.finditer("예실차", content)]
print(f"예실차 전체 위치: {yesil_positions}")

for pos in yesil_positions:
    table_start = content.rfind("<TABLE", 0, pos)
    table_end = content.find("</TABLE>", pos)
    if table_start < 0 or table_end < 0:
        continue
    if pos - table_start > 10000:
        continue
    chunk = content[table_start:table_end + 8]
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
    # 숫자 있는 테이블인지
    all_nums = re.findall(r">\s*([\d,]{3,})\s*<", chunk)
    if len(all_nums) < 3:
        continue
    print(f"\n예실차 위치 {pos}: TABLE TR={len(tr_matches)}")
    for i, tr in enumerate(tr_matches[:25]):
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
        if clean:
            print(f"  TR[{i:2d}]: {clean}")

# =========================================================
# 6. K-ICS 가용자본 상세 항목 (가용자본 세부 분류 테이블)
# =========================================================
print("\n" + "=" * 70)
print("6. K-ICS 가용자본/요구자본 세부 테이블")
print("=" * 70)

garyong_positions = [m.start() for m in re.finditer("가용자본", content)]
print(f"가용자본 위치: {garyong_positions[:10]}")

for pos in garyong_positions:
    table_start = content.rfind("<TABLE", 0, pos)
    table_end = content.find("</TABLE>", pos)
    if table_start < 0 or table_end < 0:
        continue
    if pos - table_start > 5000:
        continue
    chunk = content[table_start:table_end + 8]
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
    all_nums = re.findall(r">\s*([\d,]{4,})\s*<", chunk)
    if len(all_nums) < 3:
        continue
    print(f"\n가용자본 위치 {pos}: TABLE TR={len(tr_matches)}, 숫자={len(all_nums)}")
    for i, tr in enumerate(tr_matches[:20]):
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
        if clean:
            print(f"  TR[{i:2d}]: {clean}")
    break

# =========================================================
# 7. 민감도 분석 전체 - 금리 포함 여부
# =========================================================
print("\n" + "=" * 70)
print("7. 금리 민감도 분석 탐색")
print("=" * 70)

interest_kws = ["금리충격", "금리 충격", "이자율충격", "이자율 충격", "+50bp", "+100bp", "금리상승", "금리하락"]
for kw in interest_kws:
    all_positions = [m.start() for m in re.finditer(re.escape(kw), content, re.IGNORECASE)]
    if all_positions:
        print(f"\n'{kw}' 발견 ({len(all_positions)}개): 위치 {all_positions[:5]}")
        for pos in all_positions[:2]:
            table_start = content.rfind("<TABLE", 0, pos)
            table_end = content.find("</TABLE>", pos)
            if table_start >= 0 and table_end >= 0 and (pos - table_start) < 8000:
                chunk = content[table_start:table_end + 8]
                tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
                print(f"  TABLE TR={len(tr_matches)}")
                for i, tr in enumerate(tr_matches[:15]):
                    cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
                    clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                    clean = [c.replace("\n","").replace("\r","").strip() for c in clean if c.strip() and c.strip() not in ["&nbsp;", "\xa0"]]
                    if clean:
                        print(f"    TR[{i:2d}]: {clean}")
                break

# =========================================================
# 8. 사업보고서의 Data Sheet 관련 항목 전체 요약
# =========================================================
print("\n" + "=" * 70)
print("8. 모든 테이블 헤더 수집 (숫자가 많은 테이블 우선)")
print("=" * 70)

# 모든 TABLE 파싱해서 헤더가 의미 있는 것 추출
all_tables = re.findall(r"<TABLE[^>]*>(.*?)</TABLE>", content, re.DOTALL)
print(f"전체 TABLE 개수: {len(all_tables)}")

useful_tables = []
for i, tbl in enumerate(all_tables):
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", tbl, re.DOTALL)
    if len(tr_matches) < 3:
        continue
    all_nums = re.findall(r">\s*([\d,]{4,})\s*<", tbl)
    if len(all_nums) < 5:
        continue
    # 헤더 추출
    first_tr = tr_matches[0]
    cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", first_tr, re.DOTALL)
    header = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
    header = [c for c in header if c and c not in ["&nbsp;", "\xa0"]]
    useful_tables.append((i, len(tr_matches), len(all_nums), header))

print(f"유의미한 TABLE (TR>=3, 숫자>=5): {len(useful_tables)}개\n")
for idx, n_tr, n_num, header in useful_tables:
    print(f"TABLE[{idx:3d}] TR={n_tr:3d} 숫자={n_num:4d}: {header[:8]}")

print("\n완료")
