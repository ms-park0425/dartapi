# -*- coding: utf-8 -*-
"""
미래에셋생명 사업보고서 document.xml 심층 분석 - Part 2
"""
import sys
import os
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding='utf-8')

from dart_api import _fetch_document_content

rcept_no = "20250814003144"
content = _fetch_document_content(rcept_no)
print(f"[문서 크기: {len(content):,} bytes]\n")

# =========================================================
# 1. K-ICS 테이블 전체 내용
# =========================================================
print("=" * 70)
print("1. K-ICS 지급여력비율 테이블 (지급여력비율(A/B) 전후 2000자)")
print("=" * 70)

idx = content.find("지급여력비율(A/B)")
if idx >= 0:
    start = max(0, idx - 500)
    end = min(len(content), idx + 2000)
    chunk = content[start:end]
    tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
    print(f"TR 개수: {len(tr_matches)}")
    for i, tr in enumerate(tr_matches):
        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        clean = [c for c in clean if c and c not in ["&nbsp;", " "]]
        if clean:
            print(f"  TR[{i:2d}]: {clean}")

# =========================================================
# 2. CSM 기간별 테이블 - 더 넓게 탐색
# =========================================================
print("\n" + "=" * 70)
print("2. CSM 잔여보장기간별 기대수익인식 테이블")
print("=" * 70)

# 여러 키워드로 탐색
keywords_csm = ["잔여보장기간", "기대수익인식", "잔여 보장기간", "기대 수익인식",
                "ExpectedRecognition", "RemainingCoverage", "1년 이하", "1년이하"]
found_csm = False
for kw in keywords_csm:
    idx = content.find(kw)
    if idx >= 0:
        # 테이블 시작점 찾기
        table_start = content.rfind("<TABLE", 0, idx)
        table_end_search = content.find("</TABLE>", idx)
        if table_start >= 0 and table_end_search >= 0:
            chunk = content[table_start:table_end_search + 8]
            tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
            print(f"키워드 '{kw}' 발견 → TABLE {len(tr_matches)}개 TR")
            for i, tr in enumerate(tr_matches[:30]):
                cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
                clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                clean = [c for c in clean if c and c not in ["&nbsp;", " ", ""]]
                if clean:
                    print(f"    TR[{i:2d}]: {clean}")
            found_csm = True
            break

if not found_csm:
    print("CSM 기간별 테이블을 찾지 못했습니다.")
    # 1년이하 주변 탐색
    idx = content.find("1년 이하")
    if idx >= 0:
        start = max(0, idx - 300)
        end = min(len(content), idx + 3000)
        chunk = content[start:end]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"'1년 이하' 주변 TR {len(tr_matches)}개:")
        for i, tr in enumerate(tr_matches[:30]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c and c not in ["&nbsp;", " ", ""]]
            if clean:
                print(f"    TR[{i:2d}]: {clean}")

# =========================================================
# 3. 손해율/위험보험료/사고보험금/예실차 테이블 탐색
# =========================================================
print("\n" + "=" * 70)
print("3. 보험수리지표 테이블 (손해율, 위험보험료, 사고보험금, 예실차)")
print("=" * 70)

insurance_keywords = ["예실차", "위험보험료", "사고보험금", "손해율", "보험수리"]
for kw in insurance_keywords:
    all_idx = [m.start() for m in re.finditer(re.escape(kw), content)]
    if all_idx:
        print(f"\n키워드 '{kw}' - {len(all_idx)}개 위치: {all_idx[:5]}")
        # 각 위치에서 해당 테이블 구조 추출
        for pos in all_idx[:3]:
            table_start = content.rfind("<TABLE", 0, pos)
            table_end = content.find("</TABLE>", pos)
            if table_start >= 0 and table_end >= 0 and (pos - table_start) < 5000:
                chunk = content[table_start:table_end + 8]
                tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
                if len(tr_matches) >= 2:
                    print(f"  위치 {pos}: TABLE {len(tr_matches)}개 TR")
                    for i, tr in enumerate(tr_matches[:15]):
                        cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
                        clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                        clean = [c for c in clean if c and c not in ["&nbsp;", " ", ""]]
                        if clean:
                            print(f"      TR[{i:2d}]: {clean}")
                    break

# =========================================================
# 4. 민감도 분석 테이블
# =========================================================
print("\n" + "=" * 70)
print("4. BEL/RA 민감도 분석 테이블")
print("=" * 70)

sensitivity_keywords = ["민감도 분석", "민감도분석", "금리 충격", "금리충격", "이자율 충격", "+50bp", "+100bp"]
for kw in sensitivity_keywords:
    idx = content.find(kw)
    if idx >= 0:
        table_start = content.rfind("<TABLE", 0, idx)
        table_end = content.find("</TABLE>", idx)
        if table_start >= 0 and table_end >= 0:
            chunk = content[table_start:table_end + 8]
            tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
            print(f"키워드 '{kw}' 발견 → TABLE {len(tr_matches)}개 TR")
            for i, tr in enumerate(tr_matches[:20]):
                cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
                clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                clean = [c for c in clean if c and c not in ["&nbsp;", " ", ""]]
                if clean:
                    print(f"    TR[{i:2d}]: {clean}")
            break

# =========================================================
# 5. 해약환급금준비금 테이블
# =========================================================
print("\n" + "=" * 70)
print("5. 해약환급금준비금 상세 테이블")
print("=" * 70)

idx = content.find("해약환급금준비금 변동")
if idx < 0:
    idx = content.find("해약환급금준비금")
if idx >= 0:
    table_start = content.rfind("<TABLE", 0, idx)
    table_end = content.find("</TABLE>", idx)
    if table_start >= 0 and table_end >= 0:
        chunk = content[table_start:table_end + 8]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"해약환급금준비금 TABLE {len(tr_matches)}개 TR")
        for i, tr in enumerate(tr_matches[:20]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c and c not in ["&nbsp;", " ", ""]]
            if clean:
                print(f"    TR[{i:2d}]: {clean}")

# =========================================================
# 6. TE 태그에서 CSM/KICS 관련 ACODE 수집
# =========================================================
print("\n" + "=" * 70)
print("6. 인라인 XBRL(TE 태그) - CSM/KICS 관련 전체 데이터")
print("=" * 70)

# CSM 관련 TE 태그 추출
csm_pattern = re.compile(
    r'<TE[^>]*ACODE="([^"]*(?:ContractualService|CSM|ServiceMargin)[^"]*)"'
    r'[^>]*ACONTEXT="([^"]+)"[^>]*>([^<]*)</TE>',
    re.DOTALL
)
csm_entries = []
for m in csm_pattern.finditer(content):
    acode, ctx, val = m.group(1), m.group(2), m.group(3).strip()
    if val and val not in ["", "0"]:
        csm_entries.append((acode, ctx, val))

print(f"CSM 관련 TE 태그 수: {len(csm_entries)}")
for acode, ctx, val in csm_entries[:30]:
    print(f"  ACODE={acode}")
    print(f"    CTX={ctx}")
    print(f"    VAL={val}")

# KICS 관련
print("\nKICS/Solvency 관련 TE 태그:")
kics_pattern = re.compile(
    r'<TE[^>]*ACODE="([^"]*(?:[Ss]olvency|[Kk][Ii][Cc][Ss]|AvailableCapital|RequiredCapital)[^"]*)"'
    r'[^>]*ACONTEXT="([^"]+)"[^>]*>([^<]*)</TE>',
    re.DOTALL
)
kics_entries = []
for m in kics_pattern.finditer(content):
    acode, ctx, val = m.group(1), m.group(2), m.group(3).strip()
    if val:
        kics_entries.append((acode, ctx, val))

print(f"KICS 관련 TE 태그 수: {len(kics_entries)}")
for acode, ctx, val in kics_entries[:20]:
    print(f"  ACODE={acode}, VAL={val}")
    print(f"    CTX={ctx}")

# =========================================================
# 7. 1년이하 ~ 10년초과 구조 전체 파악 (잔여보장기간 테이블)
# =========================================================
print("\n" + "=" * 70)
print("7. 잔여보장기간별 테이블 전체 탐색 (1년이하/1~2년/2~3년... 형태)")
print("=" * 70)

# '1년 이하' ~ '합 계' 구간 전체 테이블 찾기
duration_patterns = ["1년 이하", "1년이하", "1 year or less"]
for kw in duration_patterns:
    idx = content.find(kw)
    if idx >= 0:
        # 테이블 시작
        table_start = content.rfind("<TABLE", 0, idx)
        # 테이블 끝 (더 넓게)
        table_end = content.find("</TABLE>", idx)
        if table_start >= 0 and table_end >= 0:
            chunk = content[table_start:table_end + 8]
            print(f"키워드: '{kw}', TABLE 크기: {len(chunk)}")
            tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
            print(f"TR 총 개수: {len(tr_matches)}")
            for i, tr in enumerate(tr_matches):
                cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
                clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
                clean = [c for c in clean if c and c.strip() not in ["&nbsp;", " ", "", "\xa0"]]
                if clean:
                    print(f"  TR[{i:2d}]: {clean}")
        break

# =========================================================
# 8. 위험보험료 TE 태그 탐색
# =========================================================
print("\n" + "=" * 70)
print("8. 위험보험료/사고보험금/손해율 TE 태그 탐색")
print("=" * 70)

risk_pattern = re.compile(
    r'<TE[^>]*ACODE="([^"]*(?:[Rr]isk[Pp]remium|[Cc]laim|[Ll]oss[Rr]atio|[Aa]ctual[Ee]xpected|[Pp]remium)[^"]*)"'
    r'[^>]*ACONTEXT="([^"]+)"[^>]*>([^<]*)</TE>',
    re.DOTALL
)
risk_entries = []
for m in risk_pattern.finditer(content):
    acode, ctx, val = m.group(1), m.group(2), m.group(3).strip()
    if val:
        risk_entries.append((acode, ctx, val))

print(f"위험보험료/손해율 관련 TE 태그 수: {len(risk_entries)}")
for acode, ctx, val in risk_entries[:20]:
    print(f"  ACODE={acode}, VAL={val}")

print("\n완료")
