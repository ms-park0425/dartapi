"""
미래에셋생명 사업보고서 document.xml 분석 스크립트
- rcept_no 찾기 (corp_code=00112332, year=2025, reprt_code=11011)
- document.xml 내 주요 항목 존재 여부 확인
"""
import sys
import os
import re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dart_api import find_rcept_no, _fetch_document_content

# 1. rcept_no 찾기
print("=" * 60)
print("1. 미래에셋생명 사업보고서(2025) rcept_no 조회")
print("=" * 60)
corp_code = "00112332"
try:
    rcept_no = find_rcept_no(corp_code, "2025", "11011")
    print(f"rcept_no: {rcept_no}")
except Exception as e:
    print(f"오류: {e}")
    sys.exit(1)

# 2. document.xml 가져오기
print("\n" + "=" * 60)
print("2. document.xml 다운로드")
print("=" * 60)
content = _fetch_document_content(rcept_no)
print(f"문서 크기: {len(content):,} bytes")

# 3. 키워드 검색 함수
def search_keyword(label, keywords, content, context=200):
    """키워드 목록 중 하나라도 찾으면 주변 컨텍스트 반환"""
    for kw in keywords:
        idx = content.find(kw)
        if idx >= 0:
            start = max(0, idx - 80)
            end = min(len(content), idx + context)
            snippet = content[start:end].replace("\n", " ").replace("\r", "")
            return True, kw, snippet
    return False, None, None

# 4. 각 항목 검색
print("\n" + "=" * 60)
print("3. 항목별 키워드 검색")
print("=" * 60)

checks = [
    # (레이블, 검색 키워드 목록)
    ("K-ICS 지급여력비율", ["지급여력비율", "킥스비율", "K-ICS비율", "KICS비율", "지급여력비율(A/B)"]),
    ("가용자본(A)", ["가용자본", "지급여력금액(A)", "지급여력금액 (A)"]),
    ("요구자본(B)", ["요구자본", "지급여력기준(B)", "지급여력기준 (B)"]),
    ("CSM 기대수익인식(잔여보장기간별)", ["기대수익인식", "잔여보장기간", "잔여 보장기간"]),
    ("1년이하 기대수익인식", ["1년 이하", "1년이하"]),
    ("1~3년 기대수익인식", ["1년 초과 3년", "1년초과 3년", "1~3년"]),
    ("3~5년 기대수익인식", ["3년 초과 5년", "3년초과 5년", "3~5년"]),
    ("5~10년 기대수익인식", ["5년 초과 10년", "5년초과 10년", "5~10년"]),
    ("10년 이상 기대수익인식", ["10년 초과", "10년초과", "10년 이상"]),
    ("누적해약환급금준비금", ["해약환급금준비금", "누적 해약환급금"]),
    ("손해율", ["손해율"]),
    ("위험보험료", ["위험보험료"]),
    ("사고보험금", ["사고보험금"]),
    ("예실차", ["예실차"]),
    ("보험수리지표", ["보험수리지표", "보험수리"]),
    ("BEL 민감도", ["BEL", "최선추정부채"]),
    ("RA 민감도", ["위험조정", "Risk Adjustment"]),
    ("금리 민감도", ["금리충격", "금리 민감도", "이자율충격"]),
    ("CSM 변동", ["CSM", "계약서비스마진"]),
    ("신계약 CSM", ["최초 인식", "신계약CSM", "신계약 CSM"]),
]

results = []
for label, keywords in checks:
    found, matched_kw, snippet = search_keyword(label, keywords, content)
    results.append((label, found, matched_kw, snippet))

# 결과 출력
print(f"\n{'항목':<35} {'발견':^5} {'매칭키워드'}")
print("-" * 80)
for label, found, matched_kw, snippet in results:
    status = "O" if found else "X"
    kw_info = matched_kw if matched_kw else "-"
    print(f"{label:<35} [{status}]   {kw_info}")

# 5. 발견된 항목 상세 스니펫
print("\n" + "=" * 60)
print("4. 발견된 항목 상세 컨텍스트")
print("=" * 60)
for label, found, matched_kw, snippet in results:
    if found:
        print(f"\n[{label}] (keyword: {matched_kw})")
        print(f"  ...{snippet}...")

# 6. K-ICS 관련 테이블 전체 구조 파악
print("\n" + "=" * 60)
print("5. K-ICS 테이블 구조 (지급여력비율 전후 1000자)")
print("=" * 60)
for kw in ["지급여력비율(A/B)", "지급여력비율", "킥스비율"]:
    idx = content.find(kw)
    if idx >= 0:
        start = max(0, idx - 300)
        end = min(len(content), idx + 1500)
        chunk = content[start:end]
        # 숫자 추출
        nums = re.findall(r">\s*([\d,\.]+)\s*</T[DH]>", chunk)
        print(f"\n키워드 '{kw}' 발견, 주변 숫자들: {nums[:20]}")
        # TR 단위 파싱
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        for i, tr in enumerate(tr_matches[:15]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c]
            if clean:
                print(f"  TR[{i}]: {clean}")
        break

# 7. CSM 기간별 테이블 구조
print("\n" + "=" * 60)
print("6. CSM 기대수익인식 테이블 구조")
print("=" * 60)
for kw in ["기대수익인식", "잔여보장기간"]:
    idx = content.find(kw)
    if idx >= 0:
        start = max(0, idx - 300)
        end = min(len(content), idx + 2000)
        chunk = content[start:end]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"\n키워드 '{kw}' 발견, TR 수: {len(tr_matches)}")
        for i, tr in enumerate(tr_matches[:20]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c]
            if clean:
                print(f"  TR[{i}]: {clean}")
        break

# 8. 손해율/예실차 테이블 구조
print("\n" + "=" * 60)
print("7. 보험수리지표 (손해율/예실차/위험보험료) 테이블 구조")
print("=" * 60)
for kw in ["예실차", "손해율", "위험보험료"]:
    idx = content.find(kw)
    if idx >= 0:
        start = max(0, idx - 300)
        end = min(len(content), idx + 2000)
        chunk = content[start:end]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"\n키워드 '{kw}' 발견, TR 수: {len(tr_matches)}")
        for i, tr in enumerate(tr_matches[:20]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c]
            if clean:
                print(f"  TR[{i}]: {clean}")
        break

# 9. BEL/RA 민감도 분석
print("\n" + "=" * 60)
print("8. BEL/RA 민감도 분석 구조")
print("=" * 60)
for kw in ["민감도", "금리충격", "BEL"]:
    idx = content.find(kw)
    if idx >= 0:
        start = max(0, idx - 200)
        end = min(len(content), idx + 2000)
        chunk = content[start:end]
        tr_matches = re.findall(r"<TR[^>]*>(.*?)</TR>", chunk, re.DOTALL)
        print(f"\n키워드 '{kw}' 발견, TR 수: {len(tr_matches)}")
        for i, tr in enumerate(tr_matches[:20]):
            cells = re.findall(r"<T[DH][^>]*>(.*?)</T[DH]>", tr, re.DOTALL)
            clean = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
            clean = [c for c in clean if c]
            if clean:
                print(f"  TR[{i}]: {clean}")
        break

# 10. TE 태그(XBRL 인라인) 기반 검색
print("\n" + "=" * 60)
print("9. TE 태그(인라인 XBRL) 내 주요 ACODE 목록 (상위 50개)")
print("=" * 60)
acodes = re.findall(r'ACODE="([^"]+)"', content)
from collections import Counter
acode_counts = Counter(acodes)
print(f"총 고유 ACODE 수: {len(acode_counts)}")
# 보험/CSM/KICS 관련만 필터
relevant = [(k, v) for k, v in acode_counts.most_common(200)
            if any(x in k.lower() for x in ["csm", "kics", "solvency", "contractual", "expected", "loss", "risk", "claim", "premium", "sensitivity"])]
print(f"\n관련 ACODE (csm/kics/solvency/contractual/loss/risk/claim/premium/sensitivity 포함):")
for k, v in relevant[:50]:
    print(f"  [{v:3}] {k}")

print("\n완료")
