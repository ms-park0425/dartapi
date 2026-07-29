"""11개 회사 2025 사업보고서 XBRL 다운로드 스크립트"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dart_api import find_corp_code, get_disclosures, download_xbrl

COMPANIES = [
    "삼성생명",
    "교보생명",
    "한화생명",
    "신한라이프생명",
    "동양생명",
    "KB생명",
    "삼성화재",
    "현대해상",
    "메리츠화재",
    "KB손해보험",
    "DB손해보험",
]


def find_report(items):
    """사업보고서 (2025.12) 중 가장 최근 것 선택. 기재정정 우선."""
    candidates = []
    for item in items:
        rn = item.get("report_nm", "")
        if "사업보고서" not in rn:
            continue
        if "2025.12" not in rn:
            continue
        candidates.append(item)

    if not candidates:
        return None

    # 기재정정 버전이 있으면 우선
    corrected = [c for c in candidates if "기재정정" in c.get("report_nm", "")]
    if corrected:
        return max(corrected, key=lambda x: x.get("rcept_dt", ""))
    return max(candidates, key=lambda x: x.get("rcept_dt", ""))


def main():
    success = []
    failed = []

    for company in COMPANIES:
        print(f"\n{'='*60}")
        print(f"처리 중: {company}")
        print(f"{'='*60}")

        try:
            # 1. corp_code 조회
            print(f"  1) corp_code 조회...")
            corp_code = find_corp_code(company)
            print(f"     corp_code: {corp_code}")

            # 2. 공시 목록 조회
            print(f"  2) 공시 목록 조회...")
            data = get_disclosures(
                corp_code=corp_code,
                bgn_de='20260101',
                end_de='20260401',
                page_count=100
            )
            items = data.get("list", [])
            print(f"     공시 {len(items)}건 조회됨")

            # 3. 사업보고서 선택
            report = find_report(items)
            if report is None:
                print(f"  [ERROR] 사업보고서 (2025.12) 를 찾을 수 없습니다.")
                # 후보 출력
                for item in items[:5]:
                    print(f"    - {item.get('report_nm', '')} ({item.get('rcept_dt', '')})")
                failed.append((company, "사업보고서 없음"))
                continue

            rcept_no = report["rcept_no"]
            print(f"  3) 선택: {report['report_nm']} (접수일: {report['rcept_dt']}, rcept_no: {rcept_no})")

            # 4. XBRL 다운로드
            print(f"  4) XBRL 다운로드...")
            save_path = download_xbrl(
                rcept_no=rcept_no,
                corp_name=company,
                reprt_name='사업보고서',
                year='2025',
                reprt_code='11011'
            )
            success.append(company)

        except Exception as e:
            print(f"  [ERROR] {company}: {e}")
            failed.append((company, str(e)))
            continue

    # 결과 요약
    print(f"\n{'='*60}")
    print(f"완료 요약")
    print(f"{'='*60}")
    print(f"성공: {len(success)}개 - {success}")
    if failed:
        print(f"실패: {len(failed)}개")
        for name, reason in failed:
            print(f"  - {name}: {reason}")


if __name__ == "__main__":
    main()
