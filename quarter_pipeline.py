"""One quarter selection -> official sources -> Excel -> direct PDF review -> final Excel."""
from __future__ import annotations

import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent


def read_config(path):
    from quarterly.core import load
    path=Path(path).resolve()
    config=load(path)
    if path == (ROOT/'quarterly.json').resolve() and (ROOT/'quarterly.local.json').exists():
        local=load(ROOT/'quarterly.local.json')
        if not isinstance(local,dict):raise ValueError('quarterly.local.json 설정은 객체여야 합니다.')
        config={**config,**local}
    return config


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="분기 선택 → API/PDF 수집 → 엑셀 입력 → PDF 직접 검토·수정")
    parser.add_argument("--period", help="예: 2609")
    parser.add_argument("--year", type=int)
    parser.add_argument("--quarter", type=int, choices=range(1, 5))
    parser.add_argument("--config", type=Path, default=ROOT / "quarterly.json")
    parser.add_argument("--workbook", help="초기 작업 엑셀 경로 재정의")
    parser.add_argument("--companies", nargs="+", help="일부 회사로 검증할 때 사용")
    parser.add_argument("--offline", action="store_true", help="저장된 자료만 사용")
    parser.add_argument("--check-baseline", action="store_true", help="2512backup 정답본과 과거 값을 읽기 전용 대조")
    parser.add_argument("--reviewer", choices=["codex", "assistant"], default="codex")
    parser.add_argument("--resume", type=Path, help="동일 실행의 PDF 검토 재개")
    parser.add_argument("--finalize", type=Path, help="현재 Codex 채팅에서 작성한 검토 결과 반영")
    args = parser.parse_args(argv)
    from dotenv import load_dotenv
    load_dotenv(ROOT / ".env")
    from quarterly.core import Period
    from quarterly.pipeline import finalize, prepare, review_run
    try:
        if args.finalize:
            finalize(args.finalize)
            return 0
        if args.resume:
            review_run(args.resume)
            return 0
        config = read_config(args.config)
        if not isinstance(config, dict):
            raise ValueError("quarterly.json 설정 파일이 필요합니다.")
        if args.workbook:
            config["workbook"] = args.workbook
        if args.check_baseline:
            from quarterly.history import history_report
            history_report(ROOT, config, period=args.period, company_filter=args.companies)
            return 0
        if args.reviewer == "assistant":
            # The current chat handles any official-site fallback without another AI process.
            config = {**config, "ai_source_discovery": False}
        if args.period:
            period = Period.parse(args.period)
        elif args.year and args.quarter:
            period = Period(args.year, args.quarter)
        else:
            parser.error("--period 2609 또는 --year 2026 --quarter 3을 지정하세요.")
        if args.reviewer == "codex":
            from quarterly.review import CodexReviewer
            reviewer = CodexReviewer(config.get("codex", "codex"))
        run = prepare(ROOT, period, config, offline=args.offline, company_filter=args.companies)
        print(f"실행 폴더: {run}")
        if args.reviewer == "assistant":
            print("상태: PDF 직접 검토 대기. 현재 Codex 채팅에서 PDF 검토·반영까지 이어서 수행해야 합니다.")
            return 0
        review_run(run, reviewer)
        return 0
    except Exception as exc:
        print(f"실행 중단: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
