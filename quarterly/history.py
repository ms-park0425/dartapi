"""Read-only comparison against the user-designated 2512backup answer workbook."""
from __future__ import annotations

import csv
from pathlib import Path

import openpyxl

from .core import digest, dump, korea_now, numeric
from .review import RATIOS, CHECKS
from .workbook import DataBook


def read_reference(path, through="2512"):
    layout = DataBook(path, protect_through=through)
    cache = openpyxl.load_workbook(path, data_only=True)["DATA"]
    values = {}
    uncached = []
    for (period, short), row in layout.rows.items():
        if not period.isdigit() or int(period) > int(through):
            continue
        for field, ec in layout.columns.items():
            if field <= 3 or field in CHECKS:
                continue
            v = cache.cell(row, ec).value
            if numeric(v):
                values[(period, short, field)] = v
            elif layout.ws.cell(row, ec).data_type == "f":
                uncached.append(dict(period=period, company=short, field=field))
    return values, uncached


def compare(value, expected, field):
    tolerance = max(0.0001 if field in RATIOS else 1, abs(expected) * 0.001)
    if not numeric(value):
        return "NO_VALUE", None, tolerance
    delta = value - expected
    return "OK" if abs(delta) <= tolerance else "DIFF", delta, tolerance


def history_report(root, config, period=None, company_filter=None, log=print):
    root = Path(root).resolve()
    reference = (root / config["reference_workbook"]).resolve()
    working = (root / config["workbook"]).resolve()
    ref_sha, work_sha = digest(reference), digest(working)
    through = config.get("protect_through", "2512")
    answer, uncached = read_reference(reference, through)
    current, current_uncached = read_reference(working, through)
    current_uncached_keys = {(r['period'], r['company'], r['field']) for r in current_uncached}
    from .core import COMPANIES
    names = {short: name for name, short in COMPANIES}
    selected = {s for n,s in COMPANIES if not company_filter or n in company_filter or s in company_filter}
    if not selected:
        raise ValueError("대상 회사가 없습니다.")
    answer = {key: value for key, value in answer.items() if key[1] in selected and (not period or key[0] == period)}
    if not answer:
        raise ValueError("선택한 기간·회사에 정답 기준값이 없습니다.")
    comparisons, rules = [], []
    for (p,s,c), expected in answer.items():
        value = current.get((p,s,c))
        status, delta, tolerance = compare(value, expected, c)
        if status == "NO_VALUE" and (p,s,c) in current_uncached_keys:
            status = "UNCALCULATED"
        comparisons.append(dict(period=p,company=s,field=c,answer=expected,value=value,status=status,difference=delta,tolerance=tolerance))
    import ir_parse
    for p,s in sorted({(p,s) for p,s,c in answer}):
        name = names[s]
        if name not in ir_parse.SPECS:
            continue
        folder = root / "data" / name / "IR" / p
        files = sorted(folder.glob("*.xlsx"))
        if not files:
            continue
        log(f"과거 규칙 대조: {name} {p}")
        for path in files:
            try:
                values, _ = ir_parse.extract_company(name, str(path), p)
                for c,value in values.items():
                    if (p,s,c) not in answer:
                        continue
                    expected = answer[(p,s,c)]
                    status, delta, tolerance = compare(value,expected,c)
                    rules.append(dict(period=p,company=s,field=c,answer=expected,value=value,status=status,difference=delta,tolerance=tolerance,source=str(path)))
            except Exception as exc:
                rules.append(dict(period=p,company=s,status="EXTRACTION_ERROR",source=str(path),detail=type(exc).__name__))
    folder = root / "runs" / "baseline_checks" / korea_now().strftime("%Y%m%d_%H%M%S_%f")
    dump(folder / "working_vs_answer.json", comparisons)
    dump(folder / "rules_vs_answer.json", rules)
    dump(folder / "uncalculated_answer_formulas.json", uncached)
    summary = dict(reference=str(reference),reference_sha256=ref_sha,workbook=str(working),workbook_sha256=work_sha,
                   through=through,working_compared=len(comparisons),working_differences=sum(c["status"]=="DIFF" for c in comparisons),
                   working_missing=sum(c["status"]=="NO_VALUE" for c in comparisons),rules_compared=len(rules),rules_differences=sum(c["status"]=="DIFF" for c in rules),
                   working_uncalculated=sum(c["status"]=="UNCALCULATED" for c in comparisons),
                   answer_uncalculated_formulas=len(uncached))
    if digest(reference) != ref_sha or digest(working) != work_sha:
        raise ValueError("대조 중 엑셀 원본이 변경됐습니다.")
    dump(folder / "summary.json", summary)
    log(f"정답본 대조 저장: {folder}")
    log(str(summary))
    return folder
