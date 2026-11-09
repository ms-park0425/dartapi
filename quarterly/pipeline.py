from __future__ import annotations

import csv
from copy import copy
from collections import defaultdict
from pathlib import Path
import shutil
import subprocess
import uuid

from .core import COMPANIES, Period, digest, dump, load, korea_now, numeric
from .extract import extract_rules
from .review import CodexReviewer, check_finding, field_catalog, prepare_assets
from .sources import Collector, CodexSourceFinder
from .workbook import DataBook, validate
from .source_rules import company_catalog


def choose_values(candidates):
    grouped = defaultdict(list)
    for c in candidates:
        grouped[(c["short"], c["field"])].append(c)
    chosen, conflicts = {}, []
    for key, group in grouped.items():
        # Keep legacy source selection for the initial pass; PDF review can correct it.
        selected = group[0]
        chosen[key] = selected["value"]
        if any(abs(c["value"] - selected["value"]) > max(1e-6, abs(selected["value"]) * 0.001) for c in group[1:]):
            conflicts.append(dict(company=key[0], field=key[1], candidates=group))
    return chosen, conflicts


def snapshot_sources(run, sources):
    for s in sources:
        original = Path(s["path"])
        dest = Path(run) / "sources" / s["company"] / (s["sha256"] + original.suffix.lower())
        dest.parent.mkdir(parents=True, exist_ok=True)
        if not dest.exists():
            shutil.copy2(original, dest)
        s["original_path"], s["path"] = str(original), str(dest.resolve())
    dump(Path(run) / "source_manifest.json", sources)


def prepare(root, period, config, offline=False, company_filter=None, run=None, log=print):
    root = Path(root).resolve()
    workbook = (root / config["workbook"]).resolve()
    through = config.get("protect_through", "2512")
    if int(period.code) <= int(through):
        raise ValueError(f"{through}까지는 정답 기준 기간입니다. --check-baseline으로 읽기 전용 검증하세요.")
    previous = Period(period.year, period.quarter-1) if period.quarter > 1 else Period(period.year-1, 4)
    through = str(max(int(through), int(previous.code))).zfill(4)
    config = {**config, 'protect_through': through}
    if not workbook.is_file():
        raise FileNotFoundError(f"작업 엑셀을 찾을 수 없습니다: {workbook}")
    companies = [c for c in COMPANIES if not company_filter or c[0] in company_filter or c[1] in company_filter]
    if not companies:
        raise ValueError("대상 회사가 없습니다.")
    run = Path(run).resolve() if run else root / "runs" / period.code / (korea_now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6])
    if (run / "state.json").exists():
        raise ValueError("기존 실행 폴더입니다. --resume을 사용하세요.")
    run.mkdir(parents=True, exist_ok=True)
    state = dict(period=period.code, companies=companies, config=config, input=str(workbook),
                 input_sha256=digest(workbook), status="collecting", started_at=korea_now().isoformat())
    reference = (root / config["reference_workbook"]).resolve() if config.get("reference_workbook") else None
    if reference:
        if reference == workbook:
            raise ValueError("정답 기준 파일을 작업 엑셀로 사용할 수 없습니다.")
        state.update(reference=str(reference), reference_sha256=digest(reference))
    dump(run / "state.json", state)
    shutil.copy2(workbook, run / "input.xlsx")
    log(f"[{period.code}] 자료 수집 시작 ({len(companies)}개사)")
    import parse_annual
    import dart_api
    collector = Collector(root, run, period, offline=offline or config.get("cached_sources_only", False), browser=config.get("browser", False), max_pages=config.get("discovery_pages", 12))
    finder = CodexSourceFinder(config.get("codex", "codex")) if not offline and config.get("ai_source_discovery", True) else None
    for name, short in companies:
        log(f"  수집: {name}")
        collector.cached(name)
        try:
            collector.dart(name, parse_annual.ALL_CORP_CODES[name], dart_api.API_KEY)
        except Exception as exc:
            collector.issue(name, "DART", "collection_failed", type(exc).__name__)
        collector.ir(name, config.get("sources", {}).get(name))
        if config.get("use_fisis", False) and not offline and not dart_api.FISIS_API_KEY:
            collector.issue(name, "FISIS", "missing_key", "FISIS_API_KEY 미설정")
        if finder and not any(s["company"] == name and s["kind"] in ("pdf", "disclosure_pdf") for s in collector.records):
            log(f"  공식 자료실 추가 탐색: {name}")
            try:
                for found in finder.find(name, period, run):
                    try:
                        collector.download(found["url"], name, found["title"])
                    except Exception as exc:
                        collector.issue(name, "AI_discovery", "download_failed", type(exc).__name__, found["url"])
            except Exception as exc:
                collector.issue(name, "AI_discovery", "discovery_failed", type(exc).__name__)
    sources = collector.finish()
    snapshot_sources(run, sources)
    dump(run/'publication_manifest.json', [dict(company=name, short=short,
        exact_dart_receipts=sorted({s['receipt'] for s in sources if s['company']==name and s.get('receipt')}),
        pdf_sources=sum(s['company']==name and s['kind'].endswith('pdf') for s in sources),
        factsheet_sources=sum(s['company']==name and s['kind']=='factsheet' for s in sources),
        requires_content_review=True) for name,short in companies])
    state["status"] = "extracting"
    dump(run / "state.json", state)
    log("규칙 기반 추출과 엑셀 1차 입력")
    candidates = extract_rules(root, run, period, companies, sources, offline=offline, use_fisis=config.get("use_fisis", False))
    for path in sorted((run / "api_traces").glob("*.json")):
        trace = load(path, {})
        finance = trace.get("params", {}).get("financeCd")
        name = next((n for n, entry in dart_api.FISIS_CORP_CODES.items() if entry[0] == finance), None)
        if name:
            sha = digest(path)
            source = dict(id=f"{name}:fisis:{sha[:16]}", company=name, period=period.code, kind="fisis_api_json",
                          path=str(path.resolve()), sha256=sha, url=trace.get("url"), params=trace.get("params"))
            sources.append(source)
            for candidate in candidates:
                if candidate["company"] == name and candidate["method"] in ("legacy_dart_fisis_rules", "fisis_rules"):
                    candidate["source_ids"].append(source["id"])
    dump(run / "source_manifest.json", sources)
    dump(run / "candidates.json", candidates)
    book = DataBook(run / "input.xlsx", protect_through=through)
    dump(run / "protected_history.json", book.protected_snapshot)
    book.ensure_period(period, companies)
    input_fields = field_catalog(book, config.get("fields"))
    chosen, conflicts = choose_values(candidates)
    changes = []
    for (short, field), value in chosen.items():
        if field not in book.columns or str(field) not in input_fields:
            continue
        current = book.values(period.code, short).get(field)
        if isinstance(current, str) and current.startswith("="):
            continue
        if current != value:
            changes.append(book.apply(period.code, short, field, value))
    first = run / "stage1.xlsx"
    book.save(first)
    dump(run / "stage1_changes.json", changes)
    dump(run / "source_conflicts.json", conflicts)
    dump(run / "validation_before.json", validate(book, period.code, companies))
    catalog = field_catalog(book, config.get("fields"))
    dump(run / "field_catalog.json", catalog)
    dump(run / "current_values.json", {short: book.values(period.code, short) for _, short in companies})
    log("PDF 페이지 렌더링 및 2차 검토 준비")
    if config.get("review_keywords"):
        prepare_assets(run, sources, companies, config.get("max_review_pages", 0), keywords=config["review_keywords"])
    else:
        prepare_assets(run, sources, companies, config.get("max_review_pages", 0))
    state.update(status="awaiting_pdf_review", stage1_sha256=digest(first), stage1_writes=len(changes),
                 catalog_sha256=digest(run / "field_catalog.json"), current_sha256=digest(run / "current_values.json"),
                 sources=len(sources), prepared_at=korea_now().isoformat())
    dump(run / "state.json", state)
    (run / "ASSISTANT_REVIEW.md").write_text(
        "# 이 실행의 PDF 직접 검토\n\n사용자가 요청한 작업은 최종 엑셀 수정까지입니다. 이 파일은 중간 산출물입니다.\n"
        "review_assets.json의 PDF 페이지 이미지를 직접 읽고 field_catalog.json과 current_values.json에 대조하세요.\n"
        "quarterly.review.review_schema() 형식으로 각 결과에 short를 추가한 findings 배열을 review_results.json에 기록하세요.\n"
        "근거가 확인된 KEEP/FILL/CORRECT만 기록하고 모호한 건은 CONFLICT로 남기세요. 사용자에게 결과 파일 작성을 맡기지 마세요.\n"
        "모든 회사의 검토 후 quarter_pipeline.py --finalize 이_실행폴더 를 실행해 final.xlsx를 저장하세요.\n",
        encoding="utf-8")
    return run


def review_run(run, reviewer=None, log=print):
    run = Path(run).resolve()
    state = load(run / "state.json")
    if digest(run / "stage1.xlsx") != state["stage1_sha256"]:
        raise ValueError("1차 엑셀이 변경됐습니다. 새 실행이 필요합니다.")
    for file, key in (("field_catalog.json", "catalog_sha256"), ("current_values.json", "current_sha256")):
        if state.get(key) and digest(run / file) != state[key]:
            raise ValueError("항목 정의 또는 현재값이 변경됐습니다. 새 실행이 필요합니다.")
    reviewer = reviewer or CodexReviewer(state["config"].get("codex", "codex"), state["config"].get("review_timeout", 900))
    period, catalog = Period.parse(state["period"]), load(run / "field_catalog.json")
    assets, current = load(run / "review_assets.json"), load(run / "current_values.json")
    for a in assets:
        if digest(Path(a["image"])) != a["image_sha256"] or digest(Path(a["pdf_path"])) != a["pdf_sha256"]:
            raise ValueError("검토 자료가 변경됐습니다.")
        if a.get("text_sha256") and digest(Path(a["text"])) != a["text_sha256"]:
            raise ValueError("검토 텍스트가 변경됐습니다.")
    findings, pages = [], []
    index = 0
    state["status"] = "reviewing_pdfs"
    dump(run / "state.json", state)
    for company, short in state["companies"]:
        own = [a for a in assets if a["short"] == short]
        size = state["config"].get("review_batch_pages", 4)
        if not isinstance(size, int) or not 1 <= size <= 8:
            raise ValueError("review_batch_pages는 1~8이어야 합니다.")
        for start in range(0, len(own), size):
            batch = own[start:start+size]
            index += 1
            log(f"PDF 직접 검토: {company} {start+1}~{start+len(batch)}/{len(own)}쪽")
            result = reviewer.review(batch, period, company_catalog(catalog, company, state['config']), current[short], run, index)
            if not isinstance(result, list):
                raise ValueError("검토 결과 배열이 아닙니다.")
            findings.extend({**f, "short": short} for f in result)
            pages.extend({"source_id": a["source_id"], "page": a["page"], "short": short} for a in batch)
            dump(run / "review_results.json", {"findings": findings, "reviewed_pages": pages})
    dump(run / "review_results.json", {"findings": findings, "reviewed_pages": pages})
    return finalize(run, log=log)


def recalculate(path, run):
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if not executable:
        for p in (Path("C:/Program Files/LibreOffice/program/soffice.exe"), Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe")):
            if p.exists():
                executable = str(p)
                break
    if not executable:
        return "pending_excel_recalculation"
    folder = Path(run) / "recalculated"
    folder.mkdir(exist_ok=True)
    profile = (Path(run) / "lo_profile").resolve().as_uri()
    result = subprocess.run([executable, f"-env:UserInstallation={profile}", "--headless", "--convert-to", "xlsx", "--outdir", str(folder), str(path)],
                            capture_output=True, timeout=120)
    candidate = folder / Path(path).name
    if result.returncode or not candidate.exists():
        return "recalculation_failed"
    shutil.copy2(candidate, path)
    return "recalculated"


def finalize(run, log=print):
    run = Path(run).resolve()
    state, results = load(run / "state.json"), load(run / "review_results.json")
    if not state or not isinstance(results, dict) or not isinstance(results.get("findings"), list):
        raise ValueError("유효한 PDF 검토 결과가 필요합니다.")
    if digest(run / "stage1.xlsx") != state["stage1_sha256"]:
        raise ValueError("1차 엑셀 변경 충돌")
    for file, key in (("field_catalog.json", "catalog_sha256"), ("current_values.json", "current_sha256")):
        if state.get(key) and digest(run / file) != state[key]:
            raise ValueError("항목 정의 또는 현재값 변경 충돌")
    if state.get("status") in ("completed", "completed_with_unresolved") and (run / "final.xlsx").exists():
        if state.get("review_results_sha256") == digest(run / "review_results.json"):
            return run / "final.xlsx"
    period, catalog, assets = Period.parse(state["period"]), load(run / "field_catalog.json"), load(run / "review_assets.json")
    for s in load(run / "source_manifest.json"):
        if digest(Path(s["path"])) != s["sha256"]:
            raise ValueError("PDF/원본 근거가 변경됐습니다.")
    for a in assets:
        if digest(Path(a["image"])) != a["image_sha256"]:
            raise ValueError("검토 이미지가 변경됐습니다.")
        if a.get("text_sha256") and digest(Path(a["text"])) != a["text_sha256"]:
            raise ValueError("검토 텍스트가 변경됐습니다.")
    book = DataBook(run / "stage1.xlsx", protect_through=state["config"].get("protect_through", "2512"))
    book.assert_history_unchanged(load(run / "protected_history.json", book.protected_snapshot))
    if state.get("reference") and digest(Path(state["reference"])) != state["reference_sha256"]:
        raise ValueError("2512backup 정답 기준 파일이 변경됐습니다.")
    companies = [tuple(c) for c in state["companies"]]
    baseline_checks = {(c["company"], c["check"]): c["status"] for c in validate(book, period.code, companies)}
    previous_styles = {}
    company_names = {short: name for name, short in companies}
    valid, rejected, unresolved, changes = defaultdict(list), [], [], []
    for f in results["findings"]:
        short = f.get("short") if isinstance(f, dict) else None
        try:
            if short not in company_names:
                raise ValueError("대상 회사 아님")
            current = book.values(period.code, short).get(f.get("field"))
            checked = check_finding(f, period, company_catalog(catalog, company_names[short], state['config']), assets, current, company_names[short], short)
            if checked and checked.get("decision") == "CONFLICT":
                unresolved.append(dict(company=short, field=f.get("field"), reason=checked["reason"]))
            elif checked:
                valid[(short, f["field"])].append((checked, f))
            else:
                unresolved.append(dict(company=short, field=f.get("field"), reason=f.get("reason", "판단 보류")))
        except (ValueError, TypeError, KeyError) as exc:
            rejected.append(dict(finding=f, reason=str(exc)))
    verified = set()
    for (short, field), group in valid.items():
        # Compatible sources may display different precision. Prefer the
        # narrowest source interval, rather than whichever page came first.
        group.sort(key=lambda entry: entry[0]['tolerance'])
        first = group[0][0]
        # Source disagreements remain visible; do not pick whichever PDF was processed last.
        if any(abs(c["value"] - first["value"]) > max(c["tolerance"], first["tolerance"]) for c, _ in group[1:]):
            unresolved.append(dict(company=short, field=field, reason="PDF 출처 간 값 충돌"))
            continue
        if any(u["company"] == short and u["field"] == field for u in unresolved):
            continue
        verified.add((short, field))
        current = book.values(period.code, short).get(field)
        if first["value"] != current:
            previous_styles[(short, field)] = copy(book.ws.cell(book.rows[(period.code, short)], book.columns[field])._style)
            change = book.apply(period.code, short, field, first["value"], color="FFE0B2", expected=current)
            change.update(reason=group[0][1]["reason"], evidence=[f for _, f in group])
            changes.append(change)
    held = []
    for check in validate(book, period.code, companies):
        if check["status"] != "fail" or baseline_checks[(check["company"], check["check"])] != "pass":
            continue
        for change in list(changes):
            if change["company"] != check["company"] or change["field"] not in check["fields"]:
                continue
            cell = book.ws.cell(book.rows[(period.code, change["company"])], book.columns[change["field"]])
            cell.value = change["old"]
            cell._style = previous_styles[(change["company"], change["field"])]
            verified.discard((change["company"], change["field"]))
            changes.remove(change)
            held.append({**change, "hold_reason": "수정으로 기존 통과 검산이 실패: " + check["check"]})
            unresolved.append(dict(company=change["company"], field=change["field"], reason=held[-1]["hold_reason"]))
    for _, short in companies:
        values = book.values(period.code, short)
        for field in catalog:
            c = int(field)
            if (short, c) not in verified and not (isinstance(values.get(c), str) and values[c].startswith("=")):
                unresolved.append(dict(company=short, field=c, reason="PDF 원문 검증 미완료" if values.get(c) is not None else "공시값 미확인"))
    output = run / "final.xlsx"
    book.save(output)
    calc = recalculate(output, run)
    checks = validate(book, period.code, companies)
    dump(run / "changes.json", changes)
    dump(run / "held_changes.json", held)
    dump(run / "rejected_reviews.json", rejected)
    dump(run / "validation_after.json", checks)
    with (run / "unresolved.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["company", "field", "reason"])
        writer.writeheader()
        writer.writerows(unresolved)
    expected_pages = {(a["source_id"], a["page"], a["short"]) for a in assets}
    reviewed_pages = {(a["source_id"], a["page"], a["short"]) for a in results.get("reviewed_pages", [])}
    summary = dict(period=period.code, companies=len(companies), stage1_writes=state["stage1_writes"],
                   pdf_verified_fields=len(verified), stage2_changes=len(changes), unresolved_fields=len({(u["company"], u["field"]) for u in unresolved}),
                   held_changes=len(held),
                   rejected_reviews=len(rejected), pdf_pages=len(expected_pages), reviewed_pages=len(expected_pages & reviewed_pages),
                   failed_checks=sum(c["status"] == "fail" for c in checks), recalculation=calc,
                   source_issues=len(load(run / "collection_issues.json", [])),
                   render_issues=len(load(run / "review_asset_issues.json", [])), output=str(output))
    pending = unresolved or rejected or expected_pages - reviewed_pages or not assets or summary["render_issues"] or summary["failed_checks"] or calc != "recalculated"
    state.update(status="completed_with_unresolved" if pending else "completed", summary=summary,
                 review_results_sha256=digest(run / "review_results.json"), completed_at=korea_now().isoformat())
    dump(run / "summary.json", summary)
    dump(run / "state.json", state)
    (run / "summary.md").write_text("# 분기 공시 입력 결과\n\n" + "\n".join(f"- {k}: {v}" for k, v in summary.items()), encoding="utf-8")
    log(f"저장: {output}\n1차 {summary['stage1_writes']}셀 / 2차 수정 {len(changes)}셀 / PDF 검증 {len(verified)}항목 / 미해결 {summary['unresolved_fields']}항목")
    return output
