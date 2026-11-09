from __future__ import annotations

from contextlib import contextmanager
import os
import shutil
from pathlib import Path

from .core import dump, numeric


@contextmanager
def working_directory(root):
    before = Path.cwd()
    os.chdir(root)
    try:
        yield
    finally:
        os.chdir(before)


def extract_rules(root, run, period, companies, sources, offline=False, use_fisis=False):
    """Reuse proven parsers; keep each independently extracted candidate."""
    import parse_annual
    import ir_parse
    import dart_api
    out, issues = [], []
    old_offline, old_trace = os.environ.get("DART_OFFLINE"), os.environ.get("DART_TRACE_DIR")
    old_fisis = os.environ.get("ENABLE_FISIS")
    os.environ["ENABLE_FISIS"] = "1" if use_fisis else "0"
    os.environ["DART_OFFLINE"] = "1" if offline else "0"
    os.environ["DART_TRACE_DIR"] = str(Path(run).resolve() / "api_traces")
    try:
        with working_directory(root):
            for name, short in companies:
                own = [s for s in sources if s["company"] == name]
                xbrls = [s for s in own if s["kind"] == "xbrl"]
                if period.quarter == 4:
                    # The collector uses doc_YYMM.xml while legacy annual
                    # parsers expect this alias. Both must use the same snapshot.
                    docs = [s for s in own if s['kind']=='document_xml']
                    if docs:
                        alias = Path('data')/name/f'doc_annual_{period.year}.xml'
                        alias.parent.mkdir(parents=True,exist_ok=True)
                        if Path(docs[0]['path']).resolve() != alias.resolve():
                            shutil.copy2(docs[0]['path'],alias)
                if xbrls:
                    try:
                        values = parse_annual.extract_company_data(name, str(period.year), period.report_type, f"{period.month:02d}", xbrl_path=xbrls[0]["path"]) or {}
                        for field, value in values.items():
                            if numeric(value) and field > 3:
                                out.append(dict(company=name, short=short, field=field, value=value,
                                                method="legacy_dart_fisis_rules", source_ids=[s["id"] for s in own if s["kind"] in ("xbrl", "document_xml")],
                                                evidence_scope="extraction_bundle; PDF review required"))
                    except Exception as exc:
                        issues.append(dict(company=name, stage="rules", detail=type(exc).__name__))
                # FISIS can work even when a company's XBRL has not been published.
                elif use_fisis and not offline and dart_api.FISIS_API_KEY:
                    for fetch in (dart_api.fetch_fisis_balance, dart_api.fetch_fisis_kics):
                        for field, value in fetch(name, f"{period.year}{period.month:02d}").items():
                            if numeric(value):
                                out.append(dict(company=name, short=short, field=field, value=value,
                                                method="fisis_rules", source_ids=[], evidence_scope="api_traces"))
                if name in ir_parse.SPECS:
                    for source in own:
                        if source["kind"] != "factsheet":
                            continue
                        try:
                            evidence = []
                            values, missing = ir_parse.extract_company(name, source["path"], period.code, evidence=evidence)
                            for field, value in values.items():
                                if numeric(value):
                                    out.append(dict(company=name, short=short, field=field, value=value, method="factsheet_rules",
                                                    source_ids=[source["id"]], evidence=[e for e in evidence if e["field"] == field]))
                        except Exception as exc:
                            issues.append(dict(company=name, stage="factsheet", source=source["id"], detail=type(exc).__name__))
                # This legacy parser is annual-only: do not apply its year-only column heuristic to quarters.
                if period.quarter == 4:
                    from ir_pdf import parse_disclosure
                    for source in own:
                        if source["kind"] != "disclosure_pdf":
                            continue
                        try:
                            values, evidence = parse_disclosure(source["path"], period.year)
                            for field, value in values.items():
                                if numeric(value):
                                    out.append(dict(company=name, short=short, field=field, value=value, method="annual_pdf_rules",
                                                    source_ids=[source["id"]], evidence=evidence.get(field)))
                        except Exception as exc:
                            issues.append(dict(company=name, stage="pdf_rules", detail=type(exc).__name__))
    finally:
        for key, old in (("DART_OFFLINE", old_offline), ("DART_TRACE_DIR", old_trace)):
            if old is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = old
        if old_fisis is None:
            os.environ.pop("ENABLE_FISIS", None)
        else:
            os.environ["ENABLE_FISIS"] = old_fisis
    dump(Path(run) / "candidates.json", out)
    dump(Path(run) / "extraction_issues.json", issues)
    return out
