"""
손보사 5개 XBRL 파일에서 BEL/RA/CSM/OCI 구조 파악
정답값에 매칭되는 tag/context 구조를 출력
"""
import xml.etree.ElementTree as ET
from pathlib import Path

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"

COMPANIES = {
    "삼성화재": {
        "path": "C:/MALIFE/dartapi/data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
        "targets": {
            "BEL": 78736684,
            "RA": 1616009,
            "CSM": 8897609,
            "Col25": 3944864,
            "Col26": 1001660,
            "Col27": 3953,
            "Col28": -217651,
            "Col30": -4782,
        }
    },
    "현대해상": {
        "path": "C:/MALIFE/dartapi/data/현대해상/사업보고서/2025/entity00164973_2025-12-31.xbrl",
        "targets": {
            "BEL": 43022017,
            "RA": 873617,
            "CSM": 4490254,
            "Col25": -1060623,
        }
    },
    "메리츠화재": {
        "path": "C:/MALIFE/dartapi/data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
        "targets": {
            "BEL": 38011337,
            "RA": 817895,
            "CSM": 3413560,
            "Col25": -876506,
        }
    },
    "KB손해보험": {
        "path": "C:/MALIFE/dartapi/data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
        "targets": {
            "BEL": 39028451,
            "RA": 812040,
            "CSM": 3402947,
            "Col25": -827590,
        }
    },
    "DB손해보험": {
        "path": "C:/MALIFE/dartapi/data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl",
        "targets": {
            "BEL": 52738440,
            "RA": 1058497,
            "CSM": 3983500,
            "Col25": -389340,
        }
    },
}

# BEL/RA/CSM 관련 태그 키워드
BEL_KEYWORDS = ["BestEstimateLiability", "PresentValue", "FullyaimedLiability"]
RA_KEYWORDS = ["RiskAdjustment", "RiskAdjust"]
CSM_KEYWORDS = ["ContractualServiceMargin", "CSM"]
OCI_KEYWORDS = ["OtherComprehensiveIncome", "OCI", "AccumulatedOtherComprehensiveIncome",
                 "ComponentsOfEquity", "OtherComprehensiveLoss"]


def get_local_name(tag):
    """'{ns}localname' -> 'localname'"""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def get_ns_prefix(tag):
    """'{ns}localname' -> ns uri"""
    if "{" in tag and "}" in tag:
        return tag[1:].split("}", 1)[0]
    return ""


def parse_context(ctx_elem):
    """context element -> dict with period and dimensions"""
    result = {"id": ctx_elem.get("id", ""), "dims": []}
    period = ctx_elem.find(f"{{{NS_XBRLI}}}period")
    if period is not None:
        instant = period.find(f"{{{NS_XBRLI}}}instant")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if instant is not None:
            result["period_type"] = "instant"
            result["date"] = instant.text
        elif start is not None and end is not None:
            result["period_type"] = "duration"
            result["start"] = start.text
            result["end"] = end.text

    entity = ctx_elem.find(f"{{{NS_XBRLI}}}entity")
    if entity is not None:
        segment = entity.find(f"{{{NS_XBRLI}}}segment")
        if segment is not None:
            for mem in segment.findall(f"{{{NS_XBRLDI}}}explicitMember"):
                dim = mem.get("dimension", "")
                val = (mem.text or "").strip()
                result["dims"].append((get_local_name(dim), get_local_name(val) if val else val, dim, val))
    return result


def analyze_xbrl(company_name, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company_name}] {xbrl_path}")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    # 모든 context 파싱
    contexts = {}
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        info = parse_context(ctx)
        contexts[info["id"]] = info

    # instant 2025-12-31 + Separate context 목록
    sep_instant_ctx = {}
    for ctx_id, info in contexts.items():
        if info.get("period_type") != "instant":
            continue
        if info.get("date") != "2025-12-31":
            continue
        # Separate 여부
        is_separate = any("Separate" in d[1] or "SeparateMember" in d[1] for d in info["dims"])
        is_consolidated = any("Consolidated" in d[1] or "ConsolidatedMember" in d[1] for d in info["dims"])
        if is_separate:
            sep_instant_ctx[ctx_id] = info
        elif not is_consolidated:
            # dim 없는 것도 포함 (ConsolidatedAndSeparate 축 없는 경우)
            sep_instant_ctx[ctx_id] = info

    # 모든 수치 태그를 모아서 정답값과 매칭
    target_vals = list(targets.values())
    target_items = list(targets.items())

    print(f"\n  instant 2025-12-31 contexts 수: {len(sep_instant_ctx)} (Separate/기타 포함)")

    # 정답값 탐색
    found = {k: [] for k in targets.keys()}

    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            val = float(elem.text.strip())
        except ValueError:
            continue

        local = get_local_name(elem.tag)
        ns_uri = get_ns_prefix(elem.tag)

        # 정답값 매칭 (절대값 포함, 단위 감안)
        for label, tgt in target_items:
            if abs(val - tgt) < 0.5 or abs(val - tgt * 1e6) < 0.5:
                matched_val = val
                ctx_info = contexts.get(ctx_ref, {})
                found[label].append({
                    "tag": local,
                    "ns": ns_uri,
                    "val": matched_val,
                    "ctx_id": ctx_ref,
                    "ctx": ctx_info,
                })

    # 결과 출력
    print(f"\n  [BEL/RA/CSM 탐색 결과]")
    for col in ["BEL", "RA", "CSM"]:
        tgt = targets.get(col)
        if tgt is None:
            continue
        print(f"\n  {col}={tgt}")
        hits = found.get(col, [])
        if not hits:
            print(f"    -> 정확 매칭 없음. 근사값 탐색...")
            # 근사값 탐색 (±1%)
            for elem in root:
                ctx_ref = elem.get("contextRef", "")
                if not elem.text:
                    continue
                try:
                    val = float(elem.text.strip())
                except ValueError:
                    continue
                if tgt != 0 and abs(val - tgt) / abs(tgt) < 0.01:
                    local = get_local_name(elem.tag)
                    ctx_info = contexts.get(ctx_ref, {})
                    print(f"    근사 tag={local}, val={val}, dims={[d[:2] for d in ctx_info.get('dims', [])]}, period={ctx_info.get('date', ctx_info.get('end','?'))}")
        else:
            for h in hits[:5]:
                ctx = h["ctx"]
                dims = ctx.get("dims", [])
                period_str = ctx.get("date", ctx.get("end", "?"))
                is_sep = any("Separate" in d[1] for d in dims)
                is_con = any("Consolidated" in d[1] for d in dims)
                fs_label = "Separate" if is_sep else ("Consolidated" if is_con else "No_FS_dim")
                n_dim = len(dims)
                dim_summary = [(d[0], d[1]) for d in dims]
                print(f"    tag={h['tag']}, val={h['val']}, date={period_str}, {fs_label}, n_dim={n_dim}")
                for d in dim_summary:
                    print(f"      axis={d[0]}, member={d[1]}")

    print(f"\n  [OCI 누계 탐색 결과 (Col25~30)]")
    for col in ["Col25", "Col26", "Col27", "Col28", "Col30"]:
        tgt = targets.get(col)
        if tgt is None:
            continue
        print(f"\n  {col}={tgt}")
        hits = found.get(col, [])
        if not hits:
            print(f"    -> 정확 매칭 없음. 근사값 탐색 (±1%)...")
            # 근사값
            for elem in root:
                ctx_ref = elem.get("contextRef", "")
                if not elem.text:
                    continue
                try:
                    val = float(elem.text.strip())
                except ValueError:
                    continue
                if tgt != 0 and abs(val - tgt) / abs(tgt) < 0.01:
                    local = get_local_name(elem.tag)
                    ctx_info = contexts.get(ctx_ref, {})
                    dims = ctx_info.get("dims", [])
                    period_str = ctx_info.get("date", ctx_info.get("end", "?"))
                    print(f"    근사 tag={local}, val={val}, date={period_str}, n_dim={len(dims)}")
                    for d in dims:
                        print(f"      axis={d[0]}, member={d[1]}")
        else:
            for h in hits[:5]:
                ctx = h["ctx"]
                dims = ctx.get("dims", [])
                period_str = ctx.get("date", ctx.get("end", "?"))
                is_sep = any("Separate" in d[1] for d in dims)
                is_con = any("Consolidated" in d[1] for d in dims)
                fs_label = "Separate" if is_sep else ("Consolidated" if is_con else "No_FS_dim")
                n_dim = len(dims)
                dim_summary = [(d[0], d[1]) for d in dims]
                print(f"    tag={h['tag']}, val={h['val']}, date={period_str}, {fs_label}, n_dim={n_dim}")
                for d in dim_summary:
                    print(f"      axis={d[0]}, member={d[1]}")

    # 해약환급금준비금 탐색 (Col99)
    print(f"\n  [해약환급금준비금(Col99) 탐색]")
    refund_keywords = ["SurrenderValue", "Surrender", "해약환급", "RefundReserve", "GuaranteedSurrenderValue",
                       "PolicyholderDividendReserve", "ContractholderParticipation",
                       "CashSurrenderValue", "PolicyReserveForSurrenderValue"]
    refund_tags = []
    for elem in root:
        local = get_local_name(elem.tag)
        if any(kw.lower() in local.lower() for kw in ["surrender", "refund", "해약", "guarantee"]):
            ctx_ref = elem.get("contextRef", "")
            ctx_info = contexts.get(ctx_ref, {})
            if not elem.text:
                continue
            try:
                val = float(elem.text.strip())
            except ValueError:
                continue
            period_str = ctx_info.get("date", ctx_info.get("end", "?"))
            dims = ctx_info.get("dims", [])
            refund_tags.append((local, val, period_str, len(dims), [(d[0], d[1]) for d in dims]))

    if refund_tags:
        for tag, val, period, ndim, dims in refund_tags[:10]:
            print(f"    tag={tag}, val={val}, date={period}, n_dim={ndim}")
            for d in dims:
                print(f"      axis={d[0]}, member={d[1]}")
    else:
        # dart 태그 중 "reserve" 포함 탐색
        print("    직접 키워드 없음. dart reserve 태그 탐색...")
        for elem in root:
            local = get_local_name(elem.tag)
            ns_uri = get_ns_prefix(elem.tag)
            if "dart" in ns_uri.lower() and "reserve" in local.lower():
                ctx_ref = elem.get("contextRef", "")
                ctx_info = contexts.get(ctx_ref, {})
                if not elem.text:
                    continue
                try:
                    val = float(elem.text.strip())
                except ValueError:
                    continue
                period_str = ctx_info.get("date", ctx_info.get("end", "?"))
                dims = ctx_info.get("dims", [])
                print(f"    tag={local}, val={val}, date={period_str}, n_dim={len(dims)}")
                for d in dims:
                    print(f"      axis={d[0]}, member={d[1]}")


def main():
    for company, info in COMPANIES.items():
        analyze_xbrl(company, info["path"], info["targets"])

    print("\n\n" + "="*70)
    print("추가: 전체 dimension 구조 요약")
    print("="*70)
    for company, info in COMPANIES.items():
        xbrl_path = info["path"]
        tree = ET.parse(xbrl_path)
        root = tree.getroot()
        contexts = {}
        for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
            parsed = parse_context(ctx)
            contexts[parsed["id"]] = parsed

        # instant 2025-12-31 contexts의 dim 수 분포
        dim_dist = {}
        sep_ctx_count = 0
        for ctx_id, info2 in contexts.items():
            if info2.get("date") != "2025-12-31":
                continue
            dims = info2.get("dims", [])
            n = len(dims)
            dim_dist[n] = dim_dist.get(n, 0) + 1
            is_sep = any("Separate" in d[1] for d in dims)
            if is_sep:
                sep_ctx_count += 1

        print(f"\n  [{company}] instant 2025-12-31 dim 분포: {dim_dist} (Separate포함: {sep_ctx_count}개)")
        # Separate context 샘플
        shown = 0
        for ctx_id, info2 in contexts.items():
            if info2.get("date") != "2025-12-31":
                continue
            dims = info2.get("dims", [])
            is_sep = any("Separate" in d[1] for d in dims)
            if is_sep and shown < 3:
                print(f"    ctx_id={ctx_id}, n_dim={len(dims)}")
                for d in dims:
                    print(f"      {d[0]} = {d[1]}")
                shown += 1


if __name__ == "__main__":
    main()
