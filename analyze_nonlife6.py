"""
손보사 XBRL - BEL/RA/CSM 직접 탐색
- 단위 참조 확인
- InsuranceContractsLiabilityAsset + 컴포넌트 axis 탐색
- raw 값 직접 탐색 (스케일 무시)
"""
import xml.etree.ElementTree as ET

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"


def get_local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag


def parse_contexts(root):
    contexts = {}
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        cid = ctx.get("id", "")
        dims = []
        entity = ctx.find(f"{{{NS_XBRLI}}}entity")
        if entity is not None:
            seg = entity.find(f"{{{NS_XBRLI}}}segment")
            if seg is not None:
                for m in seg.findall(f"{{{NS_XBRLDI}}}explicitMember"):
                    dim_full = m.get("dimension", "")
                    val_full = (m.text or "").strip()
                    dims.append({
                        "axis": get_local(dim_full),
                        "member": get_local(val_full),
                    })
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        ptype, pdate = None, None
        if period is not None:
            inst = period.find(f"{{{NS_XBRLI}}}instant")
            en = period.find(f"{{{NS_XBRLI}}}endDate")
            if inst is not None:
                ptype, pdate = "instant", inst.text
            elif en is not None:
                ptype, pdate = "duration", en.text
        contexts[cid] = {"id": cid, "dims": dims, "ptype": ptype, "pdate": pdate}
    return contexts


def analyze_units(root):
    """유닛 정의 확인"""
    units = {}
    ns_xbrli_q = f"{{{NS_XBRLI}}}unit"
    for u in root.findall(f"{{{NS_XBRLI}}}unit"):
        uid = u.get("id", "")
        measure = u.find(f"{{{NS_XBRLI}}}measure")
        if measure is not None:
            units[uid] = measure.text
    return units


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    # 단위 확인
    units = analyze_units(root)
    print(f"  단위 정의: {list(units.items())[:5]}")

    # Sep instant 2025-12-31 contexts
    sep_2025 = {}
    for cid, c in contexts.items():
        if c["ptype"] == "instant" and c["pdate"] == "2025-12-31":
            dims = c["dims"]
            is_sep = any("Separate" in d["member"] for d in dims)
            if is_sep:
                sep_2025[cid] = c

    print(f"  Sep instant 2025-12-31: {len(sep_2025)}개")

    # Sep contexts dim 분포
    ndim_dist = {}
    for c in sep_2025.values():
        n = len(c["dims"])
        ndim_dist[n] = ndim_dist.get(n, 0) + 1
    print(f"  Sep contexts dim 분포: {ndim_dist}")

    # 1. InsuranceContracts 관련 태그 탐색 (Sep 2025-12-31)
    print(f"\n  [InsuranceContracts 관련 태그 (Sep, 2025-12-31)]")
    ic_tags = {}
    for elem in root:
        local = get_local(elem.tag)
        if "insurance" not in local.lower() and "InsurancContract" not in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in sep_2025:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
        key = (local, dim_key)
        ic_tags.setdefault(key, []).append(raw)

    # 대규모 값만 출력 (BEL 수준: 수십조 원 = 10^13 이상)
    large_vals = []
    for (local, dim_key), vals in ic_tags.items():
        v = vals[0]
        if abs(v) > 1e12:  # 1조 이상
            large_vals.append((abs(v), local, v, dim_key))
    large_vals.sort(reverse=True)
    for _, local, v, dim_key in large_vals[:20]:
        print(f"    tag={local}, val={v:.0f}")
        for a, m in dim_key[:4]:
            print(f"      {a} = {m}")

    # 2. BEL 정확값 직접 탐색 (scale variants)
    print(f"\n  [BEL/RA/CSM 직접 탐색 (Sep+All Instant 2025-12-31)]")
    all_instant_2025 = {cid for cid, c in contexts.items()
                        if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"}

    for label, tgt_mn in [("BEL", targets["BEL"]), ("RA", targets["RA"]), ("CSM", targets["CSM"])]:
        print(f"\n  {label}={tgt_mn}백만원 탐색:")
        # 여러 scale 시도
        found_any = False
        for scale_name, scale in [("raw=value(백만원)", 1), ("raw=value*1e3(천원)", 1000),
                                   ("raw=value*1e6(원)", 1_000_000),
                                   ("raw=value/1e3(십억)", 1/1000)]:
            target_raw = tgt_mn * scale
            for elem in root:
                ctx_ref = elem.get("contextRef", "")
                if ctx_ref not in all_instant_2025:
                    continue
                if not elem.text:
                    continue
                try:
                    raw = float(elem.text.strip())
                except ValueError:
                    continue
                if target_raw != 0:
                    err = abs(raw - target_raw) / abs(target_raw)
                else:
                    err = abs(raw)
                if err < 0.001:
                    local = get_local(elem.tag)
                    ctx = contexts[ctx_ref]
                    dims = ctx["dims"]
                    is_sep = any("Separate" in d["member"] for d in dims)
                    is_con = any("Consolidated" in d["member"] for d in dims)
                    fs = "Sep" if is_sep else ("Con" if is_con else "?")
                    dec = elem.get("decimals", "?")
                    unit = elem.get("unitRef", "?")
                    print(f"    [{scale_name}] err={err:.5f} | tag={local}")
                    print(f"      raw={raw}, dec={dec}, unit={unit}, {fs}, n_dim={len(dims)}")
                    for d in dims[:5]:
                        print(f"        {d['axis']} = {d['member']}")
                    found_any = True
                    break
            if found_any:
                break
        if not found_any:
            print(f"    -> 모든 scale에서 0.1% 이내 매칭 없음")

    # 3. Sep 2025-12-31 컨텍스트의 모든 보험계약 관련 값 dump (BEL 범위)
    print(f"\n  [Sep instant 2025-12-31 - 대규모 값 상위 10개]")
    large_sep = []
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in sep_2025:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        if abs(raw) < 1e10:
            continue
        local = get_local(elem.tag)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        large_sep.append((abs(raw), local, raw, len(dims), dims))

    large_sep.sort(reverse=True)
    seen_sig = set()
    cnt = 0
    for _, local, v, ndim, dims in large_sep:
        dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
        sig = (local, ndim, dim_key)
        if sig in seen_sig: continue
        seen_sig.add(sig)
        dec_s = ""
        print(f"    tag={local}, val={v:.0f}, n_dim={ndim}")
        for d in dims[:3]:
            print(f"      {d['axis']} = {d['member']}")
        cnt += 1
        if cnt >= 15: break

    # 4. 모든 Dec 분포 확인 + 유닛 참조
    print(f"\n  [unitRef 분포]")
    from collections import Counter
    unit_dist = Counter()
    for elem in root:
        u = elem.get("unitRef")
        if u: unit_dist[u] += 1
    print(f"    {dict(unit_dist.most_common(5))}")


COMPANIES = {
    "삼성화재": {
        "path": "C:/MALIFE/dartapi/data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
        "targets": {"BEL": 78736684, "RA": 1616009, "CSM": 8897609}
    },
    "메리츠화재": {
        "path": "C:/MALIFE/dartapi/data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
        "targets": {"BEL": 38011337, "RA": 817895, "CSM": 3413560}
    },
    "KB손해보험": {
        "path": "C:/MALIFE/dartapi/data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
        "targets": {"BEL": 39028451, "RA": 812040, "CSM": 3402947}
    },
}


if __name__ == "__main__":
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])
