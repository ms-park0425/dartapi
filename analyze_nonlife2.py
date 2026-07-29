"""
손보사 5개 XBRL - BEL/RA/CSM/OCI 정답값 탐색 (정밀 탐색)
"""
import xml.etree.ElementTree as ET

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"

COMPANIES = {
    "삼성화재": {
        "path": "C:/MALIFE/dartapi/data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
        "targets": {
            "BEL(Col73)": 78736684,
            "RA(Col74)": 1616009,
            "CSM(Col75)": 8897609,
            "OCI(Col25)": 3944864,
            "OCI(Col26)": 1001660,
            "OCI(Col27)": 3953,
            "OCI(Col28)": -217651,
            "OCI(Col30)": -4782,
        }
    },
    "현대해상": {
        "path": "C:/MALIFE/dartapi/data/현대해상/사업보고서/2025/entity00164973_2025-12-31.xbrl",
        "targets": {
            "BEL(Col73)": 43022017,
            "RA(Col74)": 873617,
            "CSM(Col75)": 4490254,
            "OCI(Col25)": -1060623,
        }
    },
    "메리츠화재": {
        "path": "C:/MALIFE/dartapi/data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
        "targets": {
            "BEL(Col73)": 38011337,
            "RA(Col74)": 817895,
            "CSM(Col75)": 3413560,
            "OCI(Col25)": -876506,
        }
    },
    "KB손해보험": {
        "path": "C:/MALIFE/dartapi/data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
        "targets": {
            "BEL(Col73)": 39028451,
            "RA(Col74)": 812040,
            "CSM(Col75)": 3402947,
            "OCI(Col25)": -827590,
        }
    },
    "DB손해보험": {
        "path": "C:/MALIFE/dartapi/data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl",
        "targets": {
            "BEL(Col73)": 52738440,
            "RA(Col74)": 1058497,
            "CSM(Col75)": 3983500,
            "OCI(Col25)": -389340,
        }
    },
}


def get_local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag

def get_ns(tag):
    return tag[1:].split("}", 1)[0] if "{" in tag else ""

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
                        "axis_full": dim_full,
                        "axis": get_local(dim_full),
                        "member_full": val_full,
                        "member": get_local(val_full),
                    })
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        ptype, pdate, pstart, pend = None, None, None, None
        if period is not None:
            inst = period.find(f"{{{NS_XBRLI}}}instant")
            st = period.find(f"{{{NS_XBRLI}}}startDate")
            en = period.find(f"{{{NS_XBRLI}}}endDate")
            if inst is not None:
                ptype, pdate = "instant", inst.text
            elif st is not None and en is not None:
                ptype, pstart, pend = "duration", st.text, en.text
        contexts[cid] = {"id": cid, "dims": dims, "ptype": ptype, "pdate": pdate, "pstart": pstart, "pend": pend}
    return contexts


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    # instant 2025-12-31 context id 세트
    instant_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"}
    # duration ending 2025-12-31 context id 세트
    dur_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "duration" and c["pend"] == "2025-12-31"}

    print(f"  instant 2025-12-31 contexts: {len(instant_2025)}")
    print(f"  duration end 2025-12-31 contexts: {len(dur_2025)}")

    # Build value index: val -> [(tag_local, ctx_id)]
    val_index = {}
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            val = int(float(elem.text.strip()))
        except ValueError:
            continue
        local = get_local(elem.tag)
        val_index.setdefault(val, []).append((local, ctx_ref))

    for label, tgt in targets.items():
        print(f"\n  -- {label} = {tgt} --")
        # 정확 매칭
        hits = val_index.get(tgt, [])
        # instant 2025-12-31 hits
        inst_hits = [(t, c) for t, c in hits if c in instant_2025]
        dur_hits = [(t, c) for t, c in hits if c in dur_2025]
        other_hits = [(t, c) for t, c in hits if c not in instant_2025 and c not in dur_2025]

        if not hits:
            print(f"    정확 매칭 없음")
            # 부동소수점 오차 ±1 탐색
            for delta in [1, -1, 2, -2]:
                hits2 = val_index.get(tgt + delta, [])
                if hits2:
                    print(f"    ±{abs(delta)} 근사: {len(hits2)}개")
                    for t, c in hits2[:3]:
                        ctx = contexts.get(c, {})
                        dims = ctx.get("dims", [])
                        is_sep = any("Separate" in d["member"] for d in dims)
                        is_con = any("Consolidated" in d["member"] for d in dims)
                        fs = "Sep" if is_sep else ("Con" if is_con else "?")
                        print(f"      tag={t}, {fs}, n_dim={len(dims)}, date={ctx.get('pdate', ctx.get('pend','?'))}")
                        for d in dims:
                            print(f"        {d['axis']} = {d['member']}")
                    break
            continue

        all_show = []
        priority = inst_hits if inst_hits else (dur_hits if dur_hits else other_hits)
        for t, c in priority[:5]:
            ctx = contexts.get(c, {})
            dims = ctx.get("dims", [])
            is_sep = any("Separate" in d["member"] for d in dims)
            is_con = any("Consolidated" in d["member"] for d in dims)
            fs = "Sep" if is_sep else ("Con" if is_con else "?")
            pdate = ctx.get("pdate") or ctx.get("pend", "?")
            all_show.append((t, c, fs, len(dims), dims, pdate))

        # 중복 제거 (같은 tag + dims 패턴)
        seen = set()
        for t, c, fs, ndim, dims, pdate in all_show:
            dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
            sig = (t, fs, ndim, dim_key)
            if sig in seen:
                continue
            seen.add(sig)
            print(f"    tag={t}, {fs}, n_dim={ndim}, date={pdate}")
            for d in dims:
                print(f"      axis={d['axis']}")
                print(f"      member={d['member']}")

        if len(inst_hits) == 0 and dur_hits:
            print(f"    (instant 매칭 없음 - duration hits={len(dur_hits)})")
        elif len(inst_hits) > 0:
            print(f"    (instant hits: {len(inst_hits)}, total hits: {len(hits)})")

    # 해약환급금준비금 탐색
    print(f"\n  -- 해약환급금준비금(Col99) 탐색 --")
    refund_found = []
    for elem in root:
        local = get_local(elem.tag)
        # 태그명에 surrender / refund / 해약 / PolicyReserve 포함
        low = local.lower()
        if any(k in low for k in ["surrender", "refundreserve", "policyreserve", "cashvalue"]):
            ctx_ref = elem.get("contextRef", "")
            if not elem.text:
                continue
            try:
                val = float(elem.text.strip())
            except ValueError:
                continue
            ctx = contexts.get(ctx_ref, {})
            dims = ctx.get("dims", [])
            pdate = ctx.get("pdate") or ctx.get("pend", "?")
            is_sep = any("Separate" in d["member"] for d in dims)
            is_con = any("Consolidated" in d["member"] for d in dims)
            fs = "Sep" if is_sep else ("Con" if is_con else "?")
            refund_found.append((local, val, fs, len(dims), dims, pdate))

    if refund_found:
        seen = set()
        for local, val, fs, ndim, dims, pdate in refund_found[:15]:
            dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
            sig = (local, fs, ndim, dim_key, pdate)
            if sig in seen:
                continue
            seen.add(sig)
            print(f"    tag={local}, val={val}, {fs}, n_dim={ndim}, date={pdate}")
            for d in dims[:4]:
                print(f"      {d['axis']} = {d['member']}")
    else:
        # dart 전용 태그에서 탐색
        print("    (surrender 계열 없음. dart 태그 탐색...)")
        for elem in root:
            ns_uri = get_ns(elem.tag)
            local = get_local(elem.tag)
            if "dart" not in ns_uri.lower():
                continue
            low = local.lower()
            if any(k in low for k in ["reserve", "refund", "liability", "provision"]):
                ctx_ref = elem.get("contextRef", "")
                if not elem.text:
                    continue
                try:
                    val = float(elem.text.strip())
                except ValueError:
                    continue
                if abs(val) < 1000:
                    continue
                ctx = contexts.get(ctx_ref, {})
                dims = ctx.get("dims", [])
                pdate = ctx.get("pdate") or ctx.get("pend", "?")
                is_sep = any("Separate" in d["member"] for d in dims)
                is_con = any("Consolidated" in d["member"] for d in dims)
                fs = "Sep" if is_sep else ("Con" if is_con else "?")
                print(f"    tag={local}, val={val}, {fs}, n_dim={len(dims)}, date={pdate}")
                for d in dims[:3]:
                    print(f"      {d['axis']} = {d['member']}")


def main():
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])


if __name__ == "__main__":
    main()
