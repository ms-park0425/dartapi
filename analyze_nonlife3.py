"""
손보사 XBRL - 단위/스케일 파악 및 정답값 탐색 (스케일 고려)
"""
import xml.etree.ElementTree as ET
from collections import Counter

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


def check_decimals(xbrl_path):
    """XBRL 파일에서 decimals 속성 분포 확인"""
    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    dec_counter = Counter()
    for elem in root:
        dec = elem.get("decimals")
        if dec is not None:
            dec_counter[dec] += 1
    return dec_counter


def search_values_scaled(xbrl_path, targets, scale=1000000):
    """targets (백만원) * scale 을 XBRL에서 탐색"""
    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    # instant 2025-12-31
    instant_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"}
    # duration end 2025-12-31
    dur_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "duration" and c["pend"] == "2025-12-31"}

    # 값 인덱스 구축 (raw float)
    val_index = {}
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            v = float(elem.text.strip())
        except ValueError:
            continue
        local = get_local(elem.tag)
        val_index.setdefault(v, []).append((local, ctx_ref, elem.get("decimals", "?")))

    results = {}
    for label, tgt_mn in targets.items():
        tgt_raw = tgt_mn * scale  # 백만원 -> 원
        # 정확 매칭 (float)
        hits = []
        # scale variants 시도
        for s in [1, 1000, 1000000]:
            tgt_try = tgt_mn * s
            # 근사: ±0.5 (정수 반올림 오차)
            for delta in [0, 0.5, -0.5, 1, -1, 1000, -1000]:
                v_try = tgt_try + delta
                if v_try in val_index:
                    for local, ctx_ref, decimals in val_index[v_try]:
                        ctx = contexts.get(ctx_ref, {})
                        dims = ctx.get("dims", [])
                        is_sep = any("Separate" in d["member"] for d in dims)
                        is_con = any("Consolidated" in d["member"] for d in dims)
                        ptype = ctx.get("ptype", "?")
                        pdate = ctx.get("pdate") or ctx.get("pend", "?")
                        fs = "Sep" if is_sep else ("Con" if is_con else "?")
                        hits.append({
                            "scale": s, "delta": delta,
                            "tag": local, "val": v_try, "decimals": decimals,
                            "fs": fs, "ptype": ptype, "pdate": pdate,
                            "ndim": len(dims), "dims": dims, "ctx_ref": ctx_ref
                        })
                    break
            if hits:
                break
        results[label] = hits
    return results, contexts, instant_2025, dur_2025


def main():
    for company, info in COMPANIES.items():
        print(f"\n{'='*70}")
        print(f"[{company}]")
        print(f"{'='*70}")

        # decimals 분포 확인
        dec_dist = check_decimals(info["path"])
        print(f"  decimals 분포: {dict(dec_dist.most_common(5))}")

        results, contexts, inst_2025, dur_2025 = search_values_scaled(
            info["path"], info["targets"]
        )

        for label, hits in results.items():
            tgt = info["targets"][label]
            print(f"\n  [{label}={tgt}]")
            if not hits:
                print(f"    -> 모든 scale (1/1000/1000000) 에서 매칭 없음")
                # BEL/RA/CSM 키워드로 직접 탐색
                tree = ET.parse(info["path"])
                root = tree.getroot()
                ctxs = parse_contexts(root)
                bel_kws = ["BestEstimate", "PresentValue", "FulfilmentCashFlow"]
                ra_kws = ["RiskAdjustment"]
                csm_kws = ["ContractualServiceMargin"]
                oci_kws = ["OtherComprehensiveIncome", "AccumulatedOCI", "ComponentsOfEquity",
                           "RevaluationSurplus", "FairValueReserve", "OtherReserves"]

                kws = []
                if "BEL" in label: kws = bel_kws
                elif "RA" in label: kws = ra_kws
                elif "CSM" in label: kws = csm_kws
                elif "OCI" in label: kws = oci_kws

                found_kw = []
                for elem in root:
                    local = get_local(elem.tag)
                    if not any(k.lower() in local.lower() for k in kws):
                        continue
                    ctx_ref = elem.get("contextRef", "")
                    if ctx_ref not in inst_2025 and ctx_ref not in dur_2025:
                        continue
                    if not elem.text:
                        continue
                    try:
                        v = float(elem.text.strip())
                    except ValueError:
                        continue
                    ctx = ctxs.get(ctx_ref, {})
                    dims = ctx.get("dims", [])
                    is_sep = any("Separate" in d["member"] for d in dims)
                    is_con = any("Consolidated" in d["member"] for d in dims)
                    fs = "Sep" if is_sep else ("Con" if is_con else "?")
                    pdate = ctx.get("pdate") or ctx.get("pend", "?")
                    found_kw.append((local, v, fs, len(dims), dims, pdate, elem.get("decimals", "?")))

                if found_kw:
                    seen = set()
                    for local, v, fs, ndim, dims, pdate, dec in found_kw[:20]:
                        dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
                        sig = (local, fs, ndim, dim_key, pdate)
                        if sig in seen: continue
                        seen.add(sig)
                        print(f"    keyword_tag={local}, val={v}, dec={dec}, {fs}, n_dim={ndim}, date={pdate}")
                        for d in dims[:4]:
                            print(f"      {d['axis']} = {d['member']}")
                else:
                    print(f"    keyword 탐색도 결과 없음")
            else:
                seen = set()
                for h in hits[:5]:
                    dims = h["dims"]
                    dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
                    sig = (h["tag"], h["fs"], h["ndim"], dim_key, h["pdate"])
                    if sig in seen: continue
                    seen.add(sig)
                    print(f"    tag={h['tag']}, val={h['val']}, dec={h['decimals']}, scale={h['scale']}, {h['fs']}, n_dim={h['ndim']}, ptype={h['ptype']}, date={h['pdate']}")
                    for d in dims[:5]:
                        print(f"      {d['axis']} = {d['member']}")


if __name__ == "__main__":
    main()
