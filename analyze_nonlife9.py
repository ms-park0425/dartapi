"""
손보사 XBRL - BEL/RA/CSM 집계 태그 탐색
TheCarryingAmountOfInsuranceContractLiabilities 또는 dart 전용 태그를 찾음
"""
import xml.etree.ElementTree as ET
import sys
sys.stdout.reconfigure(encoding="utf-8")

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"


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
                    dims.append((get_local(dim_full), get_local(val_full)))
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        ptype, pdate = None, None
        if period is not None:
            inst = period.find(f"{{{NS_XBRLI}}}instant")
            en = period.find(f"{{{NS_XBRLI}}}endDate")
            if inst is not None:
                ptype, pdate = "instant", inst.text
            elif en is not None:
                ptype, pdate = "duration", en.text
        contexts[cid] = {"id": cid, "dims": tuple(dims), "ptype": ptype, "pdate": pdate}
    return contexts


def raw_to_mn(raw, decimals):
    try:
        dec = int(decimals)
        actual = raw * (10 ** (-dec))
        return actual / 1e6
    except:
        return raw / 1e6


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    sep_2025 = {cid for cid, c in contexts.items()
                if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"
                and any("Separate" in m for _, m in c["dims"])}

    all_2025_instant = {cid for cid, c in contexts.items()
                        if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"}

    # 1. dart 전용 태그 탐색 (BEL/RA/CSM 관련)
    print(f"\n  [dart 네임스페이스 BEL/RA/CSM 관련 태그 (Sep 2025-12-31)]")
    for elem in root:
        ns = get_ns(elem.tag)
        local = get_local(elem.tag)
        if "dart" not in ns.lower():
            continue
        if not any(k in local for k in ["BEL", "BestEstimate", "FulfilmentCash", "RiskAdj", "CSM", "ContractualService",
                                         "InsuranceLiability", "InsuranceContractLiab", "Liability"]):
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in all_2025_instant:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except:
            continue
        dec = elem.get("decimals", "?")
        val_mn = raw_to_mn(raw, dec)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        sep = any("Separate" in m for _, m in dims)
        fs = "Sep" if sep else "Con"
        print(f"    ns={ns[-30:]}, tag={local[:60]}")
        print(f"      val_mn={val_mn:.0f}, {fs}, n_dim={len(dims)}")
        for ax, mem in dims[:3]:
            print(f"      {ax[:50]} = {mem}")

    # 2. TheCarryingAmount 류 태그 탐색
    print(f"\n  [TheCarryingAmount / BookValue 류 태그 (Sep 2025-12-31)]")
    for elem in root:
        local = get_local(elem.tag)
        if not any(k in local for k in ["CarryingAmount", "BookValue", "TheBookValue", "TheCarrying"]):
            continue
        if "Insurance" not in local and "Liability" not in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in all_2025_instant:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except:
            continue
        dec = elem.get("decimals", "?")
        val_mn = raw_to_mn(raw, dec)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        sep = any("Separate" in m for _, m in dims)
        fs = "Sep" if sep else "Con"
        print(f"    tag={local[:80]}")
        print(f"      val_mn={val_mn:.0f}, {fs}, n_dim={len(dims)}")
        for ax, mem in dims[:4]:
            print(f"      {ax[:60]} = {mem}")

    # 3. 정답과 근사한 값을 모든 태그에서 탐색 (Sep only, all instant 2025)
    print(f"\n  [정답값 근사 탐색 (전체 태그, Sep 2025-12-31 instant, ±2%)]")
    for label, tgt_mn in targets.items():
        if "OCI" in label:
            continue
        print(f"\n  {label} = {tgt_mn}백만원:")
        hits = []
        for elem in root:
            ctx_ref = elem.get("contextRef", "")
            if ctx_ref not in sep_2025:
                continue
            if not elem.text:
                continue
            try:
                raw = float(elem.text.strip())
            except:
                continue
            dec = elem.get("decimals", "?")
            val_mn = raw_to_mn(raw, dec)
            if tgt_mn != 0:
                err = abs(val_mn - tgt_mn) / abs(tgt_mn)
            else:
                err = abs(val_mn)
            if err < 0.02:
                local = get_local(elem.tag)
                ctx = contexts[ctx_ref]
                dims = ctx["dims"]
                hits.append((err, local, val_mn, raw, dec, len(dims), dims))

        hits.sort(key=lambda x: x[0])
        seen_sig = set()
        for err, local, val_mn, raw, dec, ndim, dims in hits[:10]:
            sig = (local, dims)
            if sig in seen_sig: continue
            seen_sig.add(sig)
            print(f"    err={err:.4f} | {local[:70]}")
            print(f"      val_mn={val_mn:.0f} | raw={raw} | dec={dec} | n_dim={ndim}")
            for ax, mem in dims[:4]:
                print(f"      {ax[:60]} = {mem}")


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
    "현대해상": {
        "path": "C:/MALIFE/dartapi/data/현대해상/사업보고서/2025/entity00164973_2025-12-31.xbrl",
        "targets": {"BEL": 43022017, "RA": 873617, "CSM": 4490254}
    },
    "DB손해보험": {
        "path": "C:/MALIFE/dartapi/data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl",
        "targets": {"BEL": 52738440, "RA": 1058497, "CSM": 3983500}
    },
}


if __name__ == "__main__":
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])
