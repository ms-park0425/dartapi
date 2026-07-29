"""
손보사 XBRL BEL/RA/CSM 탐색 v8
- decimals=-6: raw값이 백만원 단위 (raw * 10^6 원 = raw 백만원)
- BEL/RA/CSM은 InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis의
  EstimatesOfPresentValueOfFutureCashFlowsMember / RiskAdjustmentForNonfinancialRiskMember / ContractualServiceMarginMember
  또는 InsuranceContractsByComponentsAxis
"""
import xml.etree.ElementTree as ET
import sys

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"
sys.stdout.reconfigure(encoding="utf-8")


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
        actual = raw * (10 ** (-dec))  # 원
        return actual / 1e6  # 백만원
    except (ValueError, TypeError):
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

    print(f"  Sep instant 2025-12-31: {len(sep_2025)}")

    # BEL/RA/CSM 컴포넌트 axis 멤버 파악
    comp_axes = {}
    for cid in sep_2025:
        c = contexts[cid]
        for ax, mem in c["dims"]:
            if "ByRemainingCoverage" in ax or "ByComponents" in ax or "ByRemainingCoverageAndIncurred" in ax:
                comp_axes.setdefault(ax, set()).add(mem)

    print(f"\n  [BEL/RA/CSM 관련 axis/member]")
    for ax, mems in comp_axes.items():
        print(f"    axis: {ax}")
        for m in sorted(mems):
            print(f"      {m}")

    # 각 component member 기반 탐색
    BEL_MEM_KEYWORDS = ["EstimatesOfPresentValueOfFutureCashFlows", "FulfilmentCashFlow", "BestEstimate", "PresentValue"]
    RA_MEM_KEYWORDS = ["RiskAdjustmentForNonfinancial", "RiskAdjustment"]
    CSM_MEM_KEYWORDS = ["ContractualServiceMargin"]

    def has_member(dims, keywords):
        return any(any(k in mem for k in keywords) for _, mem in dims)

    def is_sep(dims):
        return any("Separate" in mem for _, mem in dims)

    def is_con(dims):
        return any("Consolidated" in mem for _, mem in dims)

    # 삼성화재 특정 ByComponents axis 확인
    print(f"\n  [all Sep instant 2025-12-31 BEL component hits]")
    bel_hits = []
    ra_hits = []
    csm_hits = []

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
        dec = elem.get("decimals", "?")
        val_mn = raw_to_mn(raw, dec)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        local = get_local(elem.tag)

        if has_member(dims, BEL_MEM_KEYWORDS):
            bel_hits.append((val_mn, local, raw, dec, len(dims), dims))
        if has_member(dims, RA_MEM_KEYWORDS):
            ra_hits.append((val_mn, local, raw, dec, len(dims), dims))
        if has_member(dims, CSM_MEM_KEYWORDS):
            csm_hits.append((val_mn, local, raw, dec, len(dims), dims))

    def show_hits(hits, label, tgt):
        print(f"\n  {label} (정답={tgt}백만원):")
        seen_s = set()
        hits_sorted = sorted(hits, key=lambda x: abs(x[0] - tgt))
        for val_mn, local, raw, dec, ndim, dims in hits_sorted[:10]:
            sig = (local, dims)
            if sig in seen_s: continue
            seen_s.add(sig)
            err = abs(val_mn - tgt) / abs(tgt) if tgt else 0
            marker = " *** MATCH ***" if err < 0.001 else (f" err={err:.4f}" if err < 0.05 else "")
            print(f"    tag={local[:70]}{marker}")
            print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={dec}, n_dim={ndim}")
            for ax, mem in dims[:5]:
                print(f"      {ax[:60]} = {mem}")

    show_hits(bel_hits, "BEL", targets["BEL"])
    show_hits(ra_hits, "RA", targets["RA"])
    show_hits(csm_hits, "CSM", targets["CSM"])

    # Context id 패턴 분석 (ByRemainingCoverage + Separate)
    print(f"\n  [Sep 2025-12-31 InsuranceContractsLiabilityAsset 직접 탐색]")
    for elem in root:
        local = get_local(elem.tag)
        if "InsuranceContractsLiabilityAsset" not in local and "InsuranceContractsThatAreLiabilities" not in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in all_2025_instant:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        dec = elem.get("decimals", "?")
        val_mn = raw_to_mn(raw, dec)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        sep = is_sep(dims)
        con = is_con(dims)
        fs = "Sep" if sep else ("Con" if con else "?")
        print(f"    tag={local[:70]}, val_mn={val_mn:.0f}, {fs}, n_dim={len(dims)}, dec={dec}")
        for ax, mem in dims[:5]:
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
}


if __name__ == "__main__":
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])
