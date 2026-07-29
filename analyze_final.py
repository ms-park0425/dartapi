"""
손보사 XBRL - BEL/RA/CSM 최종 탐색
핵심: InsuranceContractsLiabilityAsset + ByComponents axis 합산
단위: raw = 원, /1e6 = 백만원
"""
import xml.etree.ElementTree as ET
import sys
sys.stdout.reconfigure(encoding="utf-8")

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


def val_to_mn(raw):
    """raw원 -> 백만원"""
    return raw / 1_000_000


def has_dim(dims, axis_kw, member_kw):
    return any(axis_kw in ax and member_kw in mem for ax, mem in dims)


def has_only_sep_and_one(dims):
    """dims에 SeparateMember가 있고, FS axis 외에 1개 axis만 있는 경우"""
    non_fs = [(ax, mem) for ax, mem in dims if "ConsolidatedAndSeparate" not in ax]
    return len(non_fs) == 1


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

    all_2025 = {cid for cid, c in contexts.items()
                if c["pdate"] == "2025-12-31"}

    print(f"  Sep instant 2025-12-31: {len(sep_2025)}")

    # 방법1: InsuranceContractsLiabilityAsset + ByComponents axis 합산
    # 각 component: EstimatesOfPresentValueOfFutureCashFlowsMember, RiskAdjustment..., ContractualServiceMargin
    bel_sum = ra_sum = csm_sum = 0.0
    bel_items = []
    ra_items = []
    csm_items = []

    # 집계 후보 태그
    ic_tags = ["InsuranceContractsLiabilityAsset", "InsuranceContractsThatAreLiabilities",
               "InsuranceContractsIssuedThatAreLiabilities"]

    for elem in root:
        local = get_local(elem.tag)
        if local not in ic_tags:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in sep_2025:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except:
            continue

        val_mn = val_to_mn(raw)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]

        # ByComponents axis
        for ax, mem in dims:
            if "InsuranceContractsByComponents" in ax or "InsuranceContractsByRemainingCoverage" in ax:
                if "EstimatesOfPresentValueOfFutureCashFlows" in mem:
                    bel_items.append((val_mn, local, raw, len(dims), dims))
                elif "RiskAdjustmentForNonfinancial" in mem or ("RiskAdjustment" in mem and "RiskAdjustmentForNonfinancial" in mem):
                    ra_items.append((val_mn, local, raw, len(dims), dims))
                elif "ContractualServiceMargin" in mem and "Not" not in mem and "Related" not in mem:
                    csm_items.append((val_mn, local, raw, len(dims), dims))

    # ByComponents=EstimatesOfPresent... 에서 가장 간단한 context(dim 수 최소) 선택
    # 단 보험종류별 합산(InsuranceContractsIssued) 포함
    print(f"\n  [ByComponents=EstimatesOfPresentValue 탐색 ({len(bel_items)}건)]")
    # 중복 제거 (같은 dim 패턴)
    seen_sig = set()
    bel_sum_attempt = 0.0
    for val_mn, local, raw, ndim, dims in sorted(bel_items, key=lambda x: x[3]):
        sig = dims
        if sig in seen_sig: continue
        seen_sig.add(sig)
        has_issued = any("InsuranceContractsIssuedMember" in m for _, m in dims)
        has_held = any("ReinsuranceContractsHeld" in m for _, m in dims)
        print(f"    val_mn={val_mn:.0f}, n_dim={ndim}, issued={has_issued}, held={has_held}")
        for ax, mem in dims:
            print(f"      {ax[:60]}={mem}")
        if has_issued and not has_held:
            bel_sum_attempt += val_mn

    print(f"\n  [BEL component 합산 시도 (InsuranceContractsIssuedMember)] = {bel_sum_attempt:.0f}백만원")
    print(f"  정답 BEL: {targets['BEL']}")

    # 방법2: RA component
    print(f"\n  [ByComponents=RiskAdjustment 탐색 ({len(ra_items)}건)]")
    seen_sig = set()
    for val_mn, local, raw, ndim, dims in sorted(ra_items, key=lambda x: x[3]):
        sig = dims
        if sig in seen_sig: continue
        seen_sig.add(sig)
        has_issued = any("InsuranceContractsIssuedMember" in m for _, m in dims)
        has_held = any("ReinsuranceContractsHeld" in m for _, m in dims)
        print(f"    val_mn={val_mn:.0f}, n_dim={ndim}, issued={has_issued}, held={has_held}")
        for ax, mem in dims:
            print(f"      {ax[:60]}={mem}")

    # 방법3: CSM component
    print(f"\n  [ByComponents=ContractualServiceMargin 탐색 ({len(csm_items)}건)]")
    seen_sig = set()
    for val_mn, local, raw, ndim, dims in sorted(csm_items, key=lambda x: x[3]):
        sig = dims
        if sig in seen_sig: continue
        seen_sig.add(sig)
        has_issued = any("InsuranceContractsIssuedMember" in m for _, m in dims)
        has_held = any("ReinsuranceContractsHeld" in m for _, m in dims)
        print(f"    val_mn={val_mn:.0f}, n_dim={ndim}, issued={has_issued}, held={has_held}")
        for ax, mem in dims:
            print(f"      {ax[:60]}={mem}")

    # 방법4: 직접 RiskAdjustmentForNonFinancialRisk 태그 탐색 (dart 전용)
    print(f"\n  [dart:RiskAdjustmentForNonFinancialRiskAboutLiabilitiesForRemainingCoverage 탐색]")
    for elem in root:
        local = get_local(elem.tag)
        if "RiskAdjustmentForNonFinancialRisk" not in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in sep_2025:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except:
            continue
        val_mn = val_to_mn(raw)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        print(f"    tag={local[:60]}, val_mn={val_mn:.0f}, n_dim={len(dims)}")
        for ax, mem in dims:
            print(f"      {ax[:55]}={mem}")

    # 방법5: dart CSM 전용 태그 탐색
    print(f"\n  [ContractualServiceMargin 태그 (Sep 2025-12-31)]")
    csm_direct = []
    for elem in root:
        local = get_local(elem.tag)
        if "ContractualServiceMargin" not in local:
            continue
        if "Increase" in local or "Decrease" in local or "Transfer" in local or "Basis" in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in sep_2025:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except:
            continue
        val_mn = val_to_mn(raw)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        csm_direct.append((len(dims), val_mn, local, raw, dims))

    for ndim, val_mn, local, raw, dims in sorted(csm_direct)[:10]:
        print(f"    tag={local[:70]}, val_mn={val_mn:.0f}, n_dim={ndim}, raw={raw}")
        for ax, mem in dims:
            print(f"      {ax[:55]}={mem}")


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
