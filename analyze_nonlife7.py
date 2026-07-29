"""
손보사 XBRL BEL/RA/CSM 탐색
- 정답은 백만원 단위
- XBRL 단위: KRW(원), decimals=-6 이면 raw/1e6 -> 백만원
- InsuranceContractsLiabilityAsset + InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis로 분해된 값의 합계
- 또는 dart 전용 태그로 직접 있을 가능성
"""
import xml.etree.ElementTree as ET
import sys

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
                        "axis_full": dim_full,
                        "member_full": val_full,
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


def raw_to_mn(raw, decimals):
    """XBRL raw + decimals -> 백만원"""
    try:
        dec = int(decimals)
    except (ValueError, TypeError):
        return raw / 1e6  # 기본 가정: 원 단위
    # actual_value = raw * 10^(-decimals) -> 원
    # 백만원 = actual / 1e6
    actual = raw * (10 ** (-dec))
    return actual / 1e6


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")
    sys.stdout.flush()

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    sep_2025 = {}
    all_2025_instant = {}
    for cid, c in contexts.items():
        if c["ptype"] == "instant" and c["pdate"] == "2025-12-31":
            all_2025_instant[cid] = c
            if any("Separate" in d["member"] for d in c["dims"]):
                sep_2025[cid] = c

    print(f"  Sep instant 2025-12-31: {len(sep_2025)}개")

    # BEL 정답과 비교: 78736684 백만원 = 78.7조
    # raw_to_mn 변환 기준:
    #   decimals=-6: actual = raw * 10^6 = raw * 1e6 원, 백만원 = raw * 1e6 / 1e6 = raw
    #   decimals=0: actual = raw 원, 백만원 = raw / 1e6
    # 즉 decimals=-6 이면 XBRL raw값이 직접 백만원!

    print(f"\n  [decimals=-6 기준: raw값=백만원]")
    print(f"  BEL 정답: {targets['BEL']}백만원")
    print(f"\n  --- InsuranceContracts 계열 + Sep + 2025-12-31 instant ---")

    ic_candidates = []
    for elem in root:
        local = get_local(elem.tag)
        # 보험부채 관련 태그
        if not any(k in local for k in ["InsuranceContract", "InsurancContract"]):
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
        dec = elem.get("decimals", "?")
        val_mn = raw_to_mn(raw, dec)
        ctx = contexts[ctx_ref]
        dims = ctx["dims"]
        dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
        ic_candidates.append((abs(val_mn), local, val_mn, raw, dec, len(dims), dims, dim_key))

    # 중복 제거
    seen = set()
    unique = []
    for item in ic_candidates:
        sig = (item[1], item[6])
        if sig not in seen:
            seen.add(sig)
            unique.append(item)
    unique.sort(reverse=True)

    for _, local, val_mn, raw, dec, ndim, dims, dim_key in unique[:30]:
        tgt = targets.get("BEL", 0)
        marker = " <-- BEL?" if tgt > 0 and abs(val_mn - tgt) / tgt < 0.05 else ""
        # dims 요약
        bycovincl = [d for d in dims if "InsuranceContractsByRemainingCoverage" in d["axis"]]
        bycompon = [d for d in dims if "InsuranceContractsByComponents" in d["axis"]]
        if bycovincl:
            print(f"  tag={local[:60]}")
            print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={dec}, n_dim={ndim}{marker}")
            for d in dims[:5]:
                print(f"    {d['axis'][:70]} = {d['member']}")
        elif bycompon:
            print(f"  tag={local[:60]}")
            print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={dec}, n_dim={ndim}{marker}")
            for d in dims[:5]:
                print(f"    {d['axis'][:70]} = {d['member']}")
        elif ndim <= 2:
            print(f"  tag={local[:60]}")
            print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={dec}, n_dim={ndim}{marker}")
            for d in dims[:3]:
                print(f"    {d['axis'][:70]} = {d['member']}")

    # BEL/RA/CSM component 분해 탐색
    print(f"\n  --- InsuranceContractsByRemainingCoverageAxis + EstimatesPresentValue(BEL) ---")
    bel_component = []
    ra_component = []
    csm_component = []

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

        # BEL: InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis + EstimatesOfPresentValueOfFutureCashFlowsMember
        # RA: InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis + RiskAdjustmentForNonfinancialRiskMember
        # CSM: InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis + ContractualServiceMarginMember
        bycov_members = [d["member"] for d in dims if "InsuranceContractsByRemainingCoverage" in d["axis"]]
        bycompon_members = [d["member"] for d in dims if "InsuranceContractsByComponents" in d["axis"]]

        if "EstimatesOfPresentValueOfFutureCashFlows" in "".join(bycov_members):
            bel_component.append((val_mn, local, raw, dec, len(dims), dims))
        elif "RiskAdjustmentForNonfinancialRisk" in "".join(bycov_members):
            ra_component.append((val_mn, local, raw, dec, len(dims), dims))
        elif "ContractualServiceMargin" in "".join(bycov_members):
            csm_component.append((val_mn, local, raw, dec, len(dims), dims))

        # InsuranceContractsByComponentsAxis 기반도 확인
        if "EstimatesOfPresentValueOfFutureCashFlows" in "".join(bycompon_members):
            bel_component.append((val_mn, local, raw, dec, len(dims), dims))
        elif "RiskAdjustmentForNonfinancialRisk" in "".join(bycompon_members):
            ra_component.append((val_mn, local, raw, dec, len(dims), dims))
        elif "ContractualServiceMarginMember" in "".join(bycompon_members):
            csm_component.append((val_mn, local, raw, dec, len(dims), dims))

    def show_component(items, label, tgt):
        print(f"\n  {label} 정답={tgt}:")
        seen_s = set()
        items.sort(key=lambda x: abs(x[0] - tgt))
        for val_mn, local, raw, dec, ndim, dims in items[:10]:
            dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
            sig = (local, dim_key)
            if sig in seen_s: continue
            seen_s.add(sig)
            err = abs(val_mn - tgt) / abs(tgt) if tgt else abs(val_mn)
            print(f"    err={err:.4f} | tag={local[:70]}")
            print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={dec}, n_dim={ndim}")
            for d in dims[:4]:
                print(f"      {d['axis'][:70]} = {d['member']}")

    show_component(bel_component, "BEL", targets["BEL"])
    show_component(ra_component, "RA", targets["RA"])
    show_component(csm_component, "CSM", targets["CSM"])

    # InsuranceContractsByComponentsAxis / InsuranceContractsByRemainingCoverageAxis 전체 members 파악
    print(f"\n  [Sep 2025-12-31 context의 ByComponents/ByRemainingCoverage axis members]")
    axis_members = {}
    for c in sep_2025.values():
        for d in c["dims"]:
            if "InsuranceContractsByComponents" in d["axis"] or "InsuranceContractsByRemainingCoverage" in d["axis"]:
                axis_members.setdefault(d["axis"], set()).add(d["member"])
    for ax, members in axis_members.items():
        print(f"    {ax}")
        for m in members:
            print(f"      -> {m}")


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
