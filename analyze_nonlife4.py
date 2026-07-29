"""
손보사 XBRL - BEL/RA/CSM/OCI 정답값 탐색
단위: XBRL decimals=-6 → 백만원 단위 저장, 정답값도 백만원
"""
import xml.etree.ElementTree as ET
from collections import Counter

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"

COMPANIES = {
    "삼성화재": {
        "path": "C:/MALIFE/dartapi/data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
        "targets": {
            "BEL": 78736684,
            "RA": 1616009,
            "CSM": 8897609,
            "OCI25": 3944864,
        }
    },
    "현대해상": {
        "path": "C:/MALIFE/dartapi/data/현대해상/사업보고서/2025/entity00164973_2025-12-31.xbrl",
        "targets": {
            "BEL": 43022017,
            "RA": 873617,
            "CSM": 4490254,
            "OCI25": -1060623,
        }
    },
    "메리츠화재": {
        "path": "C:/MALIFE/dartapi/data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
        "targets": {
            "BEL": 38011337,
            "RA": 817895,
            "CSM": 3413560,
            "OCI25": -876506,
        }
    },
    "KB손해보험": {
        "path": "C:/MALIFE/dartapi/data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
        "targets": {
            "BEL": 39028451,
            "RA": 812040,
            "CSM": 3402947,
            "OCI25": -827590,
        }
    },
    "DB손해보험": {
        "path": "C:/MALIFE/dartapi/data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl",
        "targets": {
            "BEL": 52738440,
            "RA": 1058497,
            "CSM": 3983500,
            "OCI25": -389340,
        }
    },
}


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
        ptype, pdate, pend = None, None, None
        if period is not None:
            inst = period.find(f"{{{NS_XBRLI}}}instant")
            en = period.find(f"{{{NS_XBRLI}}}endDate")
            if inst is not None:
                ptype, pdate = "instant", inst.text
            elif en is not None:
                ptype, pdate = "duration", en.text
        contexts[cid] = {"id": cid, "dims": dims, "ptype": ptype, "pdate": pdate}
    return contexts


def to_mn(raw_val, decimals_str):
    """raw_val(원 단위 등) + decimals -> 백만원 변환"""
    try:
        dec = int(decimals_str)
    except (ValueError, TypeError):
        return None
    # XBRL: value * 10^(-decimals) = actual
    # decimals=-6 -> actual = raw * 10^6 (백만원 단위)
    actual = raw_val * (10 ** (-dec)) if dec != 0 else raw_val
    return actual / 1_000_000  # -> 백만원


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    # decimals 분포
    dec_dist = Counter()
    for elem in root:
        d = elem.get("decimals")
        if d: dec_dist[d] += 1
    print(f"  decimals: {dict(dec_dist.most_common(3))}")

    # instant 2025-12-31 and Separate context ids
    sep_ctx = set()
    all_2025_ctx = set()
    for cid, c in contexts.items():
        if c["pdate"] == "2025-12-31":
            all_2025_ctx.add(cid)
            if any("Separate" in d["member"] for d in c["dims"]):
                sep_ctx.add(cid)

    print(f"  2025-12-31 contexts: {len(all_2025_ctx)} (Separate: {len(sep_ctx)})")

    # BEL 관련 태그 키워드 매핑
    kw_map = {
        "BEL": ["BestEstimate", "FulfilmentCashFlow", "PresentValueOfFutureCashFlow",
                "LiabilityForRemainingCoverage", "LiabilityForIncurredClaims",
                "InsuranceContractLiability", "InsuranceContractsThatAreLiabilities",
                "NetEstimatedPresentValue"],
        "RA":  ["RiskAdjustment"],
        "CSM": ["ContractualServiceMargin"],
        "OCI25": ["OtherComprehensiveIncome", "AccumulatedOtherComprehensive",
                  "ComponentsOfEquity", "EquityAttributable", "ReserveOf",
                  "OtherReserves", "FairValueGainLoss"],
    }

    for label, tgt_mn in targets.items():
        print(f"\n  [{label}={tgt_mn}백만원]")
        kws = kw_map.get(label, [])
        best = []

        for elem in root:
            local = get_local(elem.tag)
            if not any(k.lower() in local.lower() for k in kws):
                continue
            ctx_ref = elem.get("contextRef", "")
            if ctx_ref not in all_2025_ctx:
                continue
            if not elem.text:
                continue
            try:
                raw = float(elem.text.strip())
            except ValueError:
                continue

            decimals = elem.get("decimals", "?")
            # 백만원 변환
            val_mn = to_mn(raw, decimals)
            if val_mn is None:
                val_mn = raw  # 변환 불가시 raw 사용

            # 정답과 오차 계산
            if tgt_mn != 0:
                err_pct = abs(val_mn - tgt_mn) / abs(tgt_mn)
            else:
                err_pct = abs(val_mn)

            ctx = contexts.get(ctx_ref, {})
            dims = ctx.get("dims", [])
            is_sep = any("Separate" in d["member"] for d in dims)
            is_con = any("Consolidated" in d["member"] for d in dims)
            fs = "Sep" if is_sep else ("Con" if is_con else "?")

            if err_pct < 0.005:  # 0.5% 이내
                best.append((err_pct, local, val_mn, raw, decimals, fs, len(dims), dims, ctx_ref))

        if not best:
            print(f"    -> 0.5% 이내 매칭 없음. 10% 이내로 확대...")
            for elem in root:
                local = get_local(elem.tag)
                if not any(k.lower() in local.lower() for k in kws):
                    continue
                ctx_ref = elem.get("contextRef", "")
                if ctx_ref not in all_2025_ctx:
                    continue
                if not elem.text:
                    continue
                try:
                    raw = float(elem.text.strip())
                except ValueError:
                    continue
                decimals = elem.get("decimals", "?")
                val_mn = to_mn(raw, decimals)
                if val_mn is None:
                    val_mn = raw
                if tgt_mn != 0:
                    err_pct = abs(val_mn - tgt_mn) / abs(tgt_mn)
                else:
                    err_pct = abs(val_mn)
                if err_pct < 0.1:
                    ctx = contexts.get(ctx_ref, {})
                    dims = ctx.get("dims", [])
                    is_sep = any("Separate" in d["member"] for d in dims)
                    is_con = any("Consolidated" in d["member"] for d in dims)
                    fs = "Sep" if is_sep else ("Con" if is_con else "?")
                    best.append((err_pct, local, val_mn, raw, decimals, fs, len(dims), dims, ctx_ref))
            if not best:
                print(f"    -> 10% 이내도 없음")
                # 마지막 수단: 모든 2025-12-31 Sep context에서 keyword 없이 raw 값 비교
                print(f"    (Sep 2025-12-31 context에서 근사값 탐색...)")
                for elem in root:
                    ctx_ref = elem.get("contextRef", "")
                    if ctx_ref not in sep_ctx:
                        continue
                    if not elem.text:
                        continue
                    try:
                        raw = float(elem.text.strip())
                    except ValueError:
                        continue
                    decimals = elem.get("decimals", "?")
                    val_mn = to_mn(raw, decimals)
                    if val_mn is None:
                        val_mn = raw
                    if tgt_mn != 0:
                        err_pct = abs(val_mn - tgt_mn) / abs(tgt_mn)
                    else:
                        err_pct = abs(val_mn)
                    if err_pct < 0.001:  # 0.1%
                        local = get_local(elem.tag)
                        ctx = contexts.get(ctx_ref, {})
                        dims = ctx.get("dims", [])
                        is_sep = any("Separate" in d["member"] for d in dims)
                        fs = "Sep" if is_sep else "?"
                        print(f"    MATCH tag={local}, val_mn={val_mn:.0f}, raw={raw}, dec={decimals}, {fs}, n_dim={len(dims)}")
                        for d in dims[:5]:
                            print(f"      {d['axis']} = {d['member']}")

        else:
            # 중복 제거 후 출력
            best.sort(key=lambda x: (x[0], x[6]))
            seen = set()
            for err, local, val_mn, raw, decimals, fs, ndim, dims, ctx_ref in best[:8]:
                dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
                sig = (local, fs, ndim, dim_key)
                if sig in seen: continue
                seen.add(sig)
                print(f"    err={err:.4f}, tag={local}")
                print(f"    val_mn={val_mn:.0f}, raw={raw}, dec={decimals}, {fs}, n_dim={ndim}")
                for d in dims[:5]:
                    print(f"      {d['axis']} = {d['member']}")

    # 해약환급금준비금 (Col99)
    print(f"\n  [해약환급금준비금 Col99]")
    refund_tags = ["SurrenderValueReserve", "AppropriationOfSurrenderValueReserve"]
    shown_refund = 0
    for elem in root:
        local = get_local(elem.tag)
        if not any(k.lower() in local.lower() for k in ["surrendervaluereserve"] ):
            continue
        if "ToBeAdded" in local or "Appropriation" in local or "Transfer" in local:
            continue
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        decimals = elem.get("decimals", "?")
        val_mn = to_mn(raw, decimals)
        if val_mn is None: val_mn = raw
        ctx = contexts.get(ctx_ref, {})
        dims = ctx.get("dims", [])
        pdate = ctx.get("pdate", "?")
        if pdate not in ["2025-12-31", "2024-12-31"]:
            continue
        is_sep = any("Separate" in d["member"] for d in dims)
        is_con = any("Consolidated" in d["member"] for d in dims)
        fs = "Sep" if is_sep else ("Con" if is_con else "?")
        print(f"    tag={local}, val_mn={val_mn:.0f}, raw={raw}, dec={decimals}, {fs}, n_dim={len(dims)}, date={pdate}")
        for d in dims[:3]:
            print(f"      {d['axis']} = {d['member']}")
        shown_refund += 1
        if shown_refund >= 6: break


def main():
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])


if __name__ == "__main__":
    main()
