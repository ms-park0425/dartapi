"""
손보사 XBRL - BEL/RA/CSM/OCI 정답값 탐색 v5
수정: to_mn = raw / 1_000_000 (원 -> 백만원)
"""
import xml.etree.ElementTree as ET

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
            "OCI26": 1001660,
            "OCI27": 3953,
            "OCI28": -217651,
            "OCI30": -4782,
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

BEL_TAGS = [
    "EstimatesOfPresentValueOfFutureCashFlows",
    "FulfilmentCashFlows",
    "BestEstimate",
    "PresentValueOfFutureCashFlow",
    "LiabilityForRemainingCoverage",
    "LiabilityForIncurredClaims",
    "InsuranceContractLiability",
    "InsuranceContractsThatAreLiabilities",
    "InsuranceContractsAssets",
    "NetInsuranceContractLiability",
]
RA_TAGS = ["RiskAdjustment"]
CSM_TAGS = ["ContractualServiceMargin"]
OCI_TAGS = [
    "OtherComprehensiveIncome",
    "AccumulatedOtherComprehensive",
    "ComponentsOfEquity",
    "ReserveOf",
    "FairValue",
    "UnrealisedGainLoss",
    "ValueChangeReserve",
]


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


def to_mn(raw_val):
    """원 단위 raw값 -> 백만원"""
    return raw_val / 1_000_000


def fmt_dims(dims):
    return [(d["axis"], d["member"]) for d in dims]


def search_tag_in_ctx(root, contexts, target_2025_ctx, kws, tgt_mn, tolerance=0.005):
    hits = []
    for elem in root:
        local = get_local(elem.tag)
        if not any(k.lower() in local.lower() for k in kws):
            continue
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in target_2025_ctx:
            continue
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        val_mn = to_mn(raw)
        if tgt_mn != 0:
            err = abs(val_mn - tgt_mn) / abs(tgt_mn)
        else:
            err = abs(val_mn)
        if err <= tolerance:
            ctx = contexts.get(ctx_ref, {})
            dims = ctx.get("dims", [])
            is_sep = any("Separate" in d["member"] for d in dims)
            is_con = any("Consolidated" in d["member"] for d in dims)
            fs = "Sep" if is_sep else ("Con" if is_con else "?")
            ptype = ctx.get("ptype", "?")
            hits.append({
                "err": err, "tag": local, "val_mn": val_mn, "raw": raw,
                "dec": elem.get("decimals", "?"), "fs": fs, "ptype": ptype,
                "ndim": len(dims), "dims": dims
            })
    return hits


def find_all_with_value(root, contexts, tgt_mn, tolerance=0.001):
    """전체 XBRL에서 tgt_mn과 근사한 값 탐색 (Sep 2025-12-31 우선)"""
    results = []
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        val_mn = to_mn(raw)
        if tgt_mn != 0:
            err = abs(val_mn - tgt_mn) / abs(tgt_mn)
        else:
            err = abs(val_mn - tgt_mn)
        if err <= tolerance:
            ctx = contexts.get(ctx_ref, {})
            dims = ctx.get("dims", [])
            is_sep = any("Separate" in d["member"] for d in dims)
            is_con = any("Consolidated" in d["member"] for d in dims)
            fs = "Sep" if is_sep else ("Con" if is_con else "?")
            results.append({
                "err": err, "tag": get_local(elem.tag), "val_mn": val_mn, "raw": raw,
                "dec": elem.get("decimals", "?"), "fs": fs,
                "ptype": ctx.get("ptype", "?"), "pdate": ctx.get("pdate", "?"),
                "ndim": len(dims), "dims": dims
            })
    results.sort(key=lambda x: (x["err"], 0 if x["fs"] == "Sep" else 1, x["ndim"]))
    return results


def analyze(company, xbrl_path, targets):
    print(f"\n{'='*70}")
    print(f"[{company}]")
    print(f"{'='*70}")

    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    contexts = parse_contexts(root)

    # 2025-12-31 context 분류
    instant_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "instant" and c["pdate"] == "2025-12-31"}
    dur_2025 = {cid for cid, c in contexts.items() if c["ptype"] == "duration" and c["pdate"] == "2025-12-31"}
    sep_2025 = {cid for cid in instant_2025 if any("Separate" in d["member"] for d in contexts[cid]["dims"])}
    all_2025 = instant_2025 | dur_2025

    print(f"  instant_2025: {len(instant_2025)} | dur_2025: {len(dur_2025)} | sep_instant: {len(sep_2025)}")

    for label, tgt_mn in targets.items():
        print(f"\n  [{label}={tgt_mn}백만원]")

        kws = {"BEL": BEL_TAGS, "RA": RA_TAGS, "CSM": CSM_TAGS}.get(label, OCI_TAGS)

        # 1단계: keyword 기반 탐색 (2025-12-31 전체)
        hits = search_tag_in_ctx(root, contexts, all_2025, kws, tgt_mn, tolerance=0.005)

        if not hits:
            # 2단계: keyword 없이 전체에서 정확 탐색
            hits = find_all_with_value(root, contexts, tgt_mn, tolerance=0.001)
            if hits:
                print(f"    [keyword 매칭 없음 -> 전체탐색 {len(hits)}건]")
            else:
                # 3단계: tolerance 확대 (1%)
                hits = find_all_with_value(root, contexts, tgt_mn, tolerance=0.01)
                if hits:
                    print(f"    [0.1% 없음 -> 1% 확대 {len(hits)}건]")
                else:
                    print(f"    -> 1% 이내 매칭 없음")
                    continue

        # 중복 제거 & 출력
        seen = set()
        for h in hits[:10]:
            dims = h["dims"]
            dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
            sig = (h["tag"], h["fs"], h["ndim"], dim_key)
            if sig in seen: continue
            seen.add(sig)
            print(f"    err={h['err']:.5f} | tag={h['tag']}")
            print(f"      val_mn={h['val_mn']:.0f} | raw={h['raw']} | dec={h['dec']}")
            print(f"      {h['fs']} | ptype={h['ptype']} | pdate={h.get('pdate','2025-12-31')} | n_dim={h['ndim']}")
            for d in dims[:5]:
                print(f"        {d['axis']} = {d['member']}")

    # 해약환급금준비금
    print(f"\n  [해약환급금준비금 SurrenderValueReserve]")
    shown = 0
    seen_sig = set()
    for elem in root:
        local = get_local(elem.tag)
        if "surrendervaluereserve" not in local.lower():
            continue
        if any(x in local for x in ["ToBeAdded", "Appropriation", "Transfer", "Adjusted", "Net"]):
            continue
        ctx_ref = elem.get("contextRef", "")
        if not elem.text:
            continue
        try:
            raw = float(elem.text.strip())
        except ValueError:
            continue
        val_mn = to_mn(raw)
        ctx = contexts.get(ctx_ref, {})
        dims = ctx.get("dims", [])
        pdate = ctx.get("pdate", "?")
        if pdate not in ["2025-12-31", "2024-12-31"]:
            continue
        is_sep = any("Separate" in d["member"] for d in dims)
        is_con = any("Consolidated" in d["member"] for d in dims)
        fs = "Sep" if is_sep else ("Con" if is_con else "?")
        dim_key = tuple(sorted((d["axis"], d["member"]) for d in dims))
        sig = (local, fs, len(dims), dim_key, pdate)
        if sig in seen_sig: continue
        seen_sig.add(sig)
        print(f"    tag={local} | val_mn={val_mn:.0f} | {fs} | n_dim={len(dims)} | date={pdate}")
        for d in dims[:3]:
            print(f"      {d['axis']} = {d['member']}")
        shown += 1
        if shown >= 4: break


def main():
    for company, info in COMPANIES.items():
        analyze(company, info["path"], info["targets"])


if __name__ == "__main__":
    main()
