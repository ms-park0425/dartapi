"""
XBRL 파일에서 FAIL 항목 역탐색 - 더 집중적 분석
"""
import xml.etree.ElementTree as ET
import os
from collections import defaultdict

NS_XBRLI = 'http://www.xbrl.org/2003/instance'
NS_XBRLDI = 'http://xbrl.org/2006/xbrldi'


def parse_xbrl(filepath):
    tree = ET.parse(filepath)
    root = tree.getroot()
    contexts = {}
    for ctx in root.findall(f'{{{NS_XBRLI}}}context'):
        cid = ctx.get('id')
        period_el = ctx.find(f'{{{NS_XBRLI}}}period')
        period = {}
        if period_el is not None:
            instant = period_el.find(f'{{{NS_XBRLI}}}instant')
            start = period_el.find(f'{{{NS_XBRLI}}}startDate')
            end = period_el.find(f'{{{NS_XBRLI}}}endDate')
            if instant is not None:
                period['type'] = 'instant'
                period['date'] = instant.text
            elif start is not None and end is not None:
                period['type'] = 'duration'
                period['start'] = start.text
                period['end'] = end.text
        dims = {}
        seg = ctx.find(f'.//{{{NS_XBRLI}}}segment')
        if seg is not None:
            for member in seg.findall(f'{{{NS_XBRLDI}}}explicitMember'):
                dim = member.get('dimension', '')
                val = member.text or ''
                dim_local = dim.split(':')[-1] if ':' in dim else dim
                val_local = val.split(':')[-1] if ':' in val else val
                dims[dim_local] = val_local
        contexts[cid] = {'period': period, 'dims': dims}

    facts = defaultdict(list)
    for child in root:
        tag = child.tag
        if tag.startswith(f'{{{NS_XBRLI}}}'):
            continue
        local = tag.split('}')[-1] if '}' in tag else tag
        ctx_ref = child.get('contextRef')
        decimals = child.get('decimals', '')
        val_text = child.text
        if ctx_ref and val_text:
            try:
                val = float(val_text.replace(',', ''))
                facts[local].append({
                    'context_id': ctx_ref,
                    'value': val,
                    'decimals': decimals,
                    'context': contexts.get(ctx_ref, {})
                })
            except (ValueError, AttributeError):
                pass

    return facts, contexts


def fmt_ctx(e):
    ctx = e['context']
    period = ctx.get('period', {})
    dims = ctx.get('dims', {})
    pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
    return f"period={pstr}, ndim={len(dims)}, dims={dims}"


def search_near(facts, target, tol=0.01):
    results = []
    for tag, entries in facts.items():
        for e in entries:
            v = e['value']
            if abs(target) > 0 and abs(v - target) / abs(target) <= tol:
                results.append((tag, e, abs(v - target) / abs(target)))
    results.sort(key=lambda x: x[2])
    return results


base = "C:/MALIFE/dartapi"

# =================================================================
# 1. Col78 CSM 태그 분석 - 각 회사별로 태그 모든 값 출력
# 특히 ndim=1 Separate 또는 ndim=0 컨텍스트에 집중
# =================================================================
print("=" * 70)
print("1. Col78 CSM 태그 - Separate+ndim 분석")
print("=" * 70)

CSM_TAG = "IncreaseDecreaseThroughChangesInEstimatesThatAdjustContractualServiceMarginInsuranceContractsLiabilityAsset"

col78 = {
    "삼성생명": ("data/삼성생명/사업보고서/2025/entity00126256_2025-12-31.xbrl", 1757775, -1757775),
    "삼성화재": ("data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl", -3372236, -1686117),
    "동양생명": ("data/동양생명/사업보고서/2025/entity00117267_2025-12-31.xbrl", -1213760, -606880),
    "한화생명": ("data/한화생명/사업보고서/2025/entity00113058_2025-12-31.xbrl", -4064448, -2032224),
}

for company, (fpath, our_val, answer_val) in col78.items():
    fp = os.path.join(base, fpath)
    facts, _ = parse_xbrl(fp)
    entries = facts.get(CSM_TAG, [])
    print(f"\n[{company}] CSM 태그 entries: {len(entries)}개, 우리={our_val}, 정답={answer_val}")
    # 값이 있는 entries만, period=duration 2024-01-01~2024-12-31 또는 2025
    for e in entries:
        v = e['value']
        if v == 0:
            continue
        ctx = e['context']
        dims = ctx.get('dims', {})
        period = ctx.get('period', {})
        # Separate member + 당기(2024-01-01~2024-12-31) 집중
        is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
        is_duration = period.get('type') == 'duration'
        has_dis = 'DisaggregationOfInsuranceContractsAxis' in dims
        has_type = 'TypesOfContractsAxis' in dims
        has_comp = 'InsuranceContractsByComponentsAxis' in dims

        pstr = f"{period.get('start')}~{period.get('end')}" if is_duration else period.get('date')
        sep_str = "SEP" if is_sep else "---"
        re_str = dims.get('DisaggregationOfInsuranceContractsAxis', '')
        type_str = dims.get('TypesOfContractsAxis', '')
        comp_str = dims.get('InsuranceContractsByComponentsAxis', '')[:50] if 'InsuranceContractsByComponentsAxis' in dims else ''

        print(f"  {sep_str} ndim={len(dims)} val={v:.0f} dec={e['decimals']} period={pstr}")
        print(f"       Disagg={re_str}, TypesOfContracts={type_str}")
        print(f"       Comp={comp_str}")

    # 정답값에 가까운 다른 태그 탐색
    print(f"  --- 정답값 {answer_val} 근사 탐색 ---")
    hits = search_near(facts, answer_val, 0.005)
    for tag, e, diff in hits[:5]:
        ctx = e['context']
        dims = ctx.get('dims', {})
        period = ctx.get('period', {})
        pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
        print(f"    태그={tag}, val={e['value']:.0f}, period={pstr}, ndim={len(dims)}, dims_keys={list(dims.keys())}")

# =================================================================
# 2. Col99 - 해약환급금준비금 Surrender 태그 및 누적값 탐색
# =================================================================
print("\n" + "=" * 70)
print("2. Col99 해약환급금준비금 - 전기+당기 누적 탐색")
print("=" * 70)

col99 = {
    "삼성화재": ("data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl", 2213099, 4130070),
    "메리츠화재": ("data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl", 1793089, 2976566),
    "KB손보": ("data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl", 2838012, 3646236),
    "DB손보": ("data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl", 3239227, 4541287),
}

for company, (fpath, our_val, answer_val) in col99.items():
    fp = os.path.join(base, fpath)
    facts, _ = parse_xbrl(fp)

    # Surrender 관련
    surr_tags = [k for k in facts if 'surrender' in k.lower() or 'SurrenderValue' in k or 'Refund' in k.lower() or 'Cancellation' in k.lower()]
    print(f"\n[{company}] Surrender 태그: {surr_tags}")

    for st in surr_tags:
        for e in facts[st]:
            v = e['value']
            if abs(v) < 1000:
                continue
            ctx = e['context']
            dims = ctx.get('dims', {})
            period = ctx.get('period', {})
            pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
            is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
            print(f"  {st}: val={v:.0f}, period={pstr}, ndim={len(dims)}, sep={is_sep}, dims={dims}")

    # 정답/전기 근사
    prev_est = answer_val - our_val
    print(f"  우리={our_val}, 정답={answer_val}, 전기추정={prev_est}")
    for tv in [answer_val, our_val, prev_est]:
        hits = search_near(facts, tv, 0.01)
        if hits:
            print(f"  target={tv:.0f} 근사:")
            for tag, e, diff in hits[:3]:
                ctx = e['context']
                dims = ctx.get('dims', {})
                period = ctx.get('period', {})
                pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
                is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
                print(f"    {tag}: val={e['value']:.0f}, period={pstr}, ndim={len(dims)}, sep={is_sep}")

# =================================================================
# 3. Col7 AC자산 - BS instant 2025-12-31 Separate
# =================================================================
print("\n" + "=" * 70)
print("3. Col7 AC자산 - instant 2025-12-31 Separate 탐색")
print("=" * 70)

col7 = {
    "신한라이프": ("data/신한라이프생명/사업보고서/2025/entity00137517_2025-12-31.xbrl", 7402581, 3950344),
    "교보생명": ("data/교보생명/사업보고서/2025/entity00112882_2025-12-31.xbrl", 20275183, 21546474),
    "메리츠화재": ("data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl", 530914, 972104),
    "KB손보": ("data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl", 7480572, 8239538),
}

for company, (fpath, our_val, answer_val) in col7.items():
    fp = os.path.join(base, fpath)
    facts, _ = parse_xbrl(fp)

    # AC 관련 태그
    ac_tags = [k for k in facts if 'amortised' in k.lower() or 'AmortisedCost' in k or 'MeasuredAtAmortised' in k.lower()]
    print(f"\n[{company}] AC 태그 ({len(ac_tags)}개): {ac_tags[:10]}")

    # instant 2025-12-31, Separate context 필터
    for at in ac_tags:
        for e in facts[at]:
            ctx = e['context']
            period = ctx.get('period', {})
            dims = ctx.get('dims', {})
            if period.get('type') != 'instant':
                continue
            if period.get('date') != '2025-12-31':
                continue
            is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
            v = e['value']
            print(f"  {at}: val={v:.0f}, ndim={len(dims)}, sep={is_sep}, dims={dims}")

    # 정답값 근사
    print(f"  우리={our_val}, 정답={answer_val}")
    hits = search_near(facts, answer_val, 0.02)
    print(f"  정답값 {answer_val} 2% 근사:")
    for tag, e, diff in hits[:8]:
        ctx = e['context']
        dims = ctx.get('dims', {})
        period = ctx.get('period', {})
        pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
        is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
        print(f"    {tag}: val={e['value']:.0f}, period={pstr}, ndim={len(dims)}, sep={is_sep}")

    # 우리값 근사 태그 확인
    print(f"  우리값 {our_val} 2% 근사:")
    hits2 = search_near(facts, our_val, 0.02)
    for tag, e, diff in hits2[:8]:
        ctx = e['context']
        dims = ctx.get('dims', {})
        period = ctx.get('period', {})
        pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
        is_sep = dims.get('ConsolidatedAndSeparateFinancialStatementsAxis') == 'SeparateMember'
        print(f"    {tag}: val={e['value']:.0f}, period={pstr}, ndim={len(dims)}, sep={is_sep}")

print("\n완료!")
