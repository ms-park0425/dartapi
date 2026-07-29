"""
XBRL 파일에서 FAIL 항목 역탐색 스크립트
"""
import xml.etree.ElementTree as ET
import os
import re
from collections import defaultdict

NS_XBRLI = 'http://www.xbrl.org/2003/instance'
NS_XBRLDI = 'http://xbrl.org/2006/xbrldi'

def parse_xbrl(filepath):
    """XBRL 파일 파싱: {tag_local: [(context_ref, value, decimals)]}"""
    tree = ET.parse(filepath)
    root = tree.getroot()

    # context 파싱: context_id -> {period, dims}
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
                # local name only
                dim_local = dim.split(':')[-1] if ':' in dim else dim
                val_local = val.split(':')[-1] if ':' in val else val
                dims[dim_local] = val_local
        contexts[cid] = {'period': period, 'dims': dims}

    # 팩트 파싱
    facts = defaultdict(list)
    for child in root:
        tag = child.tag
        if tag.startswith(f'{{{NS_XBRLI}}}'):
            continue
        # local name
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


def search_by_value(facts, target_values, tolerance_ratio=0.02):
    """target_values에 근사한 팩트 찾기 (2% 이내)"""
    results = []
    for tag, entries in facts.items():
        for e in entries:
            v = e['value']
            for tv in target_values:
                if tv == 0:
                    continue
                if abs(v - tv) / abs(tv) <= tolerance_ratio:
                    results.append({
                        'tag': tag,
                        'value': v,
                        'target': tv,
                        'context_id': e['context_id'],
                        'context': e['context'],
                        'decimals': e['decimals']
                    })
    return results


def print_tag_details(facts, tag_name, company_name):
    """특정 태그의 모든 context/값 출력"""
    print(f"\n  [{company_name}] 태그: {tag_name}")
    entries = facts.get(tag_name, [])
    if not entries:
        # partial match
        matches = [k for k in facts if tag_name.lower() in k.lower()]
        if matches:
            print(f"    (정확히 없음, 유사 태그: {matches[:5]})")
        else:
            print(f"    (태그 없음)")
        return
    for e in entries:
        ctx = e['context']
        period = ctx.get('period', {})
        dims = ctx.get('dims', {})
        ndim = len(dims)
        period_str = ''
        if period.get('type') == 'instant':
            period_str = f"instant:{period.get('date')}"
        elif period.get('type') == 'duration':
            period_str = f"{period.get('start')}~{period.get('end')}"
        print(f"    ctx={e['context_id']}, period={period_str}, ndim={ndim}, dims={dims}, val={e['value']:.0f}, dec={e['decimals']}")


def main():
    base = "C:/MALIFE/dartapi"

    # =================================================================
    # 1. Col78 (CSM조정) 분석
    # =================================================================
    print("=" * 80)
    print("1. Col78 (CSM조정) 분석")
    print("=" * 80)

    CSM_TAG = "IncreaseDecreaseThroughChangesInEstimatesThatAdjustContractualServiceMarginInsuranceContractsLiabilityAsset"

    col78_cases = {
        "삼성생명": {
            "file": "data/삼성생명/사업보고서/2025/entity00126256_2025-12-31.xbrl",
            "our": 1757775,  # 우리 값 (부호 반대라는 뜻: 우리는 +1757775, 정답은 -1757775)
            "answer": -1757775
        },
        "삼성화재": {
            "file": "data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
            "our": -3372236,
            "answer": -1686117
        },
        "동양생명": {
            "file": "data/동양생명/사업보고서/2025/entity00117267_2025-12-31.xbrl",
            "our": -1213760,
            "answer": -606880
        },
        "한화생명": {
            "file": "data/한화생명/사업보고서/2025/entity00113058_2025-12-31.xbrl",
            "our": -4064448,
            "answer": -2032224
        },
    }

    for company, info in col78_cases.items():
        filepath = os.path.join(base, info["file"])
        print(f"\n--- {company} ---")
        if not os.path.exists(filepath):
            print(f"  파일 없음: {filepath}")
            continue
        facts, contexts = parse_xbrl(filepath)

        # 1a) CSM 태그 상세
        print_tag_details(facts, CSM_TAG, company)

        # 1b) 정답값 근사 검색 (절대값)
        answer = info["answer"]
        target_vals = [answer, -answer, answer/2, -answer/2, answer*2, -answer*2]
        print(f"\n  정답값={answer:.0f} 근사 검색 (tolerance 0.5%):")
        found = search_by_value(facts, target_vals, tolerance_ratio=0.005)
        if found:
            for r in sorted(found, key=lambda x: abs(x['value'] - x['target'])):
                ctx = r['context']
                dims = ctx.get('dims', {})
                period = ctx.get('period', {})
                pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
                print(f"    태그={r['tag']}, val={r['value']:.0f}, target={r['target']:.0f}, period={pstr}, ndim={len(dims)}, dims={dims}")
        else:
            print("    (근사값 없음)")

    # =================================================================
    # 2. Col99 (해약환급금준비금) 분석
    # =================================================================
    print("\n" + "=" * 80)
    print("2. Col99 (해약환급금준비금) 분석")
    print("=" * 80)

    col99_cases = {
        "삼성화재": {
            "file": "data/삼성화재/사업보고서/2025/entity00139214_2025-12-31.xbrl",
            "our": 2213099,
            "answer": 4130070
        },
        "메리츠화재": {
            "file": "data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
            "our": 1793089,
            "answer": 2976566
        },
        "KB손보": {
            "file": "data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
            "our": 2838012,
            "answer": 3646236
        },
        "DB손보": {
            "file": "data/DB손해보험/사업보고서/2025/entity00159102_2025-12-31.xbrl",
            "our": 3239227,
            "answer": 4541287
        },
    }

    SURRENDER_KEYWORDS = ['SurrenderValue', 'surrender_value', 'Surrender', 'RefundReserve', 'CancellationReserve']

    for company, info in col99_cases.items():
        filepath = os.path.join(base, info["file"])
        print(f"\n--- {company} ---")
        if not os.path.exists(filepath):
            print(f"  파일 없음: {filepath}")
            continue
        facts, contexts = parse_xbrl(filepath)

        # Surrender 관련 태그 목록
        surrender_tags = [k for k in facts if any(kw.lower() in k.lower() for kw in SURRENDER_KEYWORDS)]
        if surrender_tags:
            print(f"  Surrender 관련 태그: {surrender_tags}")
            for st in surrender_tags:
                print_tag_details(facts, st, company)
        else:
            print("  SurrenderValue 관련 태그 없음")

        # 정답값 근사 검색
        answer = info["answer"]
        our = info["our"]
        # 정답 = 전기 + 당기, 전기 = 정답 - 당기
        prev = answer - our  # 이전기 값
        print(f"\n  정답값={answer:.0f}, 우리값(당기)={our:.0f}, 전기추정={prev:.0f}")
        print(f"  정답값 근사 검색 (1%):")
        found = search_by_value(facts, [answer, -answer, our, prev, -prev], tolerance_ratio=0.01)
        if found:
            for r in sorted(found, key=lambda x: abs(x['value'] - x['target'])):
                ctx = r['context']
                dims = ctx.get('dims', {})
                period = ctx.get('period', {})
                pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
                print(f"    태그={r['tag']}, val={r['value']:.0f}, target={r['target']:.0f}, period={pstr}, ndim={len(dims)}, dims={dims}")
        else:
            print("    (근사값 없음)")

    # =================================================================
    # 3. Col7 (AC자산) 분석
    # =================================================================
    print("\n" + "=" * 80)
    print("3. Col7 (AC자산) 분석")
    print("=" * 80)

    col7_cases = {
        "신한라이프": {
            "file": "data/신한라이프생명/사업보고서/2025/entity00137517_2025-12-31.xbrl",
            "our": 7402581,
            "answer": 3950344
        },
        "교보생명": {
            "file": "data/교보생명/사업보고서/2025/entity00112882_2025-12-31.xbrl",
            "our": 20275183,
            "answer": 21546474
        },
        "메리츠화재": {
            "file": "data/메리츠화재/사업보고서/2025/entity00117744_2025-12-31.xbrl",
            "our": 530914,
            "answer": 972104
        },
        "KB손보": {
            "file": "data/KB손해보험/사업보고서/2025/entity00120216_2025-12-31.xbrl",
            "our": 7480572,
            "answer": 8239538
        },
    }

    AC_KEYWORDS = ['AmortisedCost', 'amortised_cost', 'amortisedcost', 'FinancialAssets', 'FinancialInstruments',
                   'MeasuredAt', 'FinancialAssetsAtAmortised', 'ACMeasurement']

    for company, info in col7_cases.items():
        filepath = os.path.join(base, info["file"])
        print(f"\n--- {company} ---")
        if not os.path.exists(filepath):
            print(f"  파일 없음: {filepath}")
            continue
        facts, contexts = parse_xbrl(filepath)

        # AC 관련 태그 목록 (instant, separate context 포함)
        ac_tags = [k for k in facts if any(kw.lower() in k.lower() for kw in AC_KEYWORDS)]
        if ac_tags:
            print(f"  AC 관련 태그 ({len(ac_tags)}개): {ac_tags[:20]}")
        else:
            print("  AC 관련 태그 없음")

        # 정답값 근사 검색 (instant 2025-12-31)
        answer = info["answer"]
        our = info["our"]
        print(f"\n  정답값={answer:.0f}, 우리값={our:.0f}")
        print(f"  정답값 근사 검색 (2%):")
        found = search_by_value(facts, [answer, -answer], tolerance_ratio=0.02)
        if found:
            for r in sorted(found, key=lambda x: abs(x['value'] - x['target'])):
                ctx = r['context']
                dims = ctx.get('dims', {})
                period = ctx.get('period', {})
                pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
                print(f"    태그={r['tag']}, val={r['value']:.0f}, target={r['target']:.0f}, period={pstr}, ndim={len(dims)}, dims={dims}")
        else:
            print("    (근사값 없음 - 5% 확대)")
            found2 = search_by_value(facts, [answer, -answer], tolerance_ratio=0.05)
            for r in sorted(found2, key=lambda x: abs(x['value'] - x['target'])):
                ctx = r['context']
                dims = ctx.get('dims', {})
                period = ctx.get('period', {})
                pstr = period.get('date') or f"{period.get('start')}~{period.get('end')}"
                print(f"    태그={r['tag']}, val={r['value']:.0f}, target={r['target']:.0f}, period={pstr}, ndim={len(dims)}, dims={dims}")

    print("\n\n완료!")

if __name__ == "__main__":
    main()
