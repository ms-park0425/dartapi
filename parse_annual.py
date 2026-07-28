"""
사업보고서 XBRL에서 DATA 시트에 필요한 값을 추출하는 스크립트.
단위: 백만원 (원 / 1e6)

DATA 시트 구조:
- Row 1: 번호 (1,2,3,...)
- Row 2: 제목
- Row 3: 카테고리 (B/S, P/L, ...)
- Row 4: 대분류
- Row 5: 소분류
- Row 6~: 데이터 (기간+회사 조합)

각 행의 Col1 = "YYMM회사약칭" (예: 2512미래)
Col2 = 기간코드 (예: 2512)
Col3 = 회사약칭

열 매핑 (Col → XBRL account_id + context):
"""
import os
import sys
import csv
import xml.etree.ElementTree as ET
from glob import glob

sys.stdout.reconfigure(encoding="utf-8")

NS_XBRLI = "http://www.xbrl.org/2003/instance"
NS_XBRLDI = "http://xbrl.org/2006/xbrldi"

COMPANIES = [
    {"name": "미래에셋생명", "short": "미래"},
    {"name": "삼성생명", "short": "삼성"},
    {"name": "교보생명", "short": "교보"},
    {"name": "한화생명", "short": "한화"},
    {"name": "신한라이프생명", "short": "신한라이프"},
    {"name": "동양생명", "short": "동양"},
    {"name": "KB생명", "short": "KB라이프"},
    {"name": "삼성화재", "short": "삼성화재"},
    {"name": "현대해상", "short": "현대해상"},
    {"name": "메리츠화재", "short": "메리츠화재"},
    {"name": "KB손해보험", "short": "KB손보"},
    {"name": "DB손해보험", "short": "DB손보"},
]

# ====== XBRL → DATA 열 매핑 ======
# BS 항목 (별도, 당기말 instant context)
BS_MAPPING = {
    4: "ifrs-full_Assets",
    5: "ifrs-full_FinancialAssetsAtFairValueThroughProfitOrLoss",
    6: "ifrs-full_FinancialAssetsAtFairValueThroughOtherComprehensiveIncome",  # dart_ fallback: SecuritiesAtFairValueThroughOtherComprehensiveIncome
    # Col7: 상각후원가 자산 - LoansAtAmortisedCost + SecuritiesAtAmortisedCost 합산도 후보
    7: "ifrs-full_FinancialAssetsAtAmortisedCost",
    8: "ifrs-full_InvestmentsInSubsidiariesJointVenturesAndAssociates",
    # fallback in extract_company_data: InvestmentAccountedForUsingEquityMethod
    12: "ifrs-full_Liabilities",
    13: "ifrs-full_InsuranceContractsIssuedThatAreLiabilities",
    15: "ifrs-full_Equity",  # 총자본
    16: "dart_HybridBonds",  # 신종자본증권
    17: "ifrs-full_AccumulatedOtherComprehensiveIncome",
    18: "ifrs-full_RetainedEarnings",
}

# PL 항목 (별도, 당기 duration context)
PL_MAPPING = {
    19: "dart_EquityAtBeginningOfPeriod",  # 기초자본
    20: "dart_EquityAtBeginningOfPeriod",  # 기초자본(신종자본증권제외) - 같은 값
    32: "ifrs-full_InsuranceServiceResult",
    33: "dart_OperatingIncomeInsurance",  # 또는 ifrs-full_OperatingIncomeInsurance
    34: "dart_OperatingExpenseInsurance",  # 또는 ifrs-full_OperatingExpenseInsurance
    35: "dart_InvestmentIncomeExpenses",  # 또는 ifrs-full_InvestmentIncomeExpenses
    39: "ifrs-full_SellingGeneralAndAdministrativeExpense",
    41: "ifrs-full_ProfitLossFromOperatingActivities",
    43: "ifrs-full_ProfitLossBeforeTax",
    44: "ifrs-full_ProfitLoss",
}

# 계산 항목 (두 PL 항목의 차이)
CALC_MAPPING = {
    36: ("dart_InsuranceFinanceIncomeFromInsuranceContractsIssuedRecognisedInProfitOrLoss",
         "dart_InsuranceFinanceExpensesFromInsuranceContractsIssuedRecognisedInProfitOrLoss",
         "subtract"),  # income - expense
    37: ("dart_FinanceIncomeFromReinsuranceContractsHeldRecognisedInProfitOrLoss",
         "dart_FinanceExpensesFromReinsuranceContractsHeldRecognisedInProfitOrLoss",
         "subtract"),
    42: ("ifrs-full_NonOperatingProfitLoss", None, "direct"),
    50: ("ifrs-full_NonOperatingProfitLoss", None, "direct"),
}

# OCI 항목 (별도, 당기 duration) - 각 col에 여러 후보 태그 (순서대로 시도)
OCI_MAPPING = {
    56: [
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncome",
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncomeTotal",
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxChangeInFairValueOfFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncome",
    ],
    # Col59는 OCI_SUM_MAPPING으로 처리 (CashFlowHedges + GainsLossesOnHedgingInstruments 합산)
    60: [
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxGainsLossesOnRevaluation",
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxGainsLossesOnRevaluationOfPropertyPlantAndEquipment",
    ],
    61: ["ifrs-full_OtherComprehensiveIncomeNetOfTaxGainsLossesOnRemeasurementsOfDefinedBenefitPlans"],
}

# Col58 = 보험OCI + 재보험OCI (두 항목 합산)
OCI_SUM_MAPPING = {
    58: (
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxInsuranceFinanceIncomeExpensesFromInsuranceContractsIssuedExcludedFromProfitOrLossThatWillBeReclassifiedToProfitOrLoss",
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxFinanceIncomeExpensesFromReinsuranceContractsHeldExcludedFromProfitOrLoss",
    ),
    59: (
        "ifrs-full_OtherComprehensiveIncomeNetOfTaxCashFlowHedges",
        "dart_OtherComprehensiveIncomeNetOfTaxGainsLossesOnHedgingInstrument",
    ),
}

# 전기말 BS 항목
PRIOR_BS_MAPPING = {
    55: "ifrs-full_AccumulatedOtherComprehensiveIncome",  # 전기말 기타포괄손익누계액
    19: "ifrs-full_Equity",                              # 기초자본 (전기말)
    65: "ifrs-full_RetainedEarnings",                    # 이익잉여금 기초 (전기말)
    71: "ifrs-full_RetainedEarnings",                    # 이익잉여금 기말 (당기말 BS와 동일)
    # 전기말 신종자본증권 (Col20 계산용)
    "_prior_hybrid": "dart_HybridBonds",
}

# OCI 누계 항목 - 당기말 BS context에서 ChangesInAccumulatedOCI axis 멤버 키워드로 탐색
# (Col24~30): 3-dim context (Separate + OtherComprehensiveIncomeLossAccumulatedAmountMember + entity-specific)
OCI_ACCUM_MAPPING = {
    # col: list of keywords (member 이름에 하나라도 포함되면 해당 col에 매핑)
    # 더 구체적인 키워드를 앞에 배치, 순서가 중요함
    24: ["OtherComprehensiveIncome"],
    25: ["FinancialAssets", "GainsAndLossesOnFinancialAssets", "GainsAndLossesFromInvestments",
         "GainsAndLossesOnDebtAndEquity"],  # FVOCI 관련 합산
    26: ["InsuranceFinanceIncomeExpenses", "InsuranceContract"],
    27: ["ReinsuranceFinance", "ReinsuranceContract"],
    28: ["CashFlowHedges", "CashFlow", "Hedging", "ExchangeDifferences"],
    29: ["Revaluation"],          # 전기말 사용
    30: ["RemeasurementsOfDefinedBenefit", "DefinedBenefit"],
}

# 계리적 가정 변경 민감도 분석 (Col146~158)
# 3-dim duration: Separate + InsuranceContractsIssuedMember + BEL or CSM component
ACTUARIAL_MAPPING = {
    # (col, component_keyword): tag_local
    (146, "BEL"): "dart_ChangeInCancellationRateAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (147, "BEL"): "dart_ChangeInRiskRateAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (148, "BEL"): "dart_ChangeInProjectRatioAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (149, "BEL"): "dart_OtherAssumptionChangesOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (150, "BEL"): "dart_FluctuationsDueToVolumeDifferencesAndDifferencesInInvestmentFactorsExpectedAndActualOfChangesInFutureServicesAttributableToFactorsOtherThanNewContractsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (153, "CSM"): "dart_ChangeInCancellationRateAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (154, "CSM"): "dart_ChangeInRiskRateAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (155, "CSM"): "dart_ChangeInProjectRatioAssumptionOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (156, "CSM"): "dart_OtherAssumptionChangesOfEffectOfChangingAssumptionsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (157, "CSM"): "dart_FluctuationsDueToVolumeDifferencesAndDifferencesInInvestmentFactorsExpectedAndActualOfChangesInFutureServicesAttributableToFactorsOtherThanNewContractsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
    (158, "CSM"): "dart_FluctuationsDueToLossFactorsOfChangesInFutureServicesAttributableToFactorsOtherThanNewContractsOfChangesInInsuranceLiabilitiesBasedOnActuarialAssumptionsTableOfItems",
}

# 추가 PL 항목
EXTRA_PL_MAPPING = {
    68: "ifrs-full_ProfitLoss",          # 순이익 (Col44와 같음)
    48: "ifrs-full_InsuranceServiceResult",  # 보험손익 (Col32와 같음)
    49: "ifrs-full_InvestmentIncomeExpenses",  # 투자손익 (Col35와 같음)
}

# Col40: 기타투자손익 = FeeIncome + RentalIncome + OtherIncomeInvestment - OtherExpenseInvestment
# 태그는 회사마다 다름, 아래 각 항목에 여러 후보
COL40_FEE = ["dart_FeeIncomeOfInvestmentIncome", "dart_FeesAndCommissionIncomeOfInvestmentIncome",
             "ifrs-full_FeeAndCommissionIncome"]
COL40_RENT = ["dart_RentalReturnOfInvestmentIncome", "ifrs-full_RentalIncome"]
COL40_OTHER_INC = ["ifrs-full_OtherOperatingIncomeInvestment", "dart_OtherOperatingIncomeInvestment"]
# 투자비용: OtherOperatingExpenseInvestment + PropertyManagementExpense + udf 계열 모두 포함
# udf_ + OfOperatingExpenseInvestment는 _get_pl_by_keyword가 아닌 별도 처리
COL40_OTHER_EXP_KEYWORDS = ["OtherOperatingExpenseInvestment", "PropertyManagementExpense"]
# DB손보 패턴: udf_IS_ + OfOperatingExpenseInvestment (당해연도 신규 정의)
COL40_UDF_EXP_PREFIX = "OfOperatingExpenseInvestment"

# 복합 축 항목 (CSM 변동표 등)
# 별도 + InsuranceContractsIssuedMember + 각 ComponentMember
INSURANCE_COMPONENT_MAPPING = {
    # Col: (component_member, tag_name) - duration context에서 기말잔액 합산
    # 이건 instant context에서 세부 disaggregation별 합산 필요
}


def parse_actuarial_sensitivity(xbrl_path, target_year="2025"):
    """
    계리적 가정 변경 민감도 분석 파싱.
    3-dim duration: Separate + InsuranceContractsIssuedMember + BEL or CSM component
    반환: {(col, component): value_won}
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    bel_member = "EstimatesOfPresentValueOfFutureCashFlowsMember"
    csm_members = {
        "ContractualServiceMarginMember",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachHasBeenAppliedMember",
        "ContractualServiceMarginNotRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachOrFairValueApproachHasBeenAppliedMember",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichFairValueApproachHasBeenAppliedMember",
    }

    # Build target context dict (3-dim 이상 지원)
    target_ctxs = {}  # ctx_id -> "BEL" or "CSM"
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or start.text != f"{target_year}-01-01":
            continue
        if end is None or end.text != f"{target_year}-12-31":
            continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue
        if "InsuranceContractsIssuedMember" != dims.get("DisaggregationOfInsuranceContractsAxis", ""):
            continue
        comp = dims.get("InsuranceContractsByComponentsAxis", "")
        if comp == bel_member:
            target_ctxs[ctx.get("id", "")] = "BEL"
        elif comp == "RiskAdjustmentForNonfinancialRiskMember":
            target_ctxs[ctx.get("id", "")] = "RA"
        elif comp in csm_members:
            target_ctxs[ctx.get("id", "")] = "CSM"

    # tag_local -> list of (col, expected_component_type)
    tag_to_col = {}
    for (col, comp), tag in ACTUARIAL_MAPPING.items():
        tag_local = tag.split("_", 1)[1]
        tag_to_col.setdefault(tag_local, []).append((col, comp))

    # 패턴2: 키워드 기반 context 탐색 (삼성생명 등 entity-specific 태그 지원)
    # 2-dim (Separate + BEL/CSM via DisaggregationAxis or InsuranceContractsByComponentsAxis) + duration
    SENS_KEYWORDS = {
        "BEL": {
            146: ["CancellationRate", "SurrenderRate", "SurrenderRatio", "LapseRate"],
            147: ["RiskRate", "RiskRatio", "RiskFactor", "MortalityRate"],
            148: ["ExpenseRate", "OperatingExpenseRate", "ProjectRatio", "ProjectExpenseRate",
                  "OperatingExpense"],
            149: ["OtherAssumption"],  # OtherChange 제외 (너무 광범위)
            150: ["VolumeDifferences", "QuantityDifferences", "FluctuationsDueTo",
                  "DifferenceOfQuantity", "DifferenceInQuantity", "VolumeDifference"],
            158: ["LossComponent", "LossFactor"],
        },
        "CSM": {
            153: ["CancellationRate", "SurrenderRate", "SurrenderRatio", "LapseRate"],
            154: ["RiskRate", "RiskRatio", "RiskFactor", "MortalityRate"],
            155: ["ExpenseRate", "OperatingExpenseRate", "ProjectRatio"],
            156: ["OtherAssumption"],
            157: ["VolumeDifferences", "QuantityDifferences", "FluctuationsDueTo",
                  "DifferenceOfQuantity", "DifferenceInQuantity"],
        },
    }

    # alt_ctxs: 패턴2 context (삼성생명 2-dim + 현대해상/DB손보/한화 5~6-dim entity-specific)
    alt_ctxs = {}  # ctx_id -> "BEL" or "CSM"
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or end is None: continue
        if start.text != f"{target_year}-01-01" or end.text != f"{target_year}-12-31": continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue
        # 이미 target_ctxs에 있으면 skip
        ctx_id = ctx.get("id", "")
        if ctx_id in target_ctxs: continue

        comp = dims.get("InsuranceContractsByComponentsAxis", "")
        disagg = dims.get("DisaggregationOfInsuranceContractsAxis", "")

        # 패턴A: 2-dim, DisaggregationAxis=BEL/RA/CSM (삼성생명)
        if len(members) == 2 and not comp:
            if bel_member in disagg:
                alt_ctxs[ctx_id] = "BEL"
            elif "RiskAdjustmentForNonfinancialRiskMember" in disagg:
                alt_ctxs[ctx_id] = "RA"
            elif any(m in disagg for m in csm_members):
                alt_ctxs[ctx_id] = "CSM"
        # 패턴B: 5~6-dim, InsuranceContractsIssuedMember + BEL/CSM component (현대해상/DB손보/한화)
        elif len(members) >= 4 and "InsuranceContractsIssuedMember" == disagg:
            if comp == bel_member:
                alt_ctxs[ctx_id] = "BEL"
            elif comp == "RiskAdjustmentForNonfinancialRiskMember":
                alt_ctxs[ctx_id] = "RA"
            elif comp in csm_members:
                alt_ctxs[ctx_id] = "CSM"

    # Accumulate: standard ACTUARIAL_MAPPING (미래에셋 패턴)
    result = {}
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in target_ctxs or not elem.text:
            continue
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag not in tag_to_col:
            continue
        try:
            v = float(elem.text.strip())
        except ValueError:
            continue
        comp_type = target_ctxs[ctx_ref]
        for col, expected_comp in tag_to_col[tag]:
            if expected_comp == "BEL" and comp_type in ("BEL", "RA"):
                result[(col, "BEL")] = result.get((col, "BEL"), 0.0) + v
            elif expected_comp == comp_type:
                result[(col, comp_type)] = result.get((col, comp_type), 0.0) + v

    # 키워드 기반 탐색: target_ctxs + alt_ctxs 모두 탐색
    # (ACTUARIAL_MAPPING에 없는 entity-specific 태그 지원)
    keyword_seen = set()  # (ctx_id, col) -> 중복 방지
    all_keyword_ctxs = {**target_ctxs, **alt_ctxs}
    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in all_keyword_ctxs or not elem.text:
            continue
        comp_type = all_keyword_ctxs[ctx_ref]
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        # ACTUARIAL_MAPPING에 있는 태그는 이미 처리됨
        if tag in tag_to_col:
            continue
        try:
            v = float(elem.text.strip())
        except ValueError:
            continue
        if v == 0:
            continue
        # 키워드 매핑 (BEL과 RA는 같은 col에 합산)
        effective_comp = "BEL" if comp_type in ("BEL", "RA") else comp_type
        for col, kwords in SENS_KEYWORDS.get(effective_comp, {}).items():
            if any(k.lower() in tag.lower() for k in kwords):
                pair = (ctx_ref, col)
                if pair not in keyword_seen:
                    keyword_seen.add(pair)
                    result[(col, effective_comp)] = result.get((col, effective_comp), 0.0) + v
                break
    return result


def parse_prior_csm(xbrl_path, target_year="2025"):
    """
    전기말 CSM 합계 파싱. 여러 패턴 지원.
    패턴A(미래에셋): 3-dim entity-disagg + ContractualServiceMarginMember, tag=InsuranceContractsIssuedThatAreLiabilities
    패턴B(삼성생명): 2-dim (Separate+IssuedMember), tag=ContractualServiceMargin
    패턴C(삼성생명 4-dim): TypesOfContracts + BEL/RA 방식에서 total 도출
    """
    prev_year = str(int(target_year) - 1)
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    # 패턴A: 3-dim entity-specific disagg + ContractualServiceMarginMember
    total_a = 0.0
    target_ctxs_a = set()
    # 패턴B: 2-dim (Separate+IssuedMember) + ContractualServiceMargin tag
    total_b = None

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        instant = period.find(f"{{{NS_XBRLI}}}instant")
        if instant is None or instant.text != f"{prev_year}-12-31":
            continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue

        # 패턴A: 3-dim entity-specific disagg
        if len(members) == 3:
            if "ContractualServiceMarginMember" != dims.get("InsuranceContractsByComponentsAxis", ""):
                continue
            disagg = dims.get("DisaggregationOfInsuranceContractsAxis", "")
            if disagg in ("InsuranceContractsIssuedMember", "ReinsuranceContractsHeldMember", ""):
                continue
            target_ctxs_a.add(ctx.get("id", ""))

        # 패턴B: 2-dim (Separate + IssuedMember)
        elif len(members) == 2:
            if "InsuranceContractsIssuedMember" == dims.get("DisaggregationOfInsuranceContractsAxis", ""):
                ctx_id = ctx.get("id", "")
                for elem in root:
                    if elem.get("contextRef") != ctx_id or not elem.text:
                        continue
                    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if tag == "ContractualServiceMargin":
                        try:
                            total_b = float(elem.text.strip())
                        except ValueError:
                            pass

    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in target_ctxs_a or not elem.text:
            continue
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag == "InsuranceContractsIssuedThatAreLiabilities":
            try:
                total_a += float(elem.text.strip())
            except ValueError:
                pass

    if total_a > 0:
        return total_a
    if total_b is not None and total_b > 0:
        return total_b
    return None


def parse_fvoci_oci(xbrl_path, target_year="2025"):
    """
    FVOCI 금융자산 평가손익 OCI 파싱. 여러 context 패턴 지원:
    1. 2-dim (Separate + ReportedAmountMember) + dart_ChangesInOtherComprehensiveIncome...
    2. 2-dim (Separate + AccumulatedOtherComprehensiveIncomeMember) + ifrs-full_...ChangeInFairValue...
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    dart_tag = "ChangesInOtherComprehensiveIncomeFairValueMeasurementGainAndLossOnValuationOfFinancialAssetsEtcOfSignificantTransactionsNotAffectingCashFlowsOfSignificantTransactionsNotAffectingCashFlowsTableOfItems"
    ifrs_tags = [
        "OtherComprehensiveIncomeNetOfTaxChangeInFairValueOfFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncome",
        "OtherComprehensiveIncomeNetOfTaxFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncome",
        "OtherComprehensiveIncomeNetOfTaxFinancialAssetsMeasuredAtFairValueThroughOtherComprehensiveIncomeTotal",
    ]

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or start.text != f"{target_year}-01-01":
            continue
        if end is None or end.text != f"{target_year}-12-31":
            continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        if len(members) != 2:
            continue
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue

        second_dim_vals = [v for k, v in dims.items() if k != "ConsolidatedAndSeparateFinancialStatementsAxis"]
        is_reported = any("ReportedAmountMember" in v for v in second_dim_vals)
        is_accum_oci = any("AccumulatedOtherComprehensiveIncomeMember" in v for v in second_dim_vals)

        if not is_reported and not is_accum_oci:
            continue

        ctx_id = ctx.get("id", "")
        for elem in root:
            if elem.get("contextRef") != ctx_id or not elem.text:
                continue
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if is_reported and tag == dart_tag:
                try:
                    return float(elem.text.strip())
                except ValueError:
                    pass
            if is_accum_oci and tag in ifrs_tags:
                try:
                    return float(elem.text.strip())
                except ValueError:
                    pass
    return None


def parse_oci_accumulated(xbrl_path, target_year="2025"):
    """
    OCI 누계 항목 파싱. 두 가지 구조 지원:
    - 패턴A (미래에셋): 3-dim (Separate + OtherComprehensiveIncomeLossAccumulatedAmountMember + entity-specific), tag=Equity
    - 패턴B (삼성생명): 2-dim (Separate + ComponentsOfEquityAxis member), tag=OtherComprehensiveIncomeLossAccumulatedAmount
    반환: {"cur": {keyword: value_won}, "prior": {keyword: value_won}}
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    prev_year = str(int(target_year) - 1)

    result = {"cur": {}, "prior": {}}

    # 패턴A: 3-dim
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        instant = period.find(f"{{{NS_XBRLI}}}instant")
        if instant is None:
            continue
        if instant.text == f"{target_year}-12-31":
            bucket = "cur"
        elif instant.text == f"{prev_year}-12-31":
            bucket = "prior"
        else:
            continue

        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue

        # 패턴A: OtherComprehensiveIncomeLossAccumulatedAmountMember 포함 + 3-dim
        if len(members) == 3 and any("OtherComprehensiveIncomeLossAccumulatedAmountMember" in v for v in dims.values()):
            change_member = ""
            for k, v in dims.items():
                if "Changes" in k or ("Accumulated" in k and "Separate" not in k):
                    change_member = v
            if change_member:
                ctx_id = ctx.get("id", "")
                for elem in root:
                    if elem.get("contextRef") != ctx_id or not elem.text:
                        continue
                    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if tag == "Equity":
                        try:
                            result[bucket][change_member] = float(elem.text.strip())
                        except ValueError:
                            pass

        # 패턴B: 2-dim with ComponentsOfEquityAxis (OCI 관련 reserve members) - 반드시 2-dim
        elif len(members) == 2 and instant is not None:
            oci_member = ""
            for k, v in dims.items():
                if k in ("ComponentsOfEquityAxis", "ReservesWithinEquityAxis"):
                    oci_member = v
            if not oci_member:
                continue
            # Reserve of FVOCI / Insurance OCI / Reinsurance OCI / CFH / Revaluation / DefinedBenefit
            if not any(kw in oci_member for kw in ("GainsAndLossesOnFinancialAssets", "InsuranceFinance",
                                                     "ReinsuranceFinance", "CashFlowHedges",
                                                     "Revaluation", "RemeasurementsOfDefinedBenefit",
                                                     "ReserveOfGainsAndLosses", "ExchangeDifferences",
                                                     "Hedging")):
                continue
            ctx_id = ctx.get("id", "")
            # ComponentsOfEquityAxis가 있는 context의 값을 우선
            oci_axis = dims.get("ComponentsOfEquityAxis") or dims.get("ReservesWithinEquityAxis", "")
            PREF_TAGS_B = (
                "OtherComprehensiveIncomeLossAccumulatedAmount",
                "AccumulatedOtherComprehensiveIncome",
                "Equity",
            )
            best_val = None
            best_prio = 99
            comp_axis_used = "ComponentsOfEquityAxis" in dims
            for elem in root:
                if elem.get("contextRef") != ctx_id or not elem.text:
                    continue
                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                try:
                    v = float(elem.text.strip())
                except ValueError:
                    continue
                if tag in PREF_TAGS_B:
                    prio = PREF_TAGS_B.index(tag)
                    if prio < best_prio:
                        best_val = v
                        best_prio = prio
            if best_val is not None:
                # ComponentsOfEquityAxis 버전을 우선 (덮어쓰기 허용)
                if comp_axis_used or oci_member not in result[bucket]:
                    result[bucket][oci_member] = best_val
    return result


def parse_csm_movement(xbrl_path, target_year="2025"):
    """
    CSM 변동 항목 파싱. 여러 dim 구조 지원 (5-dim, 4-dim, 3-dim, 2-dim).
    반환: {tag_local_name: total_value_won}
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    csm_tags = {
        "InsuranceRevenueContractualServiceMarginRecognisedInProfitOrLossBecauseOfTransferOfServices": "csm_amortization",
        "IncreaseDecreaseThroughChangesInEstimatesThatAdjustContractualServiceMarginInsuranceContractsLiabilityAsset": "csm_adjustment",
        "InsuranceFinanceIncomeExpensesFromInsuranceContractsIssuedRecognisedInProfitOrLoss": "csm_finance",
        "IncreaseDecreaseThroughEffectsOfContractsInitiallyRecognisedInPeriodInsuranceContractsLiabilityAsset": "new_csm",
        # 삼성생명 패턴
        "IncreaseDecreaseThroughRecognitionOfContractualServiceMarginInProfitOrLossToReflectTransferOfServicesInsuranceContractsLiabilityAsset": "csm_amortization",
        # 이자부리 대체 태그 (한화/교보/현대해상 등)
        "IncreaseDecreaseThroughInsuranceFinanceIncomeOrExpensesInsuranceContractsLiabilityAsset": "csm_finance",
        "IncreaseDecreaseThroughEffectsOfContractsAcquiredInPeriodInsuranceContractsLiabilityAsset": "new_csm",
    }
    component_keys = {
        "ContractualServiceMarginMember",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachHasBeenAppliedMember",
        "ContractualServiceMarginNotRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachOrFairValueApproachHasBeenAppliedMember",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichFairValueApproachHasBeenAppliedMember",
    }

    # 어떤 dim 구조인지 먼저 파악 (4-dim or 3-dim)
    has_4dim = False
    has_5dim = False
    has_3dim = False
    has_6dim = False
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or start.text != f"{target_year}-01-01": continue
        if end is None or end.text != f"{target_year}-12-31": continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {m.get("dimension", "").split(":")[-1]: (m.text or "").split(":")[-1] for m in members}
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""): continue
        if "InsuranceContractsIssuedMember" != dims.get("DisaggregationOfInsuranceContractsAxis", ""): continue
        if dims.get("InsuranceContractsByComponentsAxis", "") not in component_keys: continue
        if len(dims) == 6: has_6dim = True
        elif len(dims) == 5: has_5dim = True
        elif len(dims) == 4: has_4dim = True
        elif len(dims) == 3: has_3dim = True

    # dim 선택 전략:
    # 6-dim 있으면: 6-dim만 (삼성화재 패턴)
    # 4-dim 있으면: 4-dim (상각/신계약) + 5-dim (조정/이자부리) 병행 - 미래에셋 패턴
    # 3-dim만 있으면: 3-dim (삼성생명 패턴)
    target_ndims = set()
    if has_6dim:
        target_ndims.add(6)
    elif has_4dim:
        target_ndims.add(4)
        if has_5dim:
            target_ndims.add(5)  # 미래에셋: 5-dim에도 일부 태그 있음
    elif has_3dim:
        target_ndims.add(3)
    if not target_ndims:
        target_ndims = {3, 4, 5}

    target_ctxs = set()
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or start.text != f"{target_year}-01-01": continue
        if end is None or end.text != f"{target_year}-12-31": continue
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {m.get("dimension", "").split(":")[-1]: (m.text or "").split(":")[-1] for m in members}
        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""): continue
        if "InsuranceContractsIssuedMember" != dims.get("DisaggregationOfInsuranceContractsAxis", ""): continue
        comp = dims.get("InsuranceContractsByComponentsAxis", "")
        if comp not in component_keys: continue
        if len(dims) in target_ndims:
            target_ctxs.add(ctx.get("id", ""))

    # 태그별 합산: 각 컴포넌트 레이블에서 합산
    # 4-dim+5-dim 병행 시 상각/신계약 태그는 4-dim만, 조정/이자 태그는 5-dim에 있는 경우 해당값 사용
    AMORT_TAGS = {
        "InsuranceRevenueContractualServiceMarginRecognisedInProfitOrLossBecauseOfTransferOfServices",
        "IncreaseDecreaseThroughRecognitionOfContractualServiceMarginInProfitOrLossToReflectTransferOfServicesInsuranceContractsLiabilityAsset",
        "IncreaseDecreaseThroughChangesThatRelateToCurrentServiceInsuranceContractsLiabilityAsset",
    }
    NEW_CSM_TAGS = {
        "IncreaseDecreaseThroughEffectsOfContractsInitiallyRecognisedInPeriodInsuranceContractsLiabilityAsset",
        "IncreaseDecreaseThroughEffectsOfContractsAcquiredInPeriodInsuranceContractsLiabilityAsset",
    }

    sums = {v: 0.0 for v in csm_tags.values()}
    # 상각/신계약: 소규모 dim에서 합산 (4 또는 가장 작은 target_ndim)
    # 조정/이자: 모든 target_ndim에서 합산 (단, 상쇄 방지)
    min_ndim = min(target_ndims) if target_ndims else 4
    min_ctxs = set()
    all_ctxs = set()

    for ctx_ref, comp_key in [(c, None) for c in target_ctxs]:
        pass  # rebuild below

    # ctx_id → ndim 매핑 (target_ctxs + adj용 all_sep_ctxs)
    ctx_ndim = {}
    # csm_adjustment 전용: 다양한 context 패턴 지원
    all_sep_ctxs = set()  # adj_tag용
    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or start.text != f"{target_year}-01-01": continue
        if end is None or end.text != f"{target_year}-12-31": continue
        ctx_id = ctx.get("id", "")
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims_t = {m.get("dimension", "").split(":")[-1]: (m.text or "").split(":")[-1] for m in members}
        if "SeparateMember" != dims_t.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""): continue
        # 패턴1: DisaggregationAxis=InsuranceContractsIssuedMember + InsuranceContractsByComponentsAxis=CSM
        # 패턴2(삼성생명): DisaggregationAxis=ContractualServiceMarginMember (2-dim)
        is_pattern1 = (
            "InsuranceContractsIssuedMember" == dims_t.get("DisaggregationOfInsuranceContractsAxis", "")
            and dims_t.get("InsuranceContractsByComponentsAxis", "") in component_keys
        )
        is_pattern2 = (
            "ContractualServiceMarginMember" == dims_t.get("DisaggregationOfInsuranceContractsAxis", "")
            and "InsuranceContractsByComponentsAxis" not in dims_t
        )
        if is_pattern1:
            ctx_ndim[ctx_id] = len(members)
            all_sep_ctxs.add(ctx_id)
            if ctx_id in target_ctxs and len(members) == min_ndim:
                min_ctxs.add(ctx_id)
        elif is_pattern2:
            ctx_ndim[ctx_id] = len(members)
            all_sep_ctxs.add(ctx_id)

    # (context_id, tag) 쌍으로 중복 방지
    # 추가: 같은 tag + comp_key 조합에서 중복 합산 방지 (낮은 ndim 우선)
    seen = set()
    # comp_key별로 값 수집: (comp_key, value) 리스트
    comp_values = {v: [] for v in csm_tags.values()}  # comp_key -> [(ndim, value)]

    adj_tag = "IncreaseDecreaseThroughChangesInEstimatesThatAdjustContractualServiceMarginInsuranceContractsLiabilityAsset"

    for elem in root:
        ctx_ref = elem.get("contextRef", "")
        if not elem.text: continue
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag not in csm_tags: continue
        # adj_tag는 all_sep_ctxs에서 수집, 나머지는 target_ctxs에서
        if tag == adj_tag:
            if ctx_ref not in all_sep_ctxs: continue
        else:
            if ctx_ref not in target_ctxs: continue
        pair = (ctx_ref, tag)
        if pair in seen: continue
        seen.add(pair)
        comp_key = csm_tags[tag]
        ndim = ctx_ndim.get(ctx_ref, 99)
        try:
            v = float(elem.text.strip())
        except ValueError:
            continue
        # 같은 context에서 같은 comp_key는 한 번만 합산
        ctx_comp_pair = (ctx_ref, comp_key)
        if ctx_comp_pair in seen: continue
        if tag in AMORT_TAGS or tag in NEW_CSM_TAGS:
            if ctx_ref in min_ctxs:
                seen.add(ctx_comp_pair)
                comp_values[comp_key].append((ndim, v))
        else:
            seen.add(ctx_comp_pair)
            comp_values[comp_key].append((ndim, v))

    # 각 comp_key에 대해 합산 전략:
    # - csm_adjustment: 가장 낮은 ndim의 합산만 사용 (상위 집계값이 정답)
    # - 나머지: 같은 절댓값 중복 제거 후 합산
    adj_tag_key = "csm_adjustment"
    for comp_key, vals in comp_values.items():
        if not vals:
            continue
        if comp_key == adj_tag_key:
            # 조정: 최소 ndim에서 ContractualServiceMarginMember(집계) 값이 있으면 그것 하나만,
            # 없으면 최소 ndim의 sub-component 합산
            min_nd = min(nd for nd, _ in vals)
            min_nd_vals = [(nd, v) for nd, v in vals if nd == min_nd]
            # ContractualServiceMarginMember context는 별도로 표시됨
            # 집계값(단일)과 세부값들이 같은 ndim에 있을 때 → 집계값 하나만 사용
            # 집계값 = 세부값들의 합과 같음 → 집계값이 있으면 집계값만
            total_min = sum(v for _, v in min_nd_vals)
            # 값이 2개 이상이고 그 중 하나가 전체합과 같으면 그것 하나만 사용
            for _, v in min_nd_vals:
                if abs(v - total_min / 2) < abs(total_min) * 0.01 and len(min_nd_vals) > 1:
                    # 약 절반인 값이 있으면 집계값 = 절반 → 그것만 사용
                    sums[comp_key] = v
                    break
            else:
                sums[comp_key] = total_min
        elif comp_key in ("new_csm", "csm_amortization"):
            # 신계약/상각: 최소 ndim에서 합산하되,
            # 단일 집계값이 세부값들의 합과 같으면 세부값들만 사용
            if vals:
                min_nd = min(nd for nd, _ in vals)
                min_vals = [v for nd, v in vals if nd == min_nd]
                total_min = sum(min_vals)
                # 집계값이 세부값 합과 같으면(같은 ndim에 중복 저장) → 더 작은 세부값들만 합산
                # 예: [2037561, 1386, 2036175] → 2037561 = 1386+2036175 → 세부값 합 사용
                # 하나의 값이 나머지 합과 같으면 그것은 집계값
                non_dup = []
                for v in min_vals:
                    rest_sum = sum(x for x in min_vals if x != v or min_vals.count(x) > 1)
                    if abs(abs(v) - abs(total_min - v)) < abs(v) * 0.001:
                        # v ≈ total - v → v ≈ total/2 → 중복 없음 (이 경우는 단순 합산)
                        pass
                    # 하나가 나머지 합과 같은지 체크
                if len(min_vals) > 1:
                    for i, v in enumerate(min_vals):
                        others_sum = sum(x for j, x in enumerate(min_vals) if j != i)
                        if abs(abs(v) - abs(others_sum)) < abs(v) * 0.01:
                            # v = sum of others → v is the aggregate, exclude it
                            non_dup = [x for j, x in enumerate(min_vals) if j != i]
                            break
                    else:
                        non_dup = min_vals
                else:
                    non_dup = min_vals
                sums[comp_key] = sum(non_dup)
        elif comp_key == "csm_finance":
            # 이자부리: 최소 ndim에서만 합산 (여러 ndim에 같은 값이 중복될 수 있음)
            if vals:
                min_nd = min(nd for nd, _ in vals)
                for nd, v in vals:
                    if nd == min_nd:
                        sums[comp_key] += v
        else:
            # 나머지: 단순 합산
            for _, v in vals:
                sums[comp_key] += v
    return sums


def find_xbrl_file(corp_name, report_type="사업보고서", year="2025"):
    """XBRL 파일 경로 찾기"""
    base = os.path.join("data", corp_name, report_type, year)
    files = glob(os.path.join(base, "*.xbrl"))
    return files[0] if files else None


def parse_simple_contexts(xbrl_path, target_year="2025"):
    """
    사업보고서 XBRL 파싱:
    - BS: 별도 당기말 (instant {year}-12-31, ConsolidatedAndSeparate 축만)
    - PL: 별도 당기 (duration {year}-01-01~{year}-12-31, ConsolidatedAndSeparate 축만)
    - prior_bs: 별도 전기말 (instant {prev_year}-12-31)

    반환: {"bs": {account_id: value_won}, "pl": {account_id: value_won}, "prior_bs": {...}}
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    prev_year = str(int(target_year) - 1)

    # 단순 context (ConsolidatedAndSeparate 축 1개만) 찾기
    bs_ctx_id = None  # instant 당기말
    pl_ctx_id = None  # duration
    prior_bs_ctx_id = None  # instant 전기말

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        ctx_id = ctx.get("id", "")
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        if len(members) != 1:
            continue
        dim = members[0].get("dimension", "")
        val = (members[0].text or "").strip()
        # 두 가지 Separate 축 지원: ifrs-full_ConsolidatedAndSeparate... 또는 dart-gcd_StatementInformationAxis
        is_separate_axis = (
            ("ConsolidatedAndSeparateFinancialStatements" in dim and "Separate" in val)
            or ("StatementInformationAxis" in dim and "SeparateMember" in val)
        )
        if not is_separate_axis:
            continue

        period = ctx.find(f"{{{NS_XBRLI}}}period")
        instant = period.find(f"{{{NS_XBRLI}}}instant")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")

        if instant is not None:
            if instant.text == f"{target_year}-12-31":
                if bs_ctx_id is None or "ConsolidatedAndSeparate" in ctx_id:
                    bs_ctx_id = ctx_id
            elif instant.text == f"{prev_year}-12-31":
                if prior_bs_ctx_id is None or "ConsolidatedAndSeparate" in ctx_id:
                    prior_bs_ctx_id = ctx_id
        elif start is not None and end is not None:
            if start.text == f"{target_year}-01-01" and end.text == f"{target_year}-12-31":
                if pl_ctx_id is None or "ConsolidatedAndSeparate" in ctx_id:
                    pl_ctx_id = ctx_id

    if not bs_ctx_id:
        print(f"  WARNING: BS context not found in {xbrl_path}")
    if not pl_ctx_id:
        print(f"  WARNING: PL context not found in {xbrl_path}")

    bs_data = {}
    pl_data = {}
    prior_bs_data = {}

    target_ctxs = {bs_ctx_id, pl_ctx_id, prior_bs_ctx_id} - {None}

    for elem in root:
        tag = elem.tag
        ctx_ref = elem.get("contextRef", "")
        if not elem.text or not elem.text.strip():
            continue
        if ctx_ref not in target_ctxs:
            continue

        if "}" not in tag:
            continue
        ns_uri, local = tag[1:].split("}", 1)
        if "ifrs-full" in ns_uri or "ifrs_full" in ns_uri:
            prefix = "ifrs-full"
        elif "dart" in ns_uri and "dart-gcd" not in ns_uri:
            prefix = "dart"
        else:
            continue

        account_id = f"{prefix}_{local}"
        val = elem.text.strip()
        try:
            v = float(val)
        except ValueError:
            continue

        if ctx_ref == bs_ctx_id:
            bs_data[account_id] = v
        elif ctx_ref == pl_ctx_id:
            pl_data[account_id] = v
        elif ctx_ref == prior_bs_ctx_id:
            prior_bs_data[account_id] = v

    return {"bs": bs_data, "pl": pl_data, "prior_bs": prior_bs_data}


def parse_insurance_components(xbrl_path, target_year="2025"):
    """
    보험계약부채 구성요소별 잔액 합산 (CSM, BEL, RA).
    여러 XBRL 구조 지원:
    - 패턴A (미래에셋): 5-dim (Separate+IssuedMember+InsContractsAxis+TypesOfContracts+Component), tag=InsuranceContractsThatAreLiabilities
    - 패턴B (삼성생명): 4-dim (Separate+IssuedMember+TypesOfContracts+Component), tag=InsuranceContractsThatAreLiabilities
                       + 2-dim CSM (Separate+IssuedMember), tag=ContractualServiceMargin
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    component_sums = {"CSM": 0.0, "BEL": 0.0, "RA": 0.0}

    component_map = {
        "ContractualServiceMarginMember": "CSM",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachHasBeenAppliedMember": "CSM",
        "ContractualServiceMarginNotRelatedToContractsThatExistedAtTransitionDateToWhichModifiedRetrospectiveApproachOrFairValueApproachHasBeenAppliedMember": "CSM",
        "ContractualServiceMarginRelatedToContractsThatExistedAtTransitionDateToWhichFairValueApproachHasBeenAppliedMember": "CSM",
        "EstimatesOfPresentValueOfFutureCashFlowsMember": "BEL",
        "RiskAdjustmentForNonfinancialRiskMember": "RA",
    }

    # 분류 축 이름들 (회사별로 다름)
    CLASSIFICATION_AXES = {
        "TypesOfContractsAxis",
        "TypesOfInsuranceContractsAxis",
        "ClassificationOfInsuranceProductsAxisOfDisclosureOfReconciliationOfChangesInInsuranceContractsIssuedByComponentsTableOfAxis",
        "ClassificationOfInsuranceContractsByPresenceOrAbsenceOfParticipationFeaturesAxisOfDisclosureOfDetailedInformationAboutConcentrationsOfRiskThatArisesFromContractsWithinScopeOfIfrs17OfAxis",
        "NatureOfInsuranceContractsAxisOfDisclosureOfReconciliationOfChangesInInsuranceContractsIssuedByComponentsTableOfAxis",
        "PortfolioClassificationAxisOfDisclosureOfInsuranceLiabilitiesByAccountingModelAndPortfolioOfAxis",
        "TypesOfInsuranceContractsByDividendOfDisclosureOfReconciliationOfChangesInInsuranceContractsIssuedByRemainingCoverageAndIncurredClaimsTableOfAxis",
    }
    # 분류 축이지만 중복 계층(집계 레벨)을 만드는 축들 - 제외
    EXCLUDE_AXES = {
        "InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis",
    }

    target_ctxs = {}
    csm_2dim_ctx = None  # 2-dim CSM (삼성생명 패턴)

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        ctx_id = ctx.get("id", "")
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        instant = period.find(f"{{{NS_XBRLI}}}instant")
        if instant is None or instant.text != f"{target_year}-12-31":
            continue

        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val

        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue
        if "InsuranceContractsIssuedMember" != dims.get("DisaggregationOfInsuranceContractsAxis", ""):
            continue

        component_val = dims.get("InsuranceContractsByComponentsAxis", "")

        # 공통: InsuranceContractsIssuedMember + 분류축 1개 이상 + ComponentAxis
        # dim 수는 4~6개, 반드시 분류축이 1개 이상 있어야 함
        has_classification = any(ax in dims for ax in CLASSIFICATION_AXES)

        # 중복 계층 축이 있으면 제외
        has_exclude = any(ax in dims for ax in EXCLUDE_AXES)
        if component_val in component_map and has_classification and not has_exclude:
            target_ctxs[ctx_id] = (component_map[component_val], len(dims))

        # 패턴B CSM: 2-dim (Separate + IssuedMember), tag=ContractualServiceMargin
        elif len(dims) == 2 and component_val == "":
            csm_2dim_ctx = ctx_id

    # BEL/RA/CSM 합산: 올바른 dim 레벨 자동 선택
    # InsuranceContractsThatAreLiabilities 우선, 없으면 InsuranceContractsLiabilityAsset
    PREF_TAG = "InsuranceContractsThatAreLiabilities"
    FALLBACK_TAG = "InsuranceContractsLiabilityAsset"

    # dim별, 태그별로 분리 (InsuranceContractsThatAreAssets는 절댓값으로 합산)
    ASSET_TAG = "InsuranceContractsThatAreAssets"
    by_ndim = {}  # ndim -> {"pref": {comp: sum}, "fallback": {comp: sum}}
    for ctx_ref, (comp_key, ndim) in target_ctxs.items():
        if ndim not in by_ndim:
            by_ndim[ndim] = {
                "pref": {"CSM": 0.0, "BEL": 0.0, "RA": 0.0},
                "fallback": {"CSM": 0.0, "BEL": 0.0, "RA": 0.0},
            }
        for elem in root:
            if elem.get("contextRef") != ctx_ref or not elem.text:
                continue
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            try:
                v = float(elem.text.strip())
                if tag == PREF_TAG:
                    by_ndim[ndim]["pref"][comp_key] += v
                elif tag == ASSET_TAG:
                    # Assets는 별도 누적 (나중에 pref와 합산 여부 결정)
                    by_ndim[ndim].setdefault("assets", {"CSM": 0.0, "BEL": 0.0, "RA": 0.0})
                    by_ndim[ndim]["assets"][comp_key] += abs(v)
                elif tag == FALLBACK_TAG:
                    by_ndim[ndim]["fallback"][comp_key] += v
            except ValueError:
                pass

    if by_ndim:
        # pref 태그로 합산이 있는 ndim 중 최솟값 선택
        chosen = None
        chosen_ndim = None
        for ndim in sorted(by_ndim.keys()):
            pref = by_ndim[ndim]["pref"]
            if any(pref[k] != 0 for k in ["BEL", "RA", "CSM"]):
                chosen = dict(pref)
                chosen_ndim = ndim
                break
        if chosen is None:
            for ndim in sorted(by_ndim.keys()):
                fallback = by_ndim[ndim]["fallback"]
                if any(fallback[k] != 0 for k in ["BEL", "RA", "CSM"]):
                    chosen = dict(fallback)
                    chosen_ndim = ndim
                    break

        if chosen is not None:
            assets = by_ndim.get(chosen_ndim, {}).get("assets", {"CSM": 0.0, "BEL": 0.0, "RA": 0.0})
            # 각 component별로 Assets 추가 여부를 독립적으로 결정
            # pref + assets가 정답이 되는 조건: assets < pref * 6% (DB손보 CSM 패턴)
            # pref가 이미 정답인 경우: assets 제외
            for key in ["BEL", "RA", "CSM"]:
                pref_v = chosen[key]
                asset_v = assets.get(key, 0.0)
                if asset_v == 0 or pref_v == 0:
                    component_sums[key] = pref_v
                elif abs(asset_v) < abs(pref_v) * 0.06:
                    # assets가 pref의 6% 미만 → 별도 항목 (DB손보 CSM: 667924/11537419≈0.058)
                    component_sums[key] = pref_v + asset_v
                else:
                    # assets가 더 큰 비율 → 중복 제외
                    component_sums[key] = pref_v

    # 패턴B CSM: ContractualServiceMargin 태그 (2-dim context)
    # 이미 패턴A/B에서 CSM을 합산했으면 이 값은 무시 (중복 방지)
    if component_sums["CSM"] == 0.0 and csm_2dim_ctx:
        for elem in root:
            if elem.get("contextRef") != csm_2dim_ctx or not elem.text:
                continue
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag == "ContractualServiceMargin":
                try:
                    component_sums["CSM"] = float(elem.text.strip())
                except ValueError:
                    pass

    # 패턴B CSM fallback: 2-dim (Separate+IssuedMember) ContractualServiceMargin 합산
    if component_sums["CSM"] == 0.0:
        for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
            ctx_id = ctx.get("id", "")
            period = ctx.find(f"{{{NS_XBRLI}}}period")
            instant = period.find(f"{{{NS_XBRLI}}}instant")
            if instant is None or instant.text != f"{target_year}-12-31":
                continue
            members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
            dims = {}
            for m in members:
                dim_name = m.get("dimension", "").split(":")[-1]
                val = (m.text or "").split(":")[-1]
                dims[dim_name] = val
            if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
                continue
            if "InsuranceContractsIssuedMember" != dims.get("DisaggregationOfInsuranceContractsAxis", ""):
                continue
            if len(dims) != 2:
                continue
            for elem in root:
                if elem.get("contextRef") != ctx_id or not elem.text:
                    continue
                tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                if tag == "ContractualServiceMargin":
                    try:
                        v = float(elem.text.strip())
                        if v != 0:
                            component_sums["CSM"] = v
                    except ValueError:
                        pass

    return component_sums


def parse_new_csm(xbrl_path, target_year="2025"):
    """
    신계약 CSM 추출: 별도 + duration + InsuranceContractsIssuedMember + ContractualServiceMarginMember
    ChangesInFutureServicesDueToNewContracts... 태그
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        ctx_id = ctx.get("id", "")
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        if start is None or end is None:
            continue
        if start.text != f"{target_year}-01-01" or end.text != f"{target_year}-12-31":
            continue

        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val

        if "SeparateMember" != dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue
        if "ContractualServiceMarginMember" not in dims.get("InsuranceContractsByComponentsAxis", ""):
            continue
        if "InsuranceContractsIssuedMember" not in dims.get("DisaggregationOfInsuranceContractsAxis", ""):
            continue
        if len(dims) != 3:
            continue

        # 이 context의 신계약CSM 태그 찾기
        for elem in root:
            if elem.get("contextRef") != ctx_id or not elem.text:
                continue
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if "ChangesInFutureServicesDueToNewContracts" in tag:
                try:
                    return float(elem.text.strip())
                except ValueError:
                    pass
    return None


def parse_surrender_reserve(xbrl_path, target_year="2025"):
    """해약환급금준비금 추출.
    패턴A(생보): 2-dim (Separate + SurrenderValueReserveMember), tag=LoanLossReserveBalance 등
    패턴B(손보/삼성생명): 1-dim (Separate), tag=SurrenderValueReserve, instant 당기말
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()

    best_a = None   # 패턴A (ReservesWithinEquityAxis)
    best_b = None   # 패턴B SurrenderValueReserve
    best_b_added = None  # SurrenderValueReserveToBeAdded (당기 추가분)

    for ctx in root.findall(f"{{{NS_XBRLI}}}context"):
        ctx_id = ctx.get("id", "")
        period = ctx.find(f"{{{NS_XBRLI}}}period")
        members = ctx.findall(f".//{{{NS_XBRLDI}}}explicitMember")
        dims = {}
        for m in members:
            dim_name = m.get("dimension", "").split(":")[-1]
            val = (m.text or "").split(":")[-1]
            dims[dim_name] = val

        if "SeparateMember" not in dims.get("ConsolidatedAndSeparateFinancialStatementsAxis", ""):
            continue

        instant = period.find(f"{{{NS_XBRLI}}}instant")
        start = period.find(f"{{{NS_XBRLI}}}startDate")
        end = period.find(f"{{{NS_XBRLI}}}endDate")
        is_cur = (
            (instant is not None and instant.text == f"{target_year}-12-31") or
            (start is not None and start.text == f"{target_year}-01-01" and
             end is not None and end.text == f"{target_year}-12-31")
        )
        if not is_cur:
            continue

        has_reserve_axis = "SurrenderValueReserveMember" in dims.get("ReservesWithinEquityAxis", "")
        is_1dim_instant = (len(dims) == 1 and instant is not None)

        if not has_reserve_axis and not is_1dim_instant:
            continue

        for elem in root:
            if elem.get("contextRef") != ctx_id or not elem.text:
                continue
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            try:
                v = float(elem.text.strip())
                if has_reserve_axis and "LoanLossReserveBalance" in tag:
                    if v > 0: best_a = v
                elif is_1dim_instant and tag == "SurrenderValueReserve":
                    if v > 0: best_b = v  # 양수만
                elif is_1dim_instant and tag == "SurrenderValueReserveToBeAdded":
                    best_b_added = v  # 음수(차감)도 허용
                elif has_reserve_axis and ("SurrenderValueReserve" in tag or "LoanLossReserve" in tag):
                    if best_a is None:
                        best_a = v
            except ValueError:
                pass

    # Col99 = SurrenderValueReserve + SurrenderValueReserveToBeAdded (손보사는 두 값의 합이 정답)
    if best_a:
        return best_a
    if best_b and best_b_added:
        return best_b + best_b_added
    return best_b or best_b_added


def extract_company_data(corp_name, year="2025"):
    """한 회사의 사업보고서에서 DATA 시트에 필요한 모든 값 추출"""
    xbrl_path = find_xbrl_file(corp_name, "사업보고서", year)
    if not xbrl_path:
        print(f"  ERROR: XBRL not found for {corp_name}")
        return None

    print(f"  파싱: {xbrl_path}")
    result = {}

    # 단순 context (BS/PL)
    simple = parse_simple_contexts(xbrl_path, year)
    bs = simple["bs"]
    pl = simple["pl"]

    def _get_bs(account_id):
        v = bs.get(account_id)
        if v is not None:
            return v
        alt = account_id.replace("dart_", "ifrs-full_") if account_id.startswith("dart_") else account_id.replace("ifrs-full_", "dart_")
        return bs.get(alt)

    # BS 매핑
    for col, account_id in BS_MAPPING.items():
        val = _get_bs(account_id)
        if val is not None:
            result[col] = val / 1e6  # 원 → 백만원

    # Col6 추가 fallback (신한라이프 패턴: SecuritiesAtFairValueThroughOCI)
    if 6 not in result:
        v = bs.get("dart_SecuritiesAtFairValueThroughOtherComprehensiveIncome")
        if v is not None:
            result[6] = v / 1e6

    # Col8 추가 fallback (InvestmentAccountedForUsingEquityMethod - 신한라이프 패턴)
    if 8 not in result:
        v = bs.get("ifrs-full_InvestmentAccountedForUsingEquityMethod")
        if v is not None:
            result[8] = v / 1e6

    # PL 매핑
    for col, account_id in PL_MAPPING.items():
        val = pl.get(account_id)
        if val is not None:
            result[col] = val / 1e6
        else:
            # ifrs-full과 dart prefix 둘 다 시도
            alt_id = account_id.replace("dart_", "ifrs-full_") if account_id.startswith("dart_") else account_id.replace("ifrs-full_", "dart_")
            val = pl.get(alt_id)
            if val is not None:
                result[col] = val / 1e6

    def _get_pl(account_id):
        """PL에서 dart_/ifrs-full_ 둘 다 시도"""
        v = pl.get(account_id)
        if v is not None:
            return v
        alt = account_id.replace("dart_", "ifrs-full_") if account_id.startswith("dart_") else account_id.replace("ifrs-full_", "dart_")
        return pl.get(alt, 0)

    # 계산 항목
    for col, (id1, id2, op) in CALC_MAPPING.items():
        if op == "direct":
            val = _get_pl(id1)
            if val:
                result[col] = val / 1e6
        elif op == "subtract":
            v1 = _get_pl(id1)
            v2 = _get_pl(id2)
            if v1 or v2:
                result[col] = (v1 - v2) / 1e6

    # Col56: 2-dim (Separate + ReportedAmountMember) context 우선
    fvoci_v = parse_fvoci_oci(xbrl_path, year)
    if fvoci_v:
        result[56] = fvoci_v / 1e6

    # OCI 매핑 (태그 후보 리스트) - Col56은 이미 처리됨
    for col, candidates in OCI_MAPPING.items():
        if col in result:
            continue
        for account_id in candidates:
            val = _get_pl(account_id)
            if val:
                result[col] = val / 1e6
                break

    # OCI 합산 매핑
    for col, (id1, id2) in OCI_SUM_MAPPING.items():
        v1 = _get_pl(id1)
        v2 = _get_pl(id2)
        if v1 or v2:
            result[col] = (v1 + v2) / 1e6

    # 전기말 BS 매핑
    prior_bs = simple["prior_bs"]
    prior_hybrid = 0.0
    for col, account_id in PRIOR_BS_MAPPING.items():
        if isinstance(col, str):
            # Special keys (e.g. "_prior_hybrid")
            v = prior_bs.get(account_id)
            if v is None:
                alt = account_id.replace("dart_", "ifrs-full_") if account_id.startswith("dart_") else account_id.replace("ifrs-full_", "dart_")
                v = prior_bs.get(alt)
            if v is not None and col == "_prior_hybrid":
                prior_hybrid = v
            continue
        val = prior_bs.get(account_id)
        if val is None:
            alt = account_id.replace("dart_", "ifrs-full_") if account_id.startswith("dart_") else account_id.replace("ifrs-full_", "dart_")
            val = prior_bs.get(alt)
        if val is not None:
            result[col] = val / 1e6

    # Col7 = AC 자산 (상각후원가 금융자산)
    ac_fin = _get_bs("ifrs-full_FinancialAssetsAtAmortisedCost") or 0
    loans_ac = _get_bs("ifrs-full_LoansAtAmortisedCost") or 0
    sec_ac = _get_bs("ifrs-full_SecuritiesAtAmortisedCost") or 0
    other_ac = (_get_bs("ifrs-full_OtherFinancialAssetsAtAmortisedCost") or
                _get_bs("dart_OtherFinancialAssetsAtAmortisedCost") or 0)

    if ac_fin > 0:
        if loans_ac > ac_fin:
            # loans가 ac_fin보다 크면 loans+sec+other로 합산
            result[7] = (loans_ac + sec_ac + other_ac) / 1e6
        elif loans_ac == 0 and sec_ac == 0 and other_ac > 0:
            # 메리츠화재 패턴: FinancialAssets + OtherFinancialAssets 별도
            result[7] = (ac_fin + other_ac) / 1e6
        else:
            result[7] = ac_fin / 1e6
    elif sec_ac > 0:
        if sec_ac >= loans_ac:
            result[7] = sec_ac / 1e6  # 신한라이프: 대출채권(보험계약대출) 제외
        else:
            result[7] = (loans_ac + sec_ac + other_ac) / 1e6
    elif loans_ac > 0:
        result[7] = (loans_ac + other_ac) / 1e6

    # Col14 = 자본(신종자본증권 제외) = Col15 - Col16
    if 15 in result:
        result[14] = result[15] - result.get(16, 0)

    # Col19 = 기초자본 (전기말 Equity) → overwrite PL_MAPPING value
    if 19 in result and result[19] == 0:
        # PL에서 못 찾은 경우만 prior_bs 사용
        pass
    # Col20 = 기초자본(신종자본증권 제외) = Col19 - 전기말 HybridBonds
    if 19 in result:
        result[20] = result[19] - prior_hybrid / 1e6

    # 비율 계산
    if 4 in result and result[4] > 0:
        if 5 in result:
            result[9] = result[5] / result[4]  # FVPL 비율
        if 6 in result:
            result[10] = result[6] / result[4]  # FVOCI 비율
        if 7 in result:
            result[11] = result[7] / result[4]  # AC 비율

    # 보험계약부채 구성요소 (CSM/BEL/RA)
    components = parse_insurance_components(xbrl_path, year)
    if components["CSM"] > 0:
        result[75] = components["CSM"] / 1e6
    if components["BEL"] > 0:
        result[73] = components["BEL"] / 1e6
    if components["RA"] > 0:
        result[74] = components["RA"] / 1e6

    # 신계약 CSM
    new_csm = parse_new_csm(xbrl_path, year)
    if new_csm is not None:
        result[76] = new_csm / 1e6

    # 해약환급금준비금
    surr = parse_surrender_reserve(xbrl_path, year)
    if surr is not None:
        result[99] = surr / 1e6

    # 추가 PL 항목 (중복 컬럼)
    for col, account_id in EXTRA_PL_MAPPING.items():
        if col not in result:
            val = _get_pl(account_id)
            if val:
                result[col] = val / 1e6

    # Col40: 기타투자손익 = FeeIncome + RentalIncome + OtherIncomeInvestment - OtherExpenseInvestment
    def _get_pl_multi(candidates):
        for aid in candidates:
            v = _get_pl(aid)
            if v:
                return v
        return 0

    def _get_pl_by_keyword(keywords):
        """PL에서 tag 이름에 keyword가 포함된 값들 합산"""
        total = 0.0
        for k, v in pl.items():
            for kw in keywords:
                if kw in k:
                    # 단순 OperatingExpenseInvestment(전체) 태그 제외
                    if k in ("dart_OperatingExpenseInvestment", "ifrs-full_OperatingExpenseInvestment"):
                        continue
                    # 이익/수익/관련자 거래 태그 제외
                    if any(x in k for x in ("Gain", "Income", "Revenue", "RelatedParty",
                                             "DisclosureOfTransactions")):
                        continue
                    # udf_ 계열 OfOperatingExpenseInvestment는 포함 (DB손보 패턴)
                    total += v
                    break
            # udf_ 계열 OfOperatingExpenseInvestment는 OtherOperatingExpenseInvestment와 함께 존재할 때만 포함
            # (동일 회사 내에서 중복 방지)
        return total

    fee = _get_pl_multi(COL40_FEE)
    rent = _get_pl_multi(COL40_RENT)
    other_inc = _get_pl_multi(COL40_OTHER_INC)
    other_exp = _get_pl_by_keyword(COL40_OTHER_EXP_KEYWORDS)
    col40_val = fee + rent + other_inc - other_exp
    if col40_val:
        result[40] = col40_val / 1e6

    # Col38: 금융손익 = Col35 - Col36 - Col37 + Col39 - Col40
    if all(c in result for c in [35, 36, 37, 39, 40]):
        result[38] = result[35] - result[36] - result[37] + result[39] - result[40]

    # Col51: 세전이익 합계 (Col43과 같음)
    if 43 in result:
        result[51] = result[43]

    # OCI 누계 항목 (당기말/전기말 BS 3-dim context)
    oci_accum = parse_oci_accumulated(xbrl_path, year)
    oci_cur = oci_accum["cur"]
    oci_prior = oci_accum["prior"]
    for col, keywords in OCI_ACCUM_MAPPING.items():
        # Col29 재평가잉여금은 전기말
        src = oci_prior if col == 29 else oci_cur
        total = 0.0
        for member, v in src.items():
            if any(kw in member for kw in keywords):
                total += v
        if total != 0.0:
            result[col] = total / 1e6

    # Col24 = 기타포괄손익누계액 합계 (= Col17, AccumulatedOCI)
    if 17 in result:
        result[24] = result[17]

    # Col63: 기말 기타포괄손익누계액 (당기말 = Col17)
    if 17 in result:
        result[63] = result[17]

    # Col71: 이익잉여금 기말 (당기말 BS = Col18)
    if 18 in result:
        result[71] = result[18]

    # CSM 변동 항목
    csm_mov = parse_csm_movement(xbrl_path, year)
    if csm_mov["csm_amortization"]:
        result[77] = abs(csm_mov["csm_amortization"]) / 1e6  # 항상 양수로
    if csm_mov["csm_adjustment"]:
        result[78] = csm_mov["csm_adjustment"] / 1e6
    if csm_mov["csm_finance"]:
        result[79] = csm_mov["csm_finance"] / 1e6

    # Col76 fallback: parse_new_csm이 못 찾으면 csm_movement의 new_csm 사용
    if 76 not in result and csm_mov.get("new_csm"):
        result[76] = abs(csm_mov["new_csm"]) / 1e6

    # Col80: 전기말 CSM
    prior_csm = parse_prior_csm(xbrl_path, year)
    if prior_csm:
        result[80] = prior_csm / 1e6

    # Col88: CSM 합계 (= Col75)
    if 75 in result:
        result[88] = result[75]

    # 계리적 가정 변경 민감도 분석 (Col146~158)
    actuarial = parse_actuarial_sensitivity(xbrl_path, year)
    for (col, comp), v in actuarial.items():
        result[col] = v / 1e6

    # 계산값 (다른 컬럼 값들의 조합)
    # Col31: OCI 체크 = Col24 - (Col25+Col26+Col27+Col28+Col29+Col30)
    # OCI 세부항목이 모두 있을 때만 계산 (없으면 정확한 값 계산 불가)
    oci_parts_cols = [25, 26, 27, 28, 29, 30]
    if 24 in result and all(c in result for c in oci_parts_cols):
        result[31] = result[24] - sum(result[c] for c in oci_parts_cols)

    # Col45: 영업손익 체크 = Col41+Col42-Col43
    if all(c in result for c in [41, 42, 43]):
        result[45] = result[41] + result[42] - result[43]

    # Col46: 투자손익 체크 = Col36+Col37+Col38-Col39+Col40-Col35
    if all(c in result for c in [35, 36, 37, 38, 39, 40]):
        result[46] = result[36] + result[37] + result[38] - result[39] + result[40] - result[35]

    # Col47: 보험투자손익 체크 = Col32+Col35-Col41
    if all(c in result for c in [32, 35, 41]):
        result[47] = result[32] + result[35] - result[41]

    # Col64: OCI 체크 = Col63-Col55-SUM(Col56~62)
    oci_detail = [result.get(c, 0) for c in range(56, 63)]
    if 63 in result and 55 in result and all(c in result for c in range(56, 62)):
        result[64] = result[63] - result[55] - sum(oci_detail)

    # Col89: CSM 합계 체크 = Col88-SUM(Col83~87)
    csm_detail = [result.get(c, 0) for c in range(83, 88)]
    if 88 in result and all(c in result for c in range(83, 88)):
        result[89] = result[88] - sum(csm_detail)

    # Col82: CSM 증감 = Col80+Col76-Col77+Col78+Col79-Col75
    if all(c in result for c in [75, 76, 77, 78, 79, 80]):
        val82 = result[80] + result[76] - result[77] + result[78] + result[79] - result[75]
        result[82] = val82  # 0이어도 저장 (체크값)

    return result


def load_kics_data(csv_path="data/kics_2512.csv"):
    """K-ICS CSV 로드: {company_short: {col: value}}
    Col92~94: 당기(2026.03 기준), Col95~97: 전기(2025.12 기준)
    """
    kics = {}
    if not os.path.exists(csv_path):
        return kics
    with open(csv_path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            short = row.get("회사", "")
            try:
                d = {}
                # 전기 K-ICS (Col95~97) - 2512말 기준
                if row.get("킥스비율"): d[95] = float(row["킥스비율"]) / 100  # % → 소수
                if row.get("가용자본"): d[96] = float(row["가용자본"])
                if row.get("요구자본"): d[97] = float(row["요구자본"])
                # Col92~94는 회사마다 다른 기준이라 MISS로 남김
                if d:
                    kics[short] = d
            except (ValueError, TypeError):
                pass
    return kics


def main():
    print("=" * 60)
    print("사업보고서 XBRL → DATA 시트 매핑 추출")
    print("=" * 60)

    kics_data = load_kics_data()

    all_results = {}
    for comp in COMPANIES:
        name = comp["name"]
        short = comp["short"]
        print(f"\n[{name}] ({short})")
        data = extract_company_data(name, "2025")
        if data:
            key = f"2512{short}"
            # Col2: 기간코드
            data[2] = 2512
            # K-ICS 데이터 추가
            if short in kics_data:
                for col, val in kics_data[short].items():
                    if val:
                        data[col] = val
            all_results[key] = data
            print(f"  추출 항목: {len(data)}개")
            if 4 in data:
                print(f"  자산: {data[4]:,.2f}M")
            if 32 in data:
                print(f"  보험손익: {data[32]:,.2f}M")
            if 44 in data:
                print(f"  당기순이익: {data[44]:,.2f}M")

    # CSV 출력
    output_path = "data/data_sheet_2512.csv"
    os.makedirs("data", exist_ok=True)

    all_cols = sorted(set(col for data in all_results.values() for col in data.keys()))
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["key"] + [f"Col{c}" for c in all_cols])
        for key, data in all_results.items():
            row = [key] + [data.get(c, "") for c in all_cols]
            writer.writerow(row)

    print(f"\n저장: {output_path}")
    print(f"회사 수: {len(all_results)}, 열 수: {len(all_cols)}")


if __name__ == "__main__":
    main()
