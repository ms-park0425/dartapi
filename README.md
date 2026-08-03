# 동업사 공시비교 DATA 시트 자동화

DART 전자공시시스템에서 12개 보험사의 XBRL 파일과 원문 XML(document.xml)을 받아  
`DATA_작업_빈칸.xlsx`의 DATA 시트를 자동으로 채웁니다.

## 실행 방법

```bash
# 전체 기간 일괄 처리 (Excel에 있는 모든 기간코드)
python parse_annual.py --period all

# 특정 기간만
python parse_annual.py --period 2512    # 2025년 12월 사업보고서
python parse_annual.py --period 2506    # 2025년 6월 반기보고서
python parse_annual.py --period 2412    # 2024년 12월 사업보고서
```

**XBRL / document.xml이 없으면 DART API에서 자동 다운로드합니다.**  
결과는 `data/` 폴더에 캐시되어 다음 실행 시 재사용됩니다.

---

## XBRL Taxonomy

`data/taxonomy.csv` — XBRL 태그 ID와 한글명 매핑표. 각 회사 XBRL 다운로드 시 포함된 `*_lab-ko.xml`에서 자동 추출됩니다.

| 분류 | prefix | 개수 | 설명 |
|------|--------|-----:|------|
| IFRS 표준 | `ifrs-full_` | 1,562 | IASB가 정의한 국제 표준 태그 |
| DART 한국 추가 | `dart_` | 967 | 금융감독원이 보험업 특화로 추가 정의 |
| 회사별 커스텀 | `entity{corp_code}_` | 10,844 | 각 회사가 자체 정의한 태그 (비표준) |

> **표준 taxonomy**는 `ifrs-full_`과 `dart_` prefix인 2,529개입니다.  
> `entity` prefix 태그는 특정 회사만 사용하는 커스텀 태그로, 다른 회사/연도에 적용되지 않습니다.

**주요 활용 태그 예시:**

| 항목 | 태그 |
|------|------|
| 총자산 | `ifrs-full_Assets` |
| 보험계약부채 | `ifrs-full_InsuranceContractsIssuedThatAreLiabilities` |
| 보험손익 | `ifrs-full_InsuranceServiceResult` |
| 당기순이익 | `ifrs-full_ProfitLoss` |
| 보험금융손익 | `dart_InsuranceFinanceIncomeFromInsuranceContractsIssuedRecognisedInProfitOrLoss` |
| 신종자본증권 | `dart_HybridBonds` |
| 보험계약마진(CSM) | `dart_ContractualServiceMargin` 등 |

---

## 컬럼별 자동화 현황

> **기준: 2512 (2025년 12월 사업보고서), 12개사**  
> 분기/반기 보고서(2503·2506·2509)는 BS·PL 기본 항목 위주로 추출되며, 사업보고서 전용 항목(CSM변동·K-ICS·예실차 등)은 연간 보고서에서만 추출됩니다.  
> 전년도(2312·2412)는 전체적으로 유사하나, 일부 XBRL 태그 구조 차이로 정확도가 낮을 수 있습니다.

아래에서 **소스**는 데이터를 어디서 가져오는지를 나타냅니다.

- **XBRL**: `fnlttXbrl.xml` API로 받은 `.xbrl` 파일 (별도 재무제표 기준)
- **원문XML**: `document.xml` API로 받은 보고서 원문 HTML 테이블 파싱
- **계산**: 다른 컬럼 값으로부터 파생

---

### B/S (재무상태표)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 4 | 총자산 | XBRL | `ifrs-full_Assets` | ✅ 12/12 |
| 5 | FVPL | XBRL | `ifrs-full_FinancialAssetsAtFairValueThroughProfitOrLoss` | ✅ 12/12 |
| 6 | FVOCI | XBRL | `ifrs-full_FinancialAssetsAtFairValueThroughOtherComprehensiveIncome` | ✅ 12/12 |
| 7 | AC | XBRL | `ifrs-full_FinancialAssetsAtAmortisedCost` | ✅ 12/12 |
| 8 | 종속·관계기업 | XBRL | `ifrs-full_InvestmentsInSubsidiariesJointVenturesAndAssociates` | ✅ 11/11 |
| 9~11 | 총자산대비비중 | 계산 | Col5~7 / Col4 | ✅ 계산 정확. 정답 Excel의 KB손보 값이 이전 분기 수치를 그대로 복사한 것으로 보임 |
| 12 | 부채 | XBRL | `ifrs-full_Liabilities` | ✅ 12/12 |
| 13 | 보험계약부채 | XBRL | `ifrs-full_InsuranceContractsIssuedThatAreLiabilities` | ✅ 12/12 |
| 14 | 자본(신종제외) | 계산 | Col15 − Col16 | ✅ 12/12 |
| 15 | 자본 | XBRL | `ifrs-full_Equity` | ✅ 12/12 |
| 16 | 신종자본증권 | XBRL | `dart_HybridBonds` | ✅ 5/5 (발행 회사만) |
| 17 | OCI누계액 | XBRL | `ifrs-full_AccumulatedOtherComprehensiveIncome` | ✅ 12/12 |
| 18 | 이익잉여금 | XBRL | `ifrs-full_RetainedEarnings` | ✅ 12/12 |

---

### 자산운용

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 21 | 자산운용 총자산 | 원문XML | 지급여력 테이블의 `총 자 산(A)` | ⚠️ 삼성 MISS (doc.xml에 없음) |
| 22 | 운용자산 | 원문XML | 운용자산/자산운용률 테이블 `운용자산(B)` | ⚠️ 삼성·삼성화재·현대해상·KB손보 MISS |
| 23 | 자산운용률 | 원문XML | 운용자산/자산운용률 테이블 `자산운용률(B/A)` | ⚠️ 삼성·삼성화재·현대해상·KB손보 MISS |

> **정의**: 보험업감독업무시행세칙 기준 — 특별계정자산 제외 총자산(Col21), 비운용자산 제외 운용자산(Col22)  
> **MISS 이유**: 삼성생명·삼성화재·현대해상·KB손보는 사업보고서 원문 XML에 해당 테이블 미포함.  
> XBRL 총자산에는 특별계정자산이 포함되어 있어 단순 태그 매핑으로는 추출 불가.  
> → 감독원 업무보고서 또는 회사 IR 자료에서 수동 확인 필요.

---

### 기타포괄손익누계액 (OCI 잔액)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 24 | OCI 합계 | 계산 | = Col17 | ✅ 12/12 |
| 25 | FVOCI 평가손익 | XBRL→원문XML | XBRL 재무상태표 기타포괄손익 항목 중 FVOCI 금융자산 평가손익 잔액. fallback: 사업보고서 자본 세부 테이블의 `공정가치측정금융자산평가손익` 행 | ⚠️ 신한라이프·KB라이프·DB손보 MISS |
| 26 | 보험계약 금융손익 | XBRL→원문XML | XBRL 재무상태표 보험계약 관련 OCI 잔액. fallback: 자본 세부 테이블의 `보험계약 관련 금융손익` 행 | ⚠️ 5개사 MISS |
| 27 | 재보험계약 금융손익 | XBRL→원문XML | XBRL 재무상태표 재보험계약 관련 OCI 잔액. fallback: 자본 세부 테이블의 `재보험계약 관련 금융손익` 행 | ⚠️ 8개사 MISS |
| 28 | 현금흐름위험회피 | XBRL→원문XML | XBRL 재무상태표 현금흐름위험회피 파생상품 OCI 잔액. fallback: 자본 세부 테이블의 `현금흐름위험회피` 행 | ⚠️ 6개사 MISS |
| 29 | 재평가잉여금 | XBRL→원문XML | XBRL 재무상태표 **전기말** 재평가잉여금 잔액. fallback: 자본 세부 테이블의 `재평가잉여금` 행 | ⚠️ 3개사 MISS |
| 30 | 확정급여부채 재측정 | XBRL→원문XML | XBRL 재무상태표 확정급여부채 재측정요소 OCI 잔액. fallback: 자본 세부 테이블의 `확정급여채무 재측정요소` 행 | ⚠️ 4개사 MISS |
| 31 | CHECK | 계산 | Col24 − (Col25+…+Col30) | ⚠️ 세부항목 모두 있을 때만 |

> **MISS 이유**: XBRL 재무상태표에서 OCI 세부항목이 회사마다 다른 방식으로 구분·보고되어 일부 회사는 자동 매핑 불가.  
> 이런 경우 사업보고서 원문 XML의 자본변동표 자본 구성 테이블에서 추출하나, 그것도 없는 회사는 수동 입력 필요.

---

### 손익계산서 (P/L)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 32 | 보험손익 | XBRL | `ifrs-full_InsuranceServiceResult` | ✅ 12/12 |
| 33 | 보험서비스수익 | XBRL | `dart_OperatingIncomeInsurance` | ✅ 12/12 |
| 34 | 보험서비스비용 | XBRL | `dart_OperatingExpenseInsurance` | ✅ 12/12 |
| 35 | 투자손익 | XBRL | `dart_InvestmentIncomeExpenses` | ✅ 12/12 |
| 36 | 보험금융손익 | XBRL | `dart_InsuranceFinanceIncome` − `dart_InsuranceFinanceExpenses` | ✅ 12/12 |
| 37 | 재보험금융손익 | XBRL | `dart_FinanceIncomeFromReinsurance` − `dart_FinanceExpensesFromReinsurance` | ⚠️ 교보 오차 (income-only 구조) |
| 38 | 금융손익 | 계산 | Col35 − Col36 − Col37 + Col39 − Col40 | ⚠️ 삼성화재·DB손보 소오차, KB라이프 MISS |
| 39 | 재산관리비 | XBRL | `ifrs-full_SellingGeneralAndAdministrativeExpense` | ✅ 11/11 |
| 40 | 기타투자손익 | XBRL | 수수료+임대료+기타수익−기타비용 합산 | ⚠️ 삼성화재·DB손보 소오차 |
| 41 | 영업이익 | XBRL | `ifrs-full_ProfitLossFromOperatingActivities` | ✅ 12/12 |
| 42 | 영업외손익 | XBRL | `ifrs-full_NonOperatingProfitLoss` (없으면 Col43−Col41) | ✅ 12/12 |
| 43 | 세전손익 | XBRL | `ifrs-full_ProfitLossBeforeTax` | ✅ 12/12 |
| 44 | 당기순이익 | XBRL | `ifrs-full_ProfitLoss` | ✅ 12/12 |
| 45~47 | CHECK 1~3 | 계산 | 영업/투자/보험+투자 체크 | ✅ 12/12 |

---

### 자본변동표 — OCI 변동분

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 55 | OCI 기초잔액 | XBRL | 전기말 `ifrs-full_AccumulatedOtherComprehensiveIncome` | ✅ 12/12 |
| 56 | FVOCI 평가손익 변동 | XBRL | `OtherComprehensiveIncomeNetOfTax...FVOCI...` | ⚠️ 동양·현대해상 소오차 |
| 57 | 대손충당금 변동 | XBRL | `OtherComprehensiveIncomeNetOfTaxCreditLosses...` | ❌ 전 회사 미추출 (교보 1개사만 해당, 금액 소액) |
| 58 | 보험계약금융손익 OCI | XBRL | 보험OCI + 재보험OCI 합산 | ✅ 12/12 |
| 59 | 현금흐름위험회피 변동 | XBRL | `OtherComprehensiveIncomeNetOfTaxCashFlowHedges` | ⚠️ 교보 소오차 |
| 60 | 재평가잉여금 변동 | XBRL | `OtherComprehensiveIncomeNetOfTaxGainsLossesOnRevaluation` | ⚠️ 동양·현대해상·DB손보 오차 |
| 61 | 확정급여부채 재측정 | XBRL | `OtherComprehensiveIncomeNetOfTax...RemeasurementsOfDefinedBenefitPlans` | ✅ 12/12 |
| 62 | 해외사업환산손익 | XBRL | `OtherComprehensiveIncomeNetOfTaxExchangeDifferencesOnTranslation` | ⚠️ 동양 MISS |
| 63 | OCI 기말잔액 | 계산 | = Col17 | ✅ 12/12 |
| 64 | CHECK | 계산 | Col63 − Col55 − (Col56~62) 합산 | ⚠️ Col57 없는 회사는 오차 가능 |

---

### 자본변동표 — 이익잉여금

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 65 | 이익잉여금 기초 | XBRL | 전기말 `ifrs-full_RetainedEarnings` | ⚠️ 메리츠화재 소오차 |
| 66 | 결산배당 | XBRL | `dart_DividendsPaidClassifiedAsFinancingActivities` | ✅ 8/8 (배당 회사만) |
| 67 | 자기주식 소각/처분 | XBRL | `ifrs-full_CancellationOfTreasuryShares` (RetainedEarnings component) | ✅ 2/2 (해당 회사만) |
| 68 | 순이익 | 계산 | = Col44 | ✅ 12/12 |
| 69 | 기타 자본변동 | XBRL | `dart_TransferOfAmountRecognised...` (RetainedEarnings component) | ⚠️ 동양 소오차, KB손보 오류 |
| 70 | 신종자본증권 배당 | XBRL | `dart_DividendToHybridBond` (RetainedEarnings component) | ⚠️ 메리츠화재 소오차, 현대해상 MISS |
| 71 | 이익잉여금 기말 | 계산 | = Col18 | ✅ 12/12 |

---

### 보험계약부채 구성요소

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 73 | BEL | XBRL → 원문XML | XBRL 보험계약 주석의 BEL 구성요소 합계. 교보·신한라이프·KB생명·메리츠·KB손보는 XBRL에 BEL/RA/CSM 분리 데이터 없어 1분기보고서 원문XML의 구성요소 테이블에서 파싱 | ✅ 12/12 |
| 74 | RA | XBRL → 원문XML | 동일 | ✅ 12/12 |
| 75 | CSM | XBRL → 원문XML | 동일 | ✅ 12/12 |
| 76 | 신계약CSM | XBRL → 원문XML | XBRL 보험계약 변동 주석의 신계약 CSM 항목. fallback: 사업보고서 원문XML의 CSM 변동표 `신계약` 행 | ⚠️ 동양 MISS |
| 77 | CSM 상각 | XBRL → 원문XML | XBRL 보험계약 변동 주석의 서비스이전 상각 항목. fallback: 원문XML CSM 변동표 `보험계약마진 상각` 행 | ⚠️ 동양 MISS |
| 78 | CSM 조정 | XBRL → 원문XML | XBRL 보험계약 변동 주석의 추정치 변동 조정 항목. fallback: 원문XML CSM 변동표 `추정치 변동` 행 | ✅ 12/12 |
| 79 | 보험금융손익(CSM) | XBRL → 원문XML | XBRL 보험계약 변동 주석의 보험금융손익 항목. fallback: 원문XML CSM 변동표 `보험금융손익` 행 | ⚠️ DB손보 MISS |
| 80 | 전기말 CSM | XBRL → 원문XML | XBRL 전기말 보험계약부채의 CSM 잔액. fallback: 원문XML CSM 변동표 `기초` 행 | ⚠️ 삼성화재·DB손보 MISS |
| 82 | CHECK | 계산 | Col80+76−77+78+79−75 | ⚠️ 구성항목 일부 MISS인 경우 오차 |

---

### CSM 기간별 기대수익인식금액

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 83~87 | 1년이하~10년초과 | 원문XML | 사업보고서 CSM 만기 테이블 파싱 | ⚠️ 교보·한화·동양·DB손보·삼성화재 OK / 나머지 7개사 MISS |
| 88 | 합계 | 계산 | = Col75 | ✅ 11/11 |
| 89 | CHECK | 계산 | Col88 − (Col83~87) | ⚠️ Col83~87이 MISS인 7개사는 계산 불가 |

> **MISS 이유**: 신한라이프·삼성·미래에셋·KB라이프·현대해상·메리츠·KB손보의 사업보고서 원문 XML에  
> CSM 기간별 테이블이 없음. 해당 데이터는 다른 경로(IR 자료 등) 에서 수동 입력 필요.

---

### K-ICS (지급여력비율)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 92~94 | 잠정치 (비율·가용·요구) | 원문XML | 사업보고서 지급여력비율 테이블 | ⚠️ 삼성·KB라이프·삼성화재·현대해상·KB손보 MISS |
| 95~97 | 최종치 (비율·가용·요구) | 원문XML | 지급여력비율 테이블 당기값 | ⚠️ 삼성·KB라이프·삼성화재·현대해상·KB손보·한화 MISS (산출중) |

> **MISS 이유**: 사업보고서 제출 시점에 K-ICS가 아직 `산출중(-)` 상태인 회사 다수.  
> → 해당 회사들은 이후 별도 공시 또는 1분기보고서에서 확인 필요.

---

### 해약환급금준비금

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 99 | 해약환급금준비금 | XBRL | `dart_SurrenderValueReserve` | ⚠️ 신한라이프·KB라이프 MISS |

---

### 손해율

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 101 | 예상손해율(A) | 원문XML | 보험금 예실차비율 테이블 `예상손해율` 항목 | ⚠️ 미래에셋·현대해상 MISS (테이블 없음), 한화 소오차 |
| 102 | 실제손해율(B) | 원문XML | 동일 테이블 `실제손해율` 항목 | ⚠️ 미래에셋·현대해상 MISS, 한화 소오차 |
| 103 | 예실차비율(b−a) | 계산 | Col101 − Col102 | ⚠️ Col101·102가 MISS/오차인 경우 동일하게 영향 |
| 104 | 실제/예상보험금 | 계산 | Col102 / Col101 | ⚠️ Col101·102가 MISS/오차인 경우 동일하게 영향 |
| 107~109 | 손해율·위험보험료·사고보험금 | — | **추출 불가** | ❌ 전 회사 MISS |

> **107~109 MISS 이유**: 위험보험료·사고보험금 절대값은 XBRL 태그가 없고, 원문 XML에서도  
> 보험종류별로 분리되어 있어 합산 추출이 복잡함.

---

### 예실차 세부

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 111 | 예실차 합계 | — | ❌ 전 회사 MISS |
| 112 | 보험금예실차 | — | ❌ 전 회사 MISS |
| 113~114 | 예상/실제보험금 | — | ❌ 전 회사 MISS |
| 115 | 사업비예실차 | — | ❌ 전 회사 MISS |
| 116~117 | 예상/실제사업비 | — | ❌ 전 회사 MISS |

> **MISS 이유**: DART XBRL에 예실차 관련 단독 태그 없음.  
> 보험종류별 다차원 집계가 필요해 자동화 어려움. **수동 입력 필요.**

---

### 계리적 가정변경 민감도

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 146~150 | BEL+RA기준 (해지율·위험률·사업비·기타·물량) | XBRL → 원문XML | XBRL 보험계약 주석의 가정변경 효과 중 BEL+RA 기준 항목. fallback: 사업보고서 원문XML의 가정변경효과 테이블 | ⚠️ 동양·삼성화재 MISS, 한화 부호오류 |
| 151 | 기타 감응도 | — | ❌ MISS (일부 회사만 공시, 자동화 불가) | |
| 153~158 | CSM기준 (해지율~손실요소) | XBRL → 원문XML | XBRL 보험계약 주석의 가정변경 효과 중 CSM 기준 항목. fallback: 사업보고서 원문XML의 가정변경효과 테이블 | ⚠️ 동양·삼성화재 MISS |
| 159 | CHECK | — | ⚠️ 교보·삼성화재·현대해상 오차 |

---

## 전체 정확도 요약

정답 데이터(`202512_동업사 공시비교_DATA_작업_260320.xlsx`)와 비교한 결과입니다.  
오차 허용 기준: **0.1% 이내** = OK, 초과 = FAIL, 미추출 = MISS

| 기간 | 보고서 종류 | 비교 가능 건수 | OK | FAIL | MISS | **OK율** |
|------|-----------|-------------:|---:|-----:|-----:|--------:|
| 2312 | 2023년 사업보고서 | 1,095 | 455 | 199 | 441 | **42%** |
| 2412 | 2024년 사업보고서 | 1,219 | 716 | 75 | 428 | **59%** |
| 2503 | 2025년 1분기 | 1,033 | 558 | 43 | 432 | **54%** |
| 2506 | 2025년 반기 | 1,059 | 512 | 97 | 450 | **48%** |
| 2509 | 2025년 3분기 | 1,058 | 577 | 52 | 429 | **55%** |
| 2512 | 2025년 사업보고서 | 1,255 | 955 | 57 | 243 | **76%** |
| **합계** | | **6,719** | **3,773** | **523** | **2,423** | **56%** |

**기간별 특이사항**

- **2312 (42%)**: 전기말(2022년말) 참조 항목에서 오차 다수 — 2022년 XBRL이 없어 이전기 비교수치가 불일치.
- **2412~2509 (48~59%)**: BS·PL 기본 항목은 잘 추출되나, 사업보고서 전용 항목(CSM변동·K-ICS·운용자산·감응도 등) 미추출로 MISS 많음.
- **2512 (76%)**: 사업보고서 전용 항목까지 포함해 가장 높은 정확도. 주요 MISS는 예실차 세부(XBRL 미제공)·일부 CSM기간별·OCI 세부잔액.

> **Excel 주황색 셀**: 정답과 0.1% 이상 차이나는 셀 (FAIL). 빈 셀은 미추출(MISS).

---

## 주요 한계 및 수동입력 필요 항목

| 항목 | 이유 |
|------|------|
| Col107~117 (예실차·위험보험료) | XBRL 미제공, 원문 XML 집계 불가 |
| Col83~87 (CSM 기간별) 일부 회사 | 사업보고서 원문 XML에 테이블 없음 |
| Col92~97 (K-ICS) 일부 회사 | 사업보고서 제출 시 산출중 상태 |
| Col22~23 (운용자산) 일부 회사 | 원문 XML에 테이블 없음 |
| Col25~30 (OCI 세부잔액) 일부 회사 | XBRL에서 회사마다 항목 분류 방식이 달라 자동 매핑 불가 |
| 2212 (2022년말) | IFRS17 전환 이전 데이터로 DART에 XBRL 미제공 — 자동화 불가, 수동 입력 필요 |

---

## 파일 구조

```
parse_annual.py       실행 스크립트
dart_api.py           DART OpenAPI 연동
DATA_작업_빈칸.xlsx    채울 대상 Excel
data/
  taxonomy.csv        XBRL 태그 한글명 매핑 (자동 생성)
  xbrl_taxonomy.xlsx  taxonomy Excel 버전
  {회사명}/           XBRL·원문XML 캐시 (자동 생성)
```

