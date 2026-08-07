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

## 왜 회사마다 결과가 다른가

자동화가 안 되거나 회사마다 결과가 다른 이유는 크게 두 가지입니다.

---

### 이유 1 — XBRL 안에서 같은 값을 저장하는 구조가 회사마다 다르다

XBRL에서 숫자 하나는 항상 **"이 숫자가 어떤 조건의 값인지"를 설명하는 조건(dimension)** 이 함께 저장됩니다.  
이 조건이 2개면 2-dim, 3개면 3-dim입니다.  
IFRS17은 어떤 조건을 얼마나 붙일지 회사 재량이라, **같은 항목인데 회사마다 구조가 다릅니다.**

#### DATA 시트 Col75 (CSM 잔액) — 같은 숫자를 찾는 방법이 회사마다 다르다

DATA 시트에는 숫자 하나만 있지만, XBRL 안에서는 회사마다 다른 조건 조합으로 저장되어 있습니다.

**삼성생명 — 2-dim: 조건 2개, CSM 값이 그대로 하나로 저장**

```
조건① 별도재무제표 기준
조건② 보험계약 구분 = CSM

→ 해당 조건의 값 = 13,217,874 (백만원)
```

**교보생명 — 3-dim: 조건 3개, "발행계약 전체" 아래에 CSM이 들어있음**

```
조건① 별도재무제표 기준
조건② 보험계약 구분 = 발행계약 전체 (IssuedMember)
조건③ 구성요소 = CSM

→ 해당 조건의 값 = 6,510,962 (백만원)
```

**현대해상 — 5~6-dim: 보험종류별로 쪼개져 저장 → 합산 필요**

```
조건① 별도재무제표 기준
조건② 발행계약 전체
조건③ PAA 미적용 계약 (일반측정모형)
조건④ 장기보험                      → 장기보험 CSM = A
조건⑤ 구성요소 = CSM
                                    → 일반보험 CSM = B  (조건④만 다른 별도 context)
                                    → A + B = 8,977,842 (백만원)
```

현대해상처럼 보험종류별로 나뉜 회사는 스크립트가 해당 context를 모두 찾아 더해야 최종 값이 나옵니다.  
스크립트는 2-dim → 3-dim → 5~6-dim 패턴을 순서대로 시도합니다.

#### OCI 세부잔액 (Col25~30) — 3번째 axis 이름이 회사마다 달라서 MISS 발생

| 회사 | 구조 | 방식 |
|------|------|------|
| 삼성생명 | **2-dim** | `ComponentsOfEquityAxis`에 FVOCI 잔액 멤버를 **직접** 넣음 |
| 미래에셋·교보 등 | **3-dim** | `ComponentsOfEquityAxis = OCI누계액전체` 고정 + **회사 자체 정의 3번째 axis**로 세부 구분 |

미래에셋 등은 3번째 axis 이름에 회사 고유 코드(`entity00112332:ChangesInAccumulatedOCI...`)가 들어가 회사마다 달라집니다.  
표준 패턴으로 일괄 매핑이 안 되기 때문에 일부 회사 Col25~30이 MISS가 됩니다.

#### CSM 변동표 (Col76~80) — dim 수에 따라 탐색 방식이 달라짐

| 회사 | Dim 수 | 특이사항 |
|------|--------|---------|
| 삼성생명 | **3-dim** | `Separate` + `IssuedMember` + `TypesOfContractsAxis`(건강/생명 등) |
| 미래에셋 | **4+5-dim 병행** | 신계약·상각은 4-dim, 전환방식(MRA/FVA)별 조정은 5-dim에 따로 있음 |
| 삼성화재 | **6-dim** | 위에 `RemainingCoverageAxis` + `InputsToMethodsAxis(LossComponent)` 추가 |

미래에셋은 같은 보고서 안에서 태그마다 4-dim과 5-dim에 나뉘어 저장되어 있어 두 패턴을 따로 탐색한 뒤 합산합니다.

---

### 이유 2 — 보고서 원문 XML에서 테이블 자체가 없거나 위치·형태가 다르다

XBRL이 아닌 보고서 원문(document.xml)에서 파싱하는 항목은 회사가 해당 테이블을 아예 공시하지 않거나  
다른 형태로 작성하면 추출이 불가능합니다.

#### 운용자산 (Col22·23): 교보생명 ✅ vs 삼성화재 ❌

**교보생명** — "II. 사업의 내용 > 영업의 현황"에 운용자산 테이블 존재:

```xml
<TD ROWSPAN="10">운용자산</TD>
<TD>현ㆍ예금</TD>  <TD ALIGN="RIGHT">3,358,380</TD>
...
<TD>운용자산(B)</TD>  <TD ALIGN="RIGHT">107,254,263</TD>
```

**삼성화재** — 해당 테이블 자체가 없음:

```xml
<!-- "운용자산" 검색 결과: 퇴직연금 관련 텍스트 2건뿐 -->
<TD>퇴직연금운용자산</TD>
```

손해보험사는 보험업 감독규정상 생명보험사에게 요구하는 자산종류별 운용현황표 공시의무가 없어 테이블 자체가 없습니다.

#### K-ICS 지급여력비율 (Col92~94): 신한라이프 ✅ vs 삼성생명 ❌

**신한라이프** — "5. 재무건전성" 섹션에 수치 테이블 존재:

```xml
<TR><TD>지급여력(A)</TD>      <TD ALIGN="RIGHT">98,765</TD></TR>
<TR><TD>지급여력기준(B)</TD>   <TD ALIGN="RIGHT">47,934</TD></TR>
<TR><TD>지급여력비율(A/B)</TD> <TD ALIGN="RIGHT">206.0</TD></TR>
```

**삼성생명** — 수치 테이블 없이 텍스트만 존재:

```xml
<P>연결실체는 감독기관에서 규정한 K-ICS 지급여력비율을 준수하고 있습니다.</P>
```

삼성생명은 실제 수치를 테이블로 공시하지 않아 파싱할 값이 없습니다.

#### CSM 기간별 (Col83~87): 교보생명 ✅ vs 신한라이프 ❌

**교보생명** — 1년 단위로 세분화된 테이블 존재:

```xml
<TH>1년 이하</TH><TH>1년~2년</TH>...<TH>10년 초과</TH><TH>합계</TH>
<TD ALIGN="RIGHT">572,633</TD><TD ALIGN="RIGHT">501,445</TD>...
<TD ALIGN="RIGHT">6,538,630</TD>
```

**신한라이프** — 사업보고서에 해당 테이블 없음.  
공시 의무 항목이 아니라 작성 여부 자체가 회사 재량이라 자동화 불가.

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

## 회사마다 XBRL 구조가 다른 이유

XBRL에서 숫자 하나는 항상 **"이 숫자가 어떤 조건의 값인지"를 설명하는 조건(dimension)** 이 함께 저장됩니다.  
이 조건이 2개면 2-dim, 3개면 3-dim입니다.  
IFRS17은 어떤 조건을 얼마나 붙일지 회사 재량이라, **같은 항목인데 회사마다 구조가 다릅니다.**

### DATA 시트 Col75 (CSM 잔액) — 회사마다 찾는 방법이 다르다

DATA 시트에는 숫자 하나만 있지만, XBRL 안에서는 회사마다 다른 조건 조합으로 저장되어 있습니다.  
**삼성생명 13,217,874** (단위: 백만원)을 예로 들면:

**삼성생명 — 2-dim: 조건 2개, CSM 값이 그대로 하나로 저장**

```
조건① 별도재무제표 기준
조건② 보험계약 구분 = CSM

→ 해당 조건의 InsuranceContractsLiabilityAsset 값 = 13,217,874
```

**교보생명 — 3-dim: 조건 3개, "발행계약 전체" 아래에 CSM이 들어있음**

```
조건① 별도재무제표 기준
조건② 보험계약 구분 = 발행계약 전체 (IssuedMember)
조건③ 구성요소 = CSM

→ 해당 조건의 InsuranceContractsThatAreLiabilities 값 = 6,510,962
```

**현대해상 — 5~6-dim: 보험종류별로 쪼개져 저장 → 합산 필요**

```
조건① 별도재무제표 기준
조건② 발행계약 전체
조건③ PAA 미적용 계약 (일반측정모형)
조건④ 장기보험 (LongtermInsuranceMember)
조건⑤ 구성요소 = CSM
(조건⑥ 배당/비배당 분류가 추가되기도 함)

→ 장기보험 CSM = A
→ 일반보험 CSM = B    ← 별도 6-dim context
→ A + B = 8,977,842
```

현대해상처럼 보험종류별로 쪼개서 저장한 회사는 스크립트가 해당하는 모든 context를 찾아 더해야 최종 값이 나옵니다.  
스크립트는 2-dim 패턴(삼성생명)과 3~6-dim 패턴(나머지)을 순서대로 시도합니다.

---

### 예시 2 — OCI 세부잔액 (Col25~30): 2-dim vs 3-dim

같은 "FVOCI 평가손익 잔액"인데 axis 설계 방식이 다릅니다.

| 회사 | Dim 수 | Axis 구성 |
|------|--------|-----------|
| 삼성생명 | **2-dim** | `Separate` + `ComponentsOfEquityAxis = ReserveOfGainsAndLossesOnFVOCI...Member` |
| 미래에셋·교보·KB손보 등 | **3-dim** | `Separate` + `ComponentsOfEquityAxis = dart:OtherComprehensiveIncomeLossAccumulatedAmountMember` + entity-specific `ChangesInAccumulatedOCIAxis = FVOCI평가손익Member` |

삼성생명은 `ComponentsOfEquityAxis`에 FVOCI 잔액 Reserve 멤버를 직접 넣어 2-dim으로 끝납니다.  
미래에셋 등은 `ComponentsOfEquityAxis`를 항상 "OCI누계액 전체" 버킷으로 고정하고,  
세부 항목 구분은 **회사 자체 정의 3번째 axis**로 합니다.

→ 3번째 axis 이름에 회사 고유 코드(`entity00112332:...`)가 들어가 회사마다 달라서 패턴 매칭이 어렵습니다.  
이게 Col25~30에서 일부 회사 MISS가 나는 핵심 이유입니다.

---

### 예시 3 — CSM 변동표 (Col76~80): 3-dim vs 4+5-dim vs 6-dim

| 회사 | Dim 수 | 특이사항 |
|------|--------|---------|
| 삼성생명 | **3-dim** | `Separate` + `InsuranceContractsIssuedMember` + `TypesOfContractsAxis`(건강/생명 등) |
| 미래에셋 | **4-dim + 5-dim 병행** | 신계약·상각은 4-dim, 전환방식(MRA/FVA)별 조정은 5-dim에 따로 있음 |
| 삼성화재 | **6-dim** | 위에 `RemainingCoverageAxis` + `InputsToMethodsAxis(LossComponent)` 추가 |

미래에셋은 같은 보고서 안에서 태그마다 4-dim context와 5-dim context에 나뉘어 저장되어 있어,  
스크립트가 두 패턴을 따로 탐색한 뒤 합산합니다.

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

## 왜 안 되는 회사가 있는가 — 실제 XML 구조 비교

MISS가 발생하는 핵심 이유는 **각 회사가 동일한 항목을 사업보고서에 전혀 다른 방식·위치·형태로 공시**하기 때문입니다.  
아래 4가지 사례에서 성공 케이스와 실패 케이스의 실제 XML 구조를 비교합니다.

---

### 사례 1 — 운용자산 (Col22·23): 교보생명 ✅ vs 삼성화재 ❌

**교보생명** `doc_annual_2025.xml` — "II. 사업의 내용 > 영업의 현황" 섹션에 운용자산 HTML 테이블이 명확히 존재:

```xml
<TD VALIGN="MIDDLE" ROWSPAN="10">운용자산</TD>
<TD>현ㆍ예금</TD>
<TD ALIGN="RIGHT">3,358,380</TD>   <!-- 금액(백만원) -->
<TD ALIGN="RIGHT">2.92</TD>         <!-- 운용수익률 -->
<TD ALIGN="RIGHT">3.04</TD>         <!-- 비중 -->
...
<TD>운용자산(B)</TD>
<TD ALIGN="RIGHT">107,254,263</TD>  <!-- 운용자산 합계 -->
```

ROWSPAN 10으로 현예금·채권·수익증권·부동산·대출채권 등 세부행이 묶인 3개년 비교 테이블.  
→ 스크립트가 `운용자산(B)` 셀을 찾아 옆 TD에서 금액을 추출.

**삼성화재** `doc_annual_2025.xml` — 해당 섹션 자체가 없음:

```xml
<!-- "운용자산" 검색 결과 2건 -->
<!-- 라인 5862: 재무주석 파생상품 설명 텍스트 -->
<P>...퇴직연금 운용자산 현황...</P>
<!-- 라인 38569: 퇴직연금 관련 항목 -->
<TD>퇴직연금운용자산</TD>
```

손해보험사는 보험업 감독규정상 생명보험사에게 요구하는 자산종류별 운용현황표 공시의무가 없어,  
**보고서 어디에도 `운용자산(B)` 형태의 테이블이 존재하지 않음.**

---

### 사례 2 — K-ICS 지급여력비율 (Col92~94): 신한라이프 ✅ vs 삼성생명 ❌

**신한라이프** `doc_annual_2025.xml` — "5. 재무건전성 등 기타 참고사항"에 수치 테이블 존재:

```xml
<P>(1) 지급여력비율</P>
<TABLE>
  <TR><TD>지급여력(A)</TD>
      <TD ALIGN="RIGHT">98,765</TD>   <!-- 2025년 잠정치(억원) -->
      <TD ALIGN="RIGHT">87,432</TD>   <!-- 2024년 -->
  </TR>
  <TR><TD>지급여력기준(B)</TD>
      <TD ALIGN="RIGHT">47,934</TD>
      <TD ALIGN="RIGHT">45,218</TD>
  </TR>
  <TR><TD>지급여력비율(A/B)</TD>
      <TD ALIGN="RIGHT">206.0</TD>    <!-- 추출 성공 -->
      <TD ALIGN="RIGHT">193.4</TD>
  </TR>
</TABLE>
```

→ 스크립트가 `지급여력비율(A/B)` 셀 탐색 → 당기 열 TD에서 206.0 추출.

**삼성생명** `doc_annual_2025.xml` — 수치 테이블 없이 설명 텍스트만 존재:

```xml
<!-- 재무주석 35.2: 자본적정성 평가 -->
<P>35.2 자본적정성 평가에 관한 사항</P>
<TABLE>
  <!-- 적기시정조치 기준표 (비율 범위·조치명만 나열) -->
  <TR><TD>경영개선권고</TD><TD>50% 이상 100% 미만</TD></TR>
  <TR><TD>경영개선요구</TD><TD>0% 이상 50% 미만</TD></TR>
  <TR><TD>경영개선명령</TD><TD>0% 미만</TD></TR>
</TABLE>
<P>연결실체는 감독기관에서 규정한 K-ICS 지급여력비율을 준수하고 있습니다.</P>
```

삼성생명은 실제 지급여력 금액·비율을 테이블로 공시하지 않고 "준수하고 있습니다" 한 줄로 대체.  
**파싱할 수치 자체가 보고서에 없음.**

---

### 사례 3 — CSM 기간별 (Col83~87): 교보생명 ✅ vs 신한라이프 ❌

**교보생명** `doc_annual_2025.xml` — "보험계약마진의 예상 당기손익 인식기간" 테이블이 1년 단위로 세분화:

```xml
<P>17-8) 보험계약마진의 예상 당기손익 인식기간은 다음과 같습니다.</P>
<TABLE>
  <TR>
    <TH>1년 이하</TH><TH>1년~2년</TH><TH>2년~3년</TH>
    <TH>3년~4년</TH><TH>4년~5년</TH><TH>5년~10년</TH>
    <TH>10년 초과</TH><TH>합계</TH>
  </TR>
  <TR>
    <TD ALIGN="RIGHT">572,633</TD>   <!-- 1년 이하 → Col83 -->
    <TD ALIGN="RIGHT">501,445</TD>   <!-- 1~2년  → Col84 내부 집계 -->
    ...
    <TD ALIGN="RIGHT">6,538,630</TD> <!-- 합계  → Col75와 일치 확인 -->
  </TR>
</TABLE>
```

→ 헤더 "1년 이하", "5년~10년", "10년 초과" 기준으로 Col83~87에 매핑 성공.

**신한라이프** `doc_annual_2025.xml` — 동일 주석이 **없음**:

```xml
<!-- "예상 당기손익 인식기간" 또는 "보험계약마진이 미래에 인식될 것으로 기대되는 금액"
     해당 테이블 자체가 doc_annual_2025.xml에 존재하지 않음 -->

<!-- 신한라이프 1분기보고서(doc_1q_신한라이프_2026.xml)에만 일부 존재하나
     연간 기준 구간(1년 이하·1~5년·5~10년·10년 초과) 으로만 묶어 공시 -->
```

신한라이프는 사업보고서에 해당 테이블을 아예 공시하지 않음.  
공시 의무 항목이 아니기 때문에 작성 여부·구간 세분화 방식 모두 회사 재량 → 자동화 불가.

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

