# 동업사 공시비교 자동 추출 (parse_annual.py)

12개 보험사의 DART 공시(XBRL + document.xml)에서 DATA 시트 컬럼을 자동 추출합니다.

**현재 성능 (2512 기준):** OK=989, FAIL=48, MISS=218 / 전체 1,255건

## 실행

```bash
python parse_annual.py        # 추출 → data/data_sheet_2512.csv
python verify_2512.py         # Excel 정답과 비교 검증
```

---

## 컬럼별 추출 방법

### Col4 총자산
XBRL 별도 당기말 `ifrs-full_Assets`

### Col5 FVPL
XBRL 별도 당기말 `ifrs-full_FinancialAssetsAtFairValueThroughProfitOrLoss`

### Col6 FVOCI
XBRL 별도 당기말 `ifrs-full_FinancialAssetsAtFairValueThroughOtherComprehensiveIncome`  
→ fallback: `dart_SecuritiesAtFairValueThroughOtherComprehensiveIncome`

### Col7 AC
XBRL 별도 당기말 `ifrs-full_FinancialAssetsAtAmortisedCost`

### Col8 종속·관계기업
XBRL 별도 당기말 `ifrs-full_InvestmentsInSubsidiariesJointVenturesAndAssociates`  
→ fallback: `ifrs-full_InvestmentAccountedForUsingEquityMethod`

### Col9~11 총자산대비비중(FVPL/FVOCI/AC)
계산: Col5/Col4, Col6/Col4, Col7/Col4

### Col12 부채
XBRL 별도 당기말 `ifrs-full_Liabilities`

### Col13 보험계약부채
XBRL 별도 당기말 `ifrs-full_InsuranceContractsIssuedThatAreLiabilities`  
→ fallback: dart_보험계약부채(참여특성有) + dart_보험계약부채(참여특성無) 합산

### Col14 자본(신종자본증권 제외)
계산: Col15 - Col16

### Col15 자본
XBRL 별도 당기말 `ifrs-full_Equity`

### Col16 신종자본증권
XBRL 별도 당기말 `dart_HybridBonds`

### Col17 기타포괄손익누계액
XBRL 별도 당기말 `ifrs-full_AccumulatedOtherComprehensiveIncome`

### Col18 이익잉여금
XBRL 별도 당기말 `ifrs-full_RetainedEarnings`

### Col19 기시자본(연결)
XBRL 전기말 `ifrs-full_Equity` (별도 당기말과 동일 태그, 전기 instant context)

### Col20 기시자본(신종자본증권 제외)
계산: 전기말 Col15 - 전기말 Col16

### Col21 자산운용률·총자산
사업보고서 document.xml → 운용자산/자산운용률 테이블에서 총 자 산(A) 항목

### Col22 운용자산
사업보고서 document.xml → 운용자산(B) 항목  
단위가 억원이면 ×100 자동 변환 (총자산과 비교해 5% 미만이면 억원 단위로 판단)

### Col23 자산운용률
사업보고서 document.xml → 자산운용률(B/A) 항목 (% → 소수)

### Col24 기타포괄손익누계액 합계
= Col17 (동일 값)

### Col25 FVOCI평가손익
XBRL 별도 당기말 OCI 누계 3-dim context에서 `FinancialAssets` 키워드 멤버 합산  
→ 없으면 사업보고서 document.xml OCI 세부잔액 테이블

### Col26 보험계약자산(부채)순금융손익
XBRL 별도 당기말 OCI 누계에서 `InsuranceContract` 키워드 멤버 합산  
→ 없으면 doc.xml: `보험계약 관련 금융손익` 행

### Col27 재보험계약자산(부채)
XBRL 별도 당기말 OCI 누계에서 `ReinsuranceFinance` 키워드 멤버  
→ 없으면 doc.xml: `재보험계약 관련 금융손익` 행

### Col28 현금흐름위험회피
XBRL 별도 당기말 OCI 누계에서 `CashFlowHedges` 키워드 멤버  
→ 없으면 doc.xml: `현금흐름위험회피` 행

### Col29 재평가잉여금
XBRL **전기말** OCI 누계에서 `Revaluation` 키워드 멤버  
→ 없으면 doc.xml: `재평가잉여금` 행

### Col30 확정급여부채
XBRL 별도 당기말 OCI 누계에서 `RemeasurementsOfDefinedBenefit` 키워드 멤버  
→ 없으면 doc.xml: `확정급여채무 재측정요소` 행

### Col31 CHECK(OCI)
계산: Col24 - (Col25+Col26+Col27+Col28+Col29+Col30)  
모든 구성항목(Col25~30) 있을 때만 계산

### Col32 보험손익
XBRL 별도 당기 `ifrs-full_InsuranceServiceResult`

### Col33 보험서비스수익
XBRL 별도 당기 `dart_OperatingIncomeInsurance`

### Col34 보험서비스비용
XBRL 별도 당기 `dart_OperatingExpenseInsurance`

### Col35 투자손익
XBRL 별도 당기 `dart_InvestmentIncomeExpenses`

### Col36 보험금융손익
계산: `dart_InsuranceFinanceIncome...ProfitOrLoss` - `dart_InsuranceFinanceExpenses...ProfitOrLoss`

### Col37 재보험금융손익
계산: `dart_FinanceIncomeFromReinsurance...ProfitOrLoss` - `dart_FinanceExpensesFromReinsurance...ProfitOrLoss`

### Col38 금융손익
계산: Col35 - Col36 - Col37 + Col39 - Col40

### Col39 재산관리비
XBRL 별도 당기 `ifrs-full_SellingGeneralAndAdministrativeExpense`

### Col40 기타투자손익
XBRL 별도 당기 수수료수익 + 임대료 + 기타영업수익 - 기타영업비용 (여러 태그 합산)

### Col41 영업이익
XBRL 별도 당기 `ifrs-full_ProfitLossFromOperatingActivities`

### Col42 영업외손익
XBRL 별도 당기 `ifrs-full_NonOperatingProfitLoss`  
→ 없으면 계산: Col43 - Col41

### Col43 세전손익
XBRL 별도 당기 `ifrs-full_ProfitLossBeforeTax`

### Col44 당기순이익
XBRL 별도 당기 `ifrs-full_ProfitLoss`

### Col45 CHECK1 (영업손익)
계산: Col41 + Col42 - Col43

### Col46 CHECK2 (투자손익)
계산: Col36 + Col37 + Col38 - Col39 + Col40 - Col35  
Col35/36/38/39/40 모두 있을 때만 계산

### Col47 CHECK3 (보험투자손익)
계산: Col32 + Col35 - Col41

### Col48 보험손익(별도표)
= Col32 (동일)

### Col49 투자손익(별도표)
= Col35 (동일)

### Col50 영업외손익(별도표)
= Col42 (동일)

### Col51 세전손익 합계
= Col43 (동일)

### Col55 기타포괄손익누계액 기초
XBRL 전기말(전기말 instant) `ifrs-full_AccumulatedOtherComprehensiveIncome`

### Col56 FVOCI평가손익 변동
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxFinancialAssets...FVOCI...`

### Col57 대손충당금
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxCreditLosses...`

### Col58 보험계약금융손익(재보 포함) OCI
계산: 보험OCI + 재보험OCI  
= `ifrs-full_OtherComprehensiveIncomeNetOfTaxInsuranceFinanceIncome...` + `...FromReinsurance...`

### Col59 현금흐름위험회피파생상품
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxCashFlowHedges`  
→ fallback: `dart_OtherComprehensiveIncomeNetOfTaxGainsLossesOnHedgingInstrument`와 합산

### Col60 재평가잉여금 변동
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxGainsLossesOnRevaluation...`

### Col61 확정급여부채 재측정
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxGainsLossesOnRemeasurementsOfDefinedBenefitPlans`

### Col62 해외사업환산손익
XBRL 별도 당기 `ifrs-full_OtherComprehensiveIncomeNetOfTaxExchangeDifferencesOnTranslation`

### Col63 기타포괄손익누계액 기말
= Col17 (당기말 = BS 잔액)

### Col64 CHECK(OCI 자본변동)
계산: Col63 - Col55 - (Col56+57+58+59+60+61+62)  
Col56~62 전부 있을 때만 계산

### Col65 이익잉여금 기초
XBRL 전기말 `ifrs-full_RetainedEarnings`

### Col66 결산배당
XBRL 별도 당기 `dart_DividendsPaidClassifiedAsFinancingActivities` (음수 저장)

### Col67 자기주식 소각/처분
XBRL 별도 당기 `ifrs-full_CancellationOfTreasuryShares`  
context: 2-dim (Separate + RetainedEarnings component)

### Col68 순이익
= Col44

### Col69 기타(가업결합·OCI재분류 등)
XBRL 별도 당기 `dart_TransferOfAmountRecognisedInOtherComprehensiveIncome...`  
또는 `ifrs-full_OtherComprehensiveIncomeNetOfTax...FVOCI`  
context: 2-dim (Separate + RetainedEarnings component)

### Col70 신종자본증권 배당
XBRL 별도 당기 `dart_DividendToHybridBond` (음수 저장)  
context: 2-dim (Separate + RetainedEarnings component)

### Col71 이익잉여금 기말
= Col18 (당기말 BS 잔액)

### Col72 22년말(ifrs4)
IFRS17 전환 시점 값. 2512 기준 모두 0(수식 결과). XBRL/doc.xml 미포함.

### Col73 BEL (순부채 기준)
- **XBRL 있는 회사**: 별도 당기말 보험계약부채 현재가치추정치 축 집계
- **5개사 fallback** (교보/신한라이프/KB생명/메리츠화재/KB손해보험): 1분기보고서 document.xml → 보험계약부채 구성요소 테이블에서 `현재가치 추정치` 컬럼 합산, target=Col13 기준 당기 테이블 선택

### Col74 RA
- XBRL 또는 1분기 doc.xml 위험조정 컬럼

### Col75 CSM
- XBRL 또는 1분기 doc.xml 보험계약마진 컬럼

### Col76 신계약CSM
- XBRL: `dart/ifrs-full_ChangesInFutureServicesDueToNewContracts...ContractualServiceMarginMember` 등
- 5개사 fallback: 사업보고서 doc.xml CSM 변동표 → `신계약`, `최초 인식` 라벨 행

### Col77 CSM상각
- XBRL: 서비스이전 보험계약마진 축
- 5개사 fallback: doc.xml → `보험계약마진 상각`, `서비스 제공`, `당기손익으로 인식한 보험계약마진` 라벨 행 (항상 양수 저장)

### Col78 CSM조정
- XBRL: 보험계약마진을 조정하는 추정치 변동 축
- 5개사 fallback: doc.xml → `보험계약마진을 조정` 라벨 행

### Col79 보험금융손익(CSM)
- XBRL: 보험금융손익 CSM 축
- 5개사 fallback: doc.xml → `순보험금융손익` 또는 `보험금융손익` 행  
  ※ 신한라이프: `보험금융손익` 하위에 sub-label(당기손익)이 있어 1열 offset 처리

### Col80 전기CSM
- XBRL: 전기말 context CSM 값
- 5개사 fallback: doc.xml CSM 변동표 → `기초 순장부금액`, `기초 보험계약마진` 등 라벨 행

### Col81 보험료배분접근법
XBRL에서 PAA 관련 잔여보장부채

### Col82 CSM 증감 CHECK
계산: Col80 + Col76 - Col77 + Col78 + Col79 - Col75

### Col83~87 CSM 기간별 기대수익인식금액
사업보고서 document.xml → CSM 만기 테이블  
- target(Col88=Col75)으로 당기 테이블 선택
- 버킷 구조 자동 판별: 5컬럼(삼성화재), 7컬럼(교보), 9컬럼(DB손보), 10+컬럼(동양), 15+컬럼(한화)
- 매핑: Col83=1년이하, Col84=1~3년, Col85=3~5년, Col86=5~10년, Col87=10년초과
- **미추출**: 신한라이프/삼성생명/미래에셋/KB라이프/현대해상/메리츠화재/KB손보 (doc.xml 미포함)

### Col88 CSM 합계
= Col75

### Col89 CHECK(CSM기간별)
계산: Col88 - (Col83+84+85+86+87)  
Col83~87 모두 있을 때만 계산, 반올림 오차 ≤2 이면 0

### Col92~94 K-ICS 잠정치 (비율/가용자본/요구자본)
사업보고서 document.xml → 지급여력비율 테이블  
- 비율: `지급여력비율(A/B)` 행 첫 번째 수치 (% → 소수)  
- 억원 단위 자동 감지: 가용자본 < 100,000 이면 ×100  
- **미추출**: 삼성생명/삼성화재/현대해상 (doc.xml에 테이블 없음)

### Col95~97 K-ICS 최종치 (비율/가용자본/요구자본)
`data/kics_2512.csv` 수동 파일 (연도별로 별도 관리 필요)

### Col99 해약환급금준비금
XBRL 별도 당기말/당기 `dart_SurrenderValueReserve` 또는 관련 잔액 태그

### Col101 예상손해율(A)
사업보고서 document.xml  
- Method1: TE ACODE 태그 `ExpectedLossRateOfInsuranceClaim...`, `ExpectedLossRatioOf...`
- Method2: 테이블 파싱 → `예상손해율` 라벨 다음 수치, 또는 헤더 이후 첫 번째 수치  
  ※ 신한라이프: `예상손해율(A)|93.12%|...` 라벨 바로 다음 값  
  ※ 교보/삼성화재: `예상손해율(%)|실제손해율(%)| | |94.5%|96.9%` 헤더 이후 값

### Col102 실제손해율(B)
doc.xml → Method1/2 동일, `실제손해율` 라벨 다음 수치

### Col103 보험금예실차비율(b-a)
계산: Col101 - Col102

### Col104 실제/예상보험금 비율
계산: Col102 / Col101

### Col107 손해율(손보 별도)
doc.xml 위험보험료 대비 손해율 테이블 (미추출 회사 있음)

### Col108 위험보험료
doc.xml 테이블 (미추출 회사 있음)

### Col109 사고보험금(간접제외)
doc.xml 테이블 (미추출 회사 있음)

### Col111 예실차 합계
doc.xml (미추출 — 보험종류별 다차원 집계 필요)

### Col112 보험금예실차 / Col113 예상보험금 / Col114 실제보험금
doc.xml (미추출)

### Col115 사업비예실차 / Col116 예상사업비 / Col117 실제사업비
doc.xml (미추출)

### Col146~150 BEL+RA 기준 가정변경 (해지율/위험률/사업비/기타/물량)
XBRL 별도 당기 `dart_ChangeIn...Assumption...` BEL component 3-dim context  
→ 없으면 사업보고서 doc.xml 가정변경효과 테이블 → BEL 컬럼(첫 번째) + RA 컬럼 합산  
  ※ 교보: 단위 억원(×100), ※ 동양/삼성화재/삼성생명: doc.xml 파싱 미지원(스킵)

### Col153~158 CSM 기준 가정변경 (해지율/위험률/사업비/기타/물량/손실요소)
XBRL 별도 당기 `dart_ChangeIn...Assumption...` CSM component  
→ 없으면 doc.xml 가정변경 테이블 → CSM 컬럼(마지막)  
  Col158(손실요소): XBRL `dart_FluctuationsDueToLossFactors...` 또는 DoNotAdjust/Adjust 패턴

---

## 핵심 기술 사항

### XBRL context 구조
- **1-dim (Separate만)**: BS/PL 대부분 항목
- **2-dim (Separate + ComponentsOfEquity)**: 자본변동 항목 (Col66~70) — RetainedEarnings / CapitalAdjustments component
- **3-dim (Separate + OCI member + entity-specific)**: OCI 누계 잔액 (Col25~30)
- **Multi-dim (Separate + 보험계약 축 + CSM 축 등)**: BEL/RA/CSM, 감응도

### document.xml 파싱 방법
HTML 테이블이 XML에 embedded. `<TABLE>` 태그 추출 → 태그 제거 후 `|` 구분 셀 배열로 변환.

**섹션 분리 (연결/별도)**: 파일 내 테이블 위치 gap > 500KB → 별도 섹션 (마지막 섹션)  
**당기 테이블 선택**: target 값(CSM기말, OCI합계 등)과 가장 가까운 테이블 선택

### 단위 변환
- XBRL: 원(KRW) → 백만원 (`/ 1e6`)
- doc.xml: 회사마다 억원 또는 백만원 → 총자산 대비 비율로 자동 감지 후 변환
