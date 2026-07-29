# 동업사 공시비교 자동 추출 (parse_annual.py)

12개 보험사의 DART 공시(XBRL + document.xml)에서 DATA 시트 항목을 자동 추출합니다.

**현재 성능 (2512 기준):** OK=989, FAIL=47, MISS=219 / 전체 1,255건

---

## 실행

```bash
python parse_annual.py        # 데이터 추출 → data/data_sheet_2512.csv
python verify_2512.py         # Excel 정답과 비교 검증
```

---

## 데이터 소스별 추출 항목

### 1. XBRL (사업보고서 12월말) — `dart_api.py` + `parse_annual.py`

DART API `fnlttXbrl.xml`로 다운로드한 `.xbrl` 파일 파싱.  
별도(Separate) 기준, 당기말(instant) 또는 당기(duration) context 사용.

| 항목 | Col | XBRL 태그 / 비고 |
|------|-----|----------------|
| **재무상태표 (BS)** | | |
| 총자산 | 4 | `ifrs-full_Assets` |
| FVPL | 5 | `ifrs-full_FinancialAssetsAtFairValueThroughProfitOrLoss` |
| FVOCI | 6 | `ifrs-full_FinancialAssetsAtFairValueThroughOtherComprehensiveIncome` |
| AC | 7 | `ifrs-full_FinancialAssetsAtAmortisedCost` |
| 종속·관계기업 | 8 | `ifrs-full_InvestmentsInSubsidiaries...` |
| 부채 | 12 | `ifrs-full_Liabilities` |
| 보험계약부채 | 13 | `ifrs-full_InsuranceContractsIssuedThatAreLiabilities` |
| 자본(신종자본증권 제외) | 14 | Col15 - Col16 계산 |
| 자본 | 15 | `ifrs-full_Equity` |
| 신종자본증권 | 16 | `dart_HybridBonds` |
| 기타포괄손익누계액 | 17 | `ifrs-full_AccumulatedOtherComprehensiveIncome` |
| 이익잉여금 | 18 | `ifrs-full_RetainedEarnings` |
| **손익계산서 (PL)** | | |
| 보험손익 | 32 | `ifrs-full_InsuranceServiceResult` |
| 보험서비스수익 | 33 | `dart_OperatingIncomeInsurance` |
| 보험서비스비용 | 34 | `dart_OperatingExpenseInsurance` |
| 투자손익 | 35 | `dart_InvestmentIncomeExpenses` |
| 보험금융손익 | 36 | dart_Insurance Finance Income - Expense |
| 재보험금융손익 | 37 | dart_Finance Income/Expense From Reinsurance |
| 재산관리비 | 39 | `ifrs-full_SellingGeneralAndAdministrativeExpense` |
| 영업이익 | 41 | `ifrs-full_ProfitLossFromOperatingActivities` |
| 세전손익 | 43 | `ifrs-full_ProfitLossBeforeTax` |
| 당기순이익 | 44 | `ifrs-full_ProfitLoss` |
| **자본변동표** | | |
| 기초자본 | 19-20 | 전기말 `ifrs-full_Equity` |
| 배당 | 66 | `dart_DividendsPaidClassifiedAsFinancingActivities` |
| 자기주식 소각 | 67 | `ifrs-full_CancellationOfTreasuryShares` (RetainedEarnings component) |
| OCI→이익잉여금 재분류 | 69 | `dart_TransferOfAmountRecognised...` (RetainedEarnings component) |
| 신종자본증권 배당 | 70 | `dart_DividendToHybridBond` |
| **OCI 누계 (BS 3-dim context)** | | |
| OCI 합계 | 24 | = Col17 |
| FVOCI평가손익 | 25 | FinancialAssets 키워드 멤버 |
| 보험계약금융손익 | 26 | InsuranceContract 키워드 멤버 |
| 재보험계약금융손익 | 27 | ReinsuranceFinance 키워드 멤버 |
| 현금흐름위험회피 | 28 | CashFlowHedges 키워드 멤버 |
| 확정급여채무 재측정 | 30 | RemeasurementsOfDefinedBenefit 멤버 |
| **BEL/RA/CSM (별도 multi-dim context)** | | |
| BEL | 73 | 보험계약부채 현재가치 추정치 축 |
| RA | 74 | 위험조정 축 |
| CSM | 75 | 보험계약마진 축 |
| 신계약CSM | 76 | ChangesInFutureServicesDueToNewContracts 태그 |
| CSM상각 | 77 | 서비스이전_보험계약마진 축 |
| CSM조정 | 78 | 보험계약마진조정_추정치변동 축 |
| 보험금융손익(CSM) | 79 | 보험금융손익 축 |
| 전기CSM | 80 | 전기말 context |
| **계리적 가정변경 민감도 (3-dim context)** | | |
| BEL+RA 기준 해지율/위험률/사업비/기타/물량 | 146-150 | `dart_ChangeIn...Assumption...BEL component` |
| CSM 기준 해지율/위험률/사업비/기타/물량/손실 | 153-158 | `dart_ChangeIn...Assumption...CSM component` |
| **K-ICS 최종치** | | |
| K-ICS비율/가용자본/요구자본 (전기) | 95-97 | `data/kics_2512.csv` (수동 관리) |
| **해약환급금준비금** | 99 | `dart_SurrenderValueReserve` |

---

### 2. 사업보고서 document.xml — `_fetch_document_content()` + doc.xml 파서

DART API `document.xml`으로 다운로드. HTML 테이블 embedded XML에서 파싱.  
`data/doc_annual_{회사명}_{연도}.xml`에 캐시 저장.

| 항목 | Col | 추출 방법 | 성공 회사 |
|------|-----|---------|---------|
| **운용자산 / 총자산(별도) / 자산운용률** | 21-23 | `_extract_investment_assets` — 운용자산/자산운용률 테이블 파싱 | 11개사 |
| **K-ICS 잠정치** | 92-94 | `_extract_kics` — 지급여력비율 테이블 파싱, 억원 단위 자동 감지 | 7개사 |
| **보험금 예실차비율** | 101-104 | `_extract_loss_ratios` — Method1: TE ACODE 태그, Method2: 테이블 파싱 | 10개사 |
| **CSM 기간별 기대수익인식금액** | 83-88 | `_extract_csm_maturity` — CSM 만기 테이블, 다양한 버킷 구조 대응 | 교보/한화/동양/DB손보/삼성화재 |
| **계리적 가정변경 감응도** | 146-157 | `_extract_sensitivity` — 가정변경효과 테이블, 억원 단위 감지 | 교보/신한라이프/KB생명/메리츠 등 |
| **OCI 세부 누계잔액** | 25-30 | `_extract_oci_detail` — BS 세부항목 테이블, OCI합계 기준 섹션 선택 | 교보/KB손보/KB생명/메리츠 등 |

---

### 3. 1분기보고서 document.xml (XBRL axis 없는 5개사 fallback)

교보/신한라이프/KB생명/메리츠화재/KB손해보험은 XBRL에 BEL/RA/CSM 분리 축이 없음.  
1Q 사업보고서(2026년 제출) document.xml에서 추출.  
`data/doc_1q_{회사명}_{연도}.xml`에 캐시.

| 항목 | Col | 추출 방법 |
|------|-----|---------|
| BEL/RA/CSM 기말잔액 | 73-75 | `parse_insurance_components_from_doc` — 보험계약부채 구성요소 테이블 |
| CSM 변동표 항목 | 76-80 | `parse_csm_movement_from_doc` — 연간 doc.xml CSM 변동표, target 기반 당기 테이블 선택 |

---

### 4. 계산값 (다른 컬럼으로부터 파생)

| 항목 | Col | 계산식 |
|------|-----|-------|
| 총자산대비비중 FVPL/FVOCI/AC | 9-11 | Col5~7 / Col4 |
| 기초자본(신종자본증권 제외) | 20 | 전기말 Col15 - 전기말 Col16 |
| 금융손익 | 38 | Col35 - Col36 - Col37 + Col39 - Col40 |
| OCI 체크 | 31 | Col24 - (Col25+26+27+28+29+30) |
| CSM 증감 체크 | 82 | Col80+76-77+78+79-Col75 |
| CSM 합계 | 88 | = Col75 |
| CSM 기간별 합계 체크 | 89 | Col88 - Σ(Col83~87) |
| 예실차비율 | 103 | Col101 - Col102 (A-B) |
| 실제/예상보험금 비율 | 104 | Col102 / Col101 |

---

## 데이터 소스 파악 흐름

```
DART API
├── fnlttXbrl.xml → .xbrl 파일
│   ├── BS/PL/자본변동 → Col4~71, 95~99
│   ├── OCI 누계 (3-dim) → Col24~30
│   ├── BEL/RA/CSM (multi-dim) → Col73~80
│   ├── 계리적 가정변경 (3-dim) → Col146~158
│   └── 자기주식/OCI재분류/배당 (RetainedEarnings component) → Col67~70
│
└── document.xml → HTML 테이블 파싱
    ├── 사업보고서 (FY2025, 2026년 제출)
    │   ├── 운용자산 테이블 → Col21~23
    │   ├── K-ICS 지급여력 테이블 → Col92~94
    │   ├── 보험금 예실차 테이블 → Col101~104
    │   ├── CSM 기간별 테이블 → Col83~88
    │   ├── 가정변경 감응도 테이블 → Col146~157
    │   └── OCI 세부잔액 테이블 → Col25~30 (fallback)
    └── 1분기보고서 (2026년 제출, 5개사만)
        ├── BEL/RA/CSM 구성요소 테이블 → Col73~75
        └── CSM 변동표 → Col76~80
```

---

## 현재 미추출 항목 (219건 MISS)

| 항목 | Col | 이유 |
|------|-----|------|
| 22년말(IFRS4) | 72 | XBRL/doc.xml에 없음 |
| 예실차 세부 | 111-117 | 보험종류별 다차원 집계, 별도 처리 필요 |
| CSM기간별 (일부 회사) | 83-88 | 신한라이프/삼성생명/KB라이프/미래에셋 등 doc.xml 미포함 |
| 위험보험료/사고보험금 | 107-109 | 다차원 집계 복잡 |
| 재보험OCI 누계 | 27 | 일부 회사 3-dim context 불완전 |

---

## 주요 기술 사항

### XBRL 파싱 context 구조
- **1-dim (Separate)**: 대부분 BS/PL 항목
- **2-dim (Separate + ComponentsOfEquity)**: 자본변동표 항목 (Col67~70)
- **3-dim (Separate + OCI member + entity-specific)**: OCI 누계 (Col24~30)
- **Multi-dim (Separate + Insurance axis + CSM axis + ...)**: BEL/RA/CSM, 감응도

### document.xml 섹션 분리
- 연결/별도 구분: 파일 내 위치 gap > 500KB 기준
- 당기/전기 구분: target 값 (CSM기말 등) 기반 최근접 테이블 선택

### 단위 변환
- XBRL: 원(KRW) → 백만원 (`/ 1e6`)
- doc.xml: 회사마다 억원 또는 백만원 → 자동 감지 후 변환
