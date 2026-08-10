# 동업사 공시비교 DATA 시트 자동화

DART 전자공시시스템에서 12개 보험사의 XBRL 파일과 원문 XML(document.xml)을 받아  
금융감독원 FISIS API와 연계하여 `DATA_작업_빈칸.xlsx`의 DATA 시트를 자동으로 채웁니다.

## 실행 방법

```bash
# 전체 기간 일괄 처리 (Excel에 있는 모든 기간코드)
python parse_annual.py --period all

# 특정 기간만
python parse_annual.py --period 2512    # 2025년 12월 사업보고서
python parse_annual.py --period 2509    # 2025년 9월 3분기보고서
python parse_annual.py --period 2506    # 2025년 6월 반기보고서
python parse_annual.py --period 2503    # 2025년 3월 1분기보고서
python parse_annual.py --period 2412    # 2024년 12월 사업보고서
```

**XBRL / document.xml이 없으면 DART API에서 자동 다운로드합니다.**  
결과는 `data/` 폴더에 캐시되어 다음 실행 시 재사용됩니다.

---

## 데이터 소스

| 소스 | 설명 | 주요 항목 |
|------|------|---------|
| **XBRL** | DART `fnlttXbrl.xml` API — 별도 재무제표 기준 `.xbrl` 파일 | BS·PL·CSM변동·OCI변동·감응도 |
| **원문XML** | DART `document.xml` API — 보고서 원문 HTML 테이블 파싱 | K-ICS·운용자산·CSM기간별·손해율 |
| **FISIS API** | 금융감독원 금융통계정보시스템 Open API | RA·CSM잔액·해약환급금·운용자산·신종자본증권·킥스(분기) |

### FISIS API 키 설정

`dart_api.py`의 `FISIS_API_KEY` 변수에 인증키를 입력합니다.  
키 발급: https://fisis.fss.or.kr/page/api-key.jsp (무료, 이메일 수령)

```python
FISIS_API_KEY = "your_api_key_here"
```

---

## 전체 정확도 요약

정답 데이터(`202512_동업사 공시비교_DATA_작업_260320.xlsx`)와 비교한 결과입니다.  
오차 허용 기준: **0.1% 이내** = OK, 초과 = FAIL, 미추출 = MISS

| 기간 | 보고서 종류 | OK | FAIL | MISS | 비교대상 | **OK율** |
|------|-----------|---:|-----:|-----:|--------:|--------:|
| 2212 | 2022년 사업보고서 | 0 | 0 | 144 | 144 | **0%** |
| 2312 | 2023년 사업보고서 | 477 | 213 | 275 | 965 | **49%** |
| 2412 | 2024년 사업보고서 | 742 | 80 | 286 | 1,108 | **67%** |
| 2503 | 2025년 1분기 | 573 | 60 | 297 | 930 | **62%** |
| 2506 | 2025년 반기 | 607 | 93 | 256 | 956 | **64%** |
| 2509 | 2025년 3분기 | 673 | 49 | 230 | 952 | **71%** |
| 2512 | 2025년 사업보고서 | 942 | 34 | 137 | 1,113 | **85%** |

**기간별 특이사항**

- **2212 (0%)**: IFRS17 도입 이전(IFRS4 기준)으로 태그 체계가 완전히 달라 현재 파싱 패턴 미지원. 수동 입력 필요.
- **2312 (49%)**: IFRS17 도입 첫 해로 XBRL 태그 구조가 정착되지 않아 패턴 불일치 다수. 전기말(2022년말) XBRL 미제공으로 이전기 비교수치 불일치.
- **2412 (67%)**: XBRL에 BEL/RA/CSM 구성요소 태그가 없는 회사 다수 (2025년부터 확대 공시). FISIS로 보완.
- **2503·2506·2509 (62~71%)**: 분기/반기는 사업보고서 전용 항목(CSM변동·감응도·CSM기간별) 정답 자체가 없어 비교 대상이 적음. K-ICS는 FISIS API로 채움.
- **2512 (85%)**: 사업보고서 전용 항목 포함 최고 정확도. 주요 MISS: 예실차 세부(Col111~117), 일부 CSM기간별(Col83~87), OCI누계 세부(Col26~30 일부).

> **2412 BEL/CSM MISS 원인**: 삼성생명 기준 2024년 XBRL 태그 수 245개(BEL/RA/CSM 분리 없음) → 2025년 988개(구성요소 분리 태그 신설). 2025년부터 대부분 회사가 XBRL 공시 수준을 크게 확대해 해결됨.

---

## 컬럼별 자동화 현황

> **기준: 2512 (2025년 12월 사업보고서), 12개사**  
> 분기/반기(2503·2506·2509)는 BS·PL 기본 항목 + FISIS 연계 항목 위주.  
> 사업보고서 전용(CSM변동·감응도·CSM기간별·손해율)은 연간에서만 추출.

---

### B/S (재무상태표)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 4 | 총자산 | XBRL | `ifrs-full_Assets` | ✅ 12/12 |
| 5 | FVPL | XBRL | `ifrs-full_FinancialAssetsAtFairValueThroughProfitOrLoss` | ✅ 12/12 |
| 6 | FVOCI | XBRL | `ifrs-full_FinancialAssetsAtFairValueThroughOtherComprehensiveIncome` | ✅ 12/12 |
| 7 | AC | XBRL | `ifrs-full_FinancialAssetsAtAmortisedCost` | ✅ 12/12 |
| 8 | 종속·관계기업 | XBRL | `ifrs-full_InvestmentsInSubsidiariesJointVenturesAndAssociates` | ✅ 11/11 |
| 9~11 | 총자산대비비중 | 계산 | Col5~7 / Col4 | ✅ |
| 12 | 부채 | XBRL | `ifrs-full_Liabilities` | ✅ 12/12 |
| 13 | 보험계약부채 | XBRL | `ifrs-full_InsuranceContractsIssuedThatAreLiabilities` | ✅ 12/12 |
| 14 | 자본(신종제외) | 계산 | Col15 − Col16 | ✅ 12/12 |
| 15 | 자본 | XBRL | `ifrs-full_Equity` | ✅ 12/12 |
| 16 | 신종자본증권 | XBRL → **FISIS** | `dart_HybridBonds`; fallback: FISIS SH151/SI147 F3 | ✅ 5/5 (발행 회사만) |
| 17 | OCI누계액 | XBRL | `ifrs-full_AccumulatedOtherComprehensiveIncome` | ✅ 12/12 |
| 18 | 이익잉여금 | XBRL | `ifrs-full_RetainedEarnings` | ✅ 12/12 |

---

### 자산운용

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 21 | 자산운용 총자산 | 원문XML → XBRL | 운용자산 테이블 `총 자 산(A)`; fallback: Col4(총자산) | ⚠️ 일부 회사 MISS |
| 22 | 운용자산 | **FISIS** → 원문XML | FISIS SH150/SI146 A1 (일반계정+특별계정 자동 판별); fallback: 원문XML 운용자산 테이블 | ✅ 11/12 (삼성화재 소오차) |
| 23 | 자산운용률 | 원문XML | 운용자산 테이블 `자산운용률(B/A)` | ⚠️ 일부 회사 MISS |

> FISIS 운용자산(SH150 A1)은 일반계정 기준. 특별계정 별도 집계 회사(한화 등)는 SH152와 합산 여부를 자동 판별합니다.  
> `|A1 + 특별계정| ≈ 총자산(1.5% 이내)`이면 합산, 그렇지 않으면 A1만 사용.

---

### 기타포괄손익누계액 (OCI 잔액)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 24 | OCI 합계 | 계산 | = Col17 | ✅ 12/12 |
| 25 | FVOCI 평가손익 | XBRL→원문XML | XBRL OCI 잔액; fallback: 자본 세부 테이블 | ⚠️ 3개사 MISS |
| 26 | 보험계약 금융손익 | XBRL→원문XML | XBRL 보험계약 OCI 잔액; fallback: 자본 세부 테이블 | ⚠️ 5개사 MISS |
| 27 | 재보험계약 금융손익 | XBRL→원문XML | XBRL 재보험계약 OCI 잔액; fallback: 자본 세부 테이블 | ⚠️ 8개사 MISS |
| 28 | 현금흐름위험회피 | XBRL→원문XML | XBRL 현금흐름위험회피 OCI 잔액; fallback: 자본 세부 테이블 | ⚠️ 6개사 MISS |
| 29 | 재평가잉여금 | XBRL→원문XML | XBRL 전기말 재평가잉여금; fallback: 자본 세부 테이블 | ⚠️ 3개사 MISS |
| 30 | 확정급여부채 재측정 | XBRL→원문XML | XBRL 확정급여부채 재측정요소 OCI; fallback: 자본 세부 테이블 | ⚠️ 4개사 MISS |

---

### 손익계산서 (P/L)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 32 | 보험손익 | XBRL | `ifrs-full_InsuranceServiceResult` | ✅ 12/12 |
| 33 | 보험서비스수익 | XBRL | `dart_OperatingIncomeInsurance` | ✅ 12/12 |
| 34 | 보험서비스비용 | XBRL | `dart_OperatingExpenseInsurance` | ✅ 12/12 |
| 35 | 투자손익 | XBRL | `dart_InvestmentIncomeExpenses` | ✅ 12/12 |
| 36 | 보험금융손익 | XBRL | Insurance finance income − expense | ✅ 12/12 |
| 37 | 재보험금융손익 | XBRL | Reinsurance finance income − expense | ⚠️ 교보 오차 |
| 38 | 금융손익 | 계산 | Col35 − Col36 − Col37 + Col39 − Col40 | ⚠️ 일부 소오차 |
| 39 | 재산관리비 | XBRL | `ifrs-full_SellingGeneralAndAdministrativeExpense` | ✅ 11/11 |
| 40 | 기타투자손익 | XBRL | 수수료+임대료+기타수익−기타비용 합산 | ⚠️ 일부 소오차 |
| 41 | 영업이익 | XBRL | `ifrs-full_ProfitLossFromOperatingActivities` | ✅ 12/12 |
| 42 | 영업외손익 | XBRL | `ifrs-full_NonOperatingProfitLoss` | ✅ 12/12 |
| 43 | 세전손익 | XBRL | `ifrs-full_ProfitLossBeforeTax` | ✅ 12/12 |
| 44 | 당기순이익 | XBRL | `ifrs-full_ProfitLoss` | ✅ 12/12 |

---

### 자본변동표 — OCI 변동분

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 55 | OCI 기초잔액 | XBRL | ✅ 12/12 |
| 56 | FVOCI 평가손익 변동 | XBRL | ⚠️ 동양·현대해상 소오차 |
| 57 | 대손충당금 변동 | XBRL | ❌ 대부분 MISS (해당 회사 소액) |
| 58 | 보험계약금융손익 OCI | XBRL | ✅ 12/12 |
| 59 | 현금흐름위험회피 변동 | XBRL | ⚠️ 교보 소오차 |
| 60 | 재평가잉여금 변동 | XBRL | ⚠️ 동양·현대해상·DB손보 오차 |
| 61 | 확정급여부채 재측정 | XBRL | ✅ 12/12 |
| 62 | 해외사업환산손익 | XBRL | ⚠️ 동양 MISS |
| 63 | OCI 기말잔액 | 계산 | = Col17 ✅ |

---

### 자본변동표 — 이익잉여금

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 65 | 이익잉여금 기초 | XBRL | ⚠️ 메리츠화재 소오차 |
| 66 | 결산배당 | XBRL | ✅ 8/8 (배당 회사만) |
| 67 | 자기주식 소각/처분 | XBRL | ✅ 2/2 (해당 회사만) |
| 68 | 순이익 | 계산 | = Col44 ✅ |
| 69 | 기타 자본변동 | XBRL | ⚠️ 동양 소오차·KB손보 오류 |
| 70 | 신종자본증권 배당 | XBRL | ⚠️ 메리츠 소오차·현대해상 MISS |
| 71 | 이익잉여금 기말 | 계산 | = Col18 ✅ |

---

### 보험계약부채 구성요소

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 73 | BEL | XBRL → 원문XML | XBRL 보험계약부채 구성요소 BEL. fallback: 1분기보고서 원문XML 구성요소 테이블 | ✅ 12/12 |
| 74 | RA | XBRL → 원문XML → **FISIS** | 동일; XBRL/원문XML 실패 시 FISIS SH156/SI152 A112+A122 | ✅ 12/12 |
| 75 | CSM | XBRL → 원문XML → **FISIS** | 동일; XBRL/원문XML 실패 시 FISIS SH156/SI152 A113 | ✅ 12/12 |
| 76 | 신계약CSM | XBRL → 원문XML | XBRL CSM 변동 주석; fallback: 사업보고서 원문XML CSM 변동표 | ⚠️ 동양 MISS |
| 77 | CSM 상각 | XBRL → 원문XML | 동일 | ⚠️ 동양 MISS |
| 78 | CSM 조정 | XBRL → 원문XML | 동일 | ✅ 12/12 |
| 79 | 보험금융손익(CSM) | XBRL → 원문XML | 동일 | ⚠️ DB손보 MISS |
| 80 | 전기말 CSM | XBRL → 원문XML | XBRL 전기말 CSM; fallback: 원문XML 변동표 `기초` 행 | ⚠️ 삼성화재·DB손보 MISS |

> **RA·CSM FISIS fallback**: 분기/반기는 FISIS 값을 우선 사용(XBRL보다 정확). 사업보고서는 XBRL/원문XML에서 뽑지 못한 경우만 FISIS 사용.  
> **손보사 RA 오차**: FISIS SI152 RA 집계 기준이 엑셀 기준(재보험 제외)과 5~10% 차이 있어 XBRL 우선.

---

### CSM 기간별 기대수익인식금액

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 83~87 | 1년이하~10년초과 | 원문XML | ⚠️ 교보·한화·동양·DB손보·삼성화재 OK / 나머지 7개사 MISS (테이블 미공시) |
| 88 | 합계 | 계산 | = Col75 ✅ 11/11 |

---

### K-ICS (지급여력비율)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 92~94 | 잠정치 (비율·가용·요구) | 원문XML | 사업보고서 지급여력비율 테이블 | ⚠️ 일부 회사 MISS (산출중) |
| 95~97 | 최종치 (비율·가용·요구) | 원문XML → **FISIS** | 사업보고서는 1분기보고서 원문XML 전년도 열; 분기/반기는 FISIS SH021/SI021 | ✅ 2512 12/12, 분기 12/12 |

> - **사업보고서(연간)**: 다음 해 1분기보고서(`doc_1q_{year+1}.xml`)의 "전기" 열에서 확정치 추출.  
> - **분기/반기**: FISIS API 해당 분기 직접 조회. 잠정치(Col92~94)는 검증 제외.

---

### 해약환급금준비금

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 99 | 해약환급금준비금 | XBRL → **FISIS** | `dart_SurrenderValueReserve`; FISIS SH151/SI147 F46로 항상 갱신 | ✅ 12/12 |

> XBRL보다 FISIS 값이 더 안정적이므로 FISIS 값을 우선 적용합니다.

---

### 손해율 · 예실차

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 101 | 예상손해율(A) | 원문XML | ⚠️ 미래에셋·현대해상 MISS |
| 102 | 실제손해율(B) | 원문XML | ⚠️ 미래에셋·현대해상 MISS |
| 103 | 예실차비율(b−a) | 계산 | Col101 − Col102 |
| 104 | 실제/예상보험금 | 계산 | Col102 / Col101 |
| 107 | 손해율(누계) | — | ❌ 전 회사 MISS |
| 108 | 위험보험료 | — | ❌ 전 회사 MISS |
| 109 | 사고보험금 | — | ❌ 전 회사 MISS |
| 111 | 예실차 합계 | — | ❌ 전 회사 MISS |
| 112~117 | 예실차 세부 | — | ❌ 전 회사 MISS |

> **107~117 MISS 이유**: XBRL 태그 없음. FISIS에도 해당 통계표 미존재. 원문XML에서도 보험종류별 분리 집계라 자동 합산이 어려움. **수동 입력 필요.**

---

### 계리적 가정변경 민감도

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 146~150 | BEL+RA기준 (해지율·위험률·사업비·기타·물량) | XBRL → 원문XML | ⚠️ 동양·삼성화재 MISS, 한화 부호오류 |
| 153~158 | CSM기준 (해지율~손실요소) | XBRL → 원문XML | ⚠️ 동양·삼성화재 MISS |

---

## XBRL Taxonomy

`data/taxonomy.csv` — XBRL 태그 ID와 한글명 매핑표.

| 분류 | prefix | 개수 | 설명 |
|------|--------|-----:|------|
| IFRS 표준 | `ifrs-full_` | 1,562 | IASB 정의 국제 표준 태그 |
| DART 한국 추가 | `dart_` | 967 | 금융감독원 보험업 특화 추가 태그 |
| 회사별 커스텀 | `entity{corp_code}_` | 10,844 | 각 회사 자체 정의 태그 |

---

## 회사마다 결과가 다른 이유

### 이유 1 — 같은 항목을 저장하는 XBRL 구조가 회사마다 다르다

IFRS17은 어떤 필터(dim)를 얼마나 걸지 회사 재량이라 같은 항목인데 구조가 다릅니다.

**CSM (Col75) 예시:**

| 회사 | Dim 수 | 구조 |
|------|--------|------|
| 삼성생명 | 2-dim | `SeparateMember` + `InsuranceContractsIssuedMember`, 태그=`ContractualServiceMargin` |
| 현대해상 | 5-dim | 위에 `PAA미적용` + `LRC` + `CSM컴포넌트` 필터 추가 |

스크립트는 조건을 순서대로 시도해 처음 매칭되는 패턴을 사용합니다. 패턴을 한 번 정의하면 `--period all`로 전 기간 일괄 처리됩니다.

**2412 BEL/CSM MISS 원인:** 2024년까지 XBRL에 BEL/RA/CSM 분리 태그가 없던 회사들이 많음. 2025년 사업보고서부터 대부분 확대 공시.

### 이유 2 — 원문 XML에 테이블 자체가 없거나 형태가 다르다

공시 의무 항목이 아닌 경우 회사 재량으로 작성 여부·형태가 결정됩니다.

- **운용자산(Col22)**: 생보사는 대부분 테이블 있음 → FISIS 우선 적용. 손보사 일부 원문에만 있음.
- **CSM기간별(Col83~87)**: 교보·한화·동양·DB손보·삼성화재만 공시. 신한라이프·삼성 등 7개사는 미공시.
- **예실차(Col111~117)**: XBRL 태그 없음, FISIS에도 없음 → 전 회사 수동 입력 필요.

---

## 파일 구조

```
parse_annual.py       메인 실행 스크립트
dart_api.py           DART OpenAPI + FISIS API 연동
verify_2512.py        정답 엑셀 대비 검증 스크립트
DATA_작업_빈칸.xlsx    채울 대상 Excel
data/
  taxonomy.csv                   XBRL 태그 한글명 매핑 (자동 생성)
  data_sheet_{YYMM}.csv          기간별 추출 결과
  {회사명}/
    사업보고서/{연도}/             XBRL 파일 캐시
    doc_annual_{연도}.xml         사업보고서 원문 캐시
    doc_1q_{연도}.xml             1분기보고서 원문 캐시 (K-ICS 확정치용)
    doc_halfyear_{연도}.xml       반기보고서 원문 캐시
    doc_3q_{연도}.xml             3분기보고서 원문 캐시
```
