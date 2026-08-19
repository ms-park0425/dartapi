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
| **FISIS API** | 금융감독원 금융통계정보시스템 Open API | RA·CSM·해약환급금·운용자산·신종자본증권·OCI세부(Col26-30)·킥스(분기) |

### FISIS API 키 설정

`dart_api.py`의 `FISIS_API_KEY` 변수에 인증키를 입력합니다.  
키 발급: https://fisis.fss.or.kr/page/api-key.jsp (무료, 이메일 수령)

```python
FISIS_API_KEY = "your_api_key_here"
```

---

## XBRL 파일 구조

XBRL 파일은 크게 두 부분으로 나뉩니다.

**① 상단 — context 정의 블록**: "이 숫자가 어떤 조건의 값인가"를 정의합니다.

```xml
<context id="CFY2025eFY_Separate_IssuedMember">
  <entity><identifier>00108465</identifier></entity>
  <period>
    <startDate>2025-01-01</startDate>
    <endDate>2025-12-31</endDate>
  </period>
  <scenario>
    <xbrldi:explicitMember dimension="ConsolidatedAndSeparateFinancialStatementsAxis">
      ifrs-full:SeparateMember
    </xbrldi:explicitMember>
    <xbrldi:explicitMember dimension="DisaggregationOfInsuranceContractsAxis">
      ifrs-full:InsuranceContractsIssuedMember
    </xbrldi:explicitMember>
  </scenario>
</context>
```

> `dimension`이 필터 종류, 그 안의 값이 필터 조건입니다. 이 필터 묶음을 **dim**이라 부릅니다.

**② 하단 — 값 태그**: `contextRef`로 위의 context를 참조해 실제 숫자를 저장합니다.

```xml
<ifrs-full:ContractualServiceMargin
  contextRef="CFY2025eFY_Separate_IssuedMember"
  decimals="-6">13217874000000</ifrs-full:ContractualServiceMargin>
```

> `contextRef`가 가리키는 context id의 조건(별도 재무제표, 발행계약 전체)에 해당하는  
> CSM 잔액이 **13,217,874,000,000원** 이라는 뜻입니다.

스크립트는 context id 문자열을 직접 매칭하지 않고, **각 context 안의 dimension 목록을 읽어 조건을 확인**합니다.  
회사마다 id 문자열이 달라도 dimension 조건이 같으면 동일하게 처리됩니다.

### XBRL Taxonomy

`data/taxonomy.csv` — XBRL 태그 ID와 한글명 매핑표. 각 회사 XBRL 다운로드 시 포함된 `*_lab-ko.xml`에서 자동 추출됩니다.

| 분류 | prefix | 개수 | 설명 |
|------|--------|-----:|------|
| IFRS 표준 | `ifrs-full_` | 1,562 | IASB가 정의한 국제 표준 태그 |
| DART 한국 추가 | `dart_` | 967 | 금융감독원 보험업 특화 추가 정의 |
| 회사별 커스텀 | `entity{corp_code}_` | 10,844 | 각 회사가 자체 정의한 태그 (비표준, 다른 회사/연도에 미적용) |

**주요 활용 태그:**

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

## 전체 정확도 요약

정답 데이터(`202512_동업사 공시비교_DATA_작업_260320.xlsx`)와 비교한 결과입니다.  
오차 허용 기준: **절대차 1 이하 또는 상대오차 0.1% 이내** = OK, 초과 = FAIL, 미추출 = MISS  
FAIL·MISS 셀은 `DATA_작업_빈칸.xlsx`에 주황색으로 자동 표시됩니다.

| 기간 | 보고서 종류 | OK | FAIL | MISS | 비교대상 | **OK율** |
|------|-----------|---:|-----:|-----:|--------:|--------:|
| 2312 | 2023년 사업보고서 | 517 | 201 | 237 | 955 | **54%** |
| 2412 | 2024년 사업보고서 | 775 | 73 | 248 | 1,096 | **71%** |
| 2503 | 2025년 1분기 | 623 | 29 | 278 | 930 | **67%** |
| 2506 | 2025년 반기 | 667 | 41 | 248 | 956 | **70%** |
| 2509 | 2025년 3분기 | 760 | 22 | 170 | 952 | **80%** |
| 2512 | 2025년 사업보고서 | 978 | 18 | 117 | 1,113 | **88%** |

**기간별 특이사항**

- **2312 (53%)**: IFRS17 도입 첫 해로 XBRL 태그 구조가 정착되지 않아 패턴 불일치 다수. 전기말(2022년말) XBRL이 IFRS17 재작성 기준이라 이전기 비교수치 불일치.
- **2412 (70%)**: XBRL에 BEL/RA/CSM 구성요소 태그가 없는 회사 다수(2025년부터 확대 공시). FISIS로 보완. OCI세부(Col26-30)도 FISIS 추가.
- **2503 (67%)**: 1분기 보고서는 CSM변동·OCI누계·BEL 등 사업보고서 전용 항목이 없어 비교 대상 자체가 적음. K-ICS·RA·CSM·해약환급금·OCI세부는 FISIS로 채움.
- **2506 (68%)**: 반기 보고서. 손보사 Col36~40(보험금융/재보험금융/금융손익/재산관리비/기타투자) 반기 기재 방식 차이 처리.
- **2509 (80%)**: 3분기 보고서. doc_3q XML에서 메리츠·KB손보 RA 직접 추출. 한화·동양·미래에셋 분기 BEL 추가 패턴. XBRL MaturityAxis CSM기간별 적용.
- **2512 (88%)**: 사업보고서 전용 항목 포함 최고 정확도. XBRL MaturityAxis에서 삼성생명·한화생명·동양생명·미래에셋 CSM기간별 직접 추출. 주요 MISS: 예실차 세부(Col111-117), CSM기간별(KB손보).

> **2412 BEL/CSM MISS 원인**: 삼성생명 기준 2024년 XBRL 태그 수 245개(BEL/RA/CSM 분리 없음) → 2025년 988개(구성요소 분리 태그 신설). 2025년부터 대부분 회사가 XBRL 공시 수준을 크게 확대해 해결됨.

---

## 회사마다 결과가 다른 이유

자동화가 안 되거나 회사마다 결과가 다른 이유는 크게 두 가지입니다.

---

### 이유 1 — XBRL 안에서 같은 값을 저장하는 구조가 회사마다 다르다

XBRL은 숫자 하나를 저장할 때 **"이 숫자가 어떤 조건의 값인가"를 함께 태깅**합니다.  
엑셀 피벗테이블에서 필터를 걸어 특정 셀 값을 뽑는 것과 같고, 그 필터 개수를 **dim**이라고 부릅니다.

IFRS17은 어떤 필터를 얼마나 걸지 회사 재량이라, **같은 항목인데 회사마다 필터 구조가 다릅니다.**

#### DATA 시트 Col75 (CSM 잔액) — 같은 숫자를 찾는 필터가 회사마다 다르다

**삼성생명 — 2-dim: 필터 2개, 값 하나로 바로 저장**

```
필터① 재무제표 종류  = 별도 (SeparateMember)
필터② 보험계약 구분  = 발행계약 전체 (InsuranceContractsIssuedMember)
      + 태그 자체가 ContractualServiceMargin → CSM임을 의미
```

실제 XBRL:
```xml
<ifrs-full:ContractualServiceMargin
  contextRef="CFY2025eFY_..._SeparateMember_..._InsuranceContractsIssuedMember"
  decimals="-6">13217874000000</ifrs-full:ContractualServiceMargin>
```
contextRef가 가리키는 context 정의:
```xml
<xbrldi:explicitMember dimension="ConsolidatedAndSeparateFinancialStatementsAxis">ifrs-full:SeparateMember</xbrldi:explicitMember>
<xbrldi:explicitMember dimension="DisaggregationOfInsuranceContractsAxis">ifrs-full:InsuranceContractsIssuedMember</xbrldi:explicitMember>
```
→ **13,217,874,000,000원 = 13,217,874 백만원** ✅

---

**현대해상 — 5-dim: 필터 5개, 값 하나로 저장 (내부적으로는 6-dim 세부 내역도 병존)**

```
필터① 재무제표 종류  = 별도
필터② 보험계약 구분  = 발행계약 전체
필터③ 회계모형       = PAA 미적용 계약
필터④ 잔여보장       = LRC (잔여보장부채)
필터⑤ 구성요소       = CSM
```

실제 XBRL (5-dim 합계 태그):
```xml
<ifrs-full:InsuranceContractsLiabilityAsset
  contextRef="CFY2025eFY_..._SeparateMember_..._InsuranceContractsIssuedMember_..._InsuranceContractsOtherThanThoseToWhichPremiumAllocationApproachHasBeenAppliedMember_..._LiabilitiesForRemainingCoverageMember_..._ContractualServiceMarginMember"
  decimals="0">8977842069322</ifrs-full:InsuranceContractsLiabilityAsset>
```
contextRef가 가리키는 context 정의:
```xml
<xbrldi:explicitMember dimension="ConsolidatedAndSeparateFinancialStatementsAxis">ifrs-full:SeparateMember</xbrldi:explicitMember>
<xbrldi:explicitMember dimension="DisaggregationOfInsuranceContractsAxis">ifrs-full:InsuranceContractsIssuedMember</xbrldi:explicitMember>
<xbrldi:explicitMember dimension="InsuranceContractsAxis">ifrs-full:InsuranceContractsOtherThanThoseToWhichPremiumAllocationApproachHasBeenAppliedMember</xbrldi:explicitMember>
<xbrldi:explicitMember dimension="InsuranceContractsByRemainingCoverageAndIncurredClaimsAxis">dart:LiabilitiesForRemainingCoverageMember</xbrldi:explicitMember>
<xbrldi:explicitMember dimension="InsuranceContractsByComponentsAxis">ifrs-full:ContractualServiceMarginMember</xbrldi:explicitMember>
```
→ **8,977,842,069,322원 = 8,977,842 백만원** ✅

같은 파일 안에 6-dim 세부 내역도 함께 존재 (필터⑤ 뒤에 보험종류 필터가 추가됨):
```
장기보험 (비배당) → 8,868,911,078,887 원
장기보험 (배당)   →   108,930,990,435 원
일반보험          →               0 원
자동차보험        →               0 원
                 ───────────────────────
                   8,977,842,069,322 원  ← 5-dim 합계와 일치
```

스크립트는 5-dim 합계 태그를 직접 읽고, 없는 경우 6-dim 세부를 합산하는 방식으로 동작합니다.

#### 스크립트가 어떻게 올바른 context를 찾는가

XBRL 파일 안에는 context id가 수백~수천 개 있습니다. id 문자열을 직접 매칭하는 게 아니라,  
각 context 안의 dimension 목록을 읽어 조건을 확인합니다.

```python
for ctx in root.findall("context"):
    dims = {}
    for m in ctx.iter("explicitMember"):
        dims[m.get("dimension")] = m.text   # 필터 목록 수집

    # 원하는 조건인지 확인
    if dims.get("ConsolidatedAndSeparateFinancialStatementsAxis") != "SeparateMember":
        continue
    if dims.get("DisaggregationOfInsuranceContractsAxis") != "InsuranceContractsIssuedMember":
        continue
    # → 이 context id가 우리가 원하는 것 → 이 id를 가진 값 태그를 찾아 읽음
```

회사를 미리 식별하는 게 아니라 **조건 자체를 순서대로 시도**합니다.  
각 컬럼마다 "이 조건 → 안 되면 이 조건 → 그것도 안 되면 이 조건" 순서로 패턴이 정의되어 있고,  
처음 매칭되는 패턴의 값을 사용합니다.

```python
def parse_insurance_components(root):
    # 패턴A 먼저 시도 (삼성생명 2-dim: DisaggregationAxis에 BEL/RA/CSM 직접)
    for ctx in ...:
        if dims["DisaggregationAxis"] == "EstimatesOfPresentValueOfFutureCashFlowsMember":
            bel = 이 context의 값
            return bel          # 찾으면 바로 반환

    # 패턴A 실패 → 패턴B 시도 (교보 등 3-dim: IssuedMember + ComponentsAxis)
    for ctx in ...:
        if dims["DisaggregationAxis"] == "InsuranceContractsIssuedMember"
        and dims["ComponentsAxis"] == "EstimatesOfPresentValueOfFutureCashFlowsMember":
            bel = 이 context의 값
            return bel

    # 패턴B도 실패 → 패턴C 시도 (현대해상 등 5~6-dim: 보험종류별 합산)
    for ctx in ...:
        if dims["InsuranceContractsAxis"] == "InsuranceContractsOtherThan..."
        and dims["ComponentsAxis"] == "EstimatesOfPresentValueOfFutureCashFlowsMember":
            bel += 이 context의 값     # 여러 개 합산
    return bel
```

이 구조 덕분에 새 회사가 추가되거나 구조가 바뀌어도 기존 패턴에 매칭되면 자동 처리되고,  
안 되면 새 패턴만 추가하면 됩니다.

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
| 9-11 | 총자산대비비중 | 계산 | Col5-7 / Col4 | ✅ |
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
| 25 | FVOCI 평가손익 | XBRL→원문XML | XBRL OCI 잔액; fallback: 자본 세부 테이블 | ⚠️ 소수 MISS/오차 |
| 26 | 보험계약 금융손익 | XBRL→원문XML→**FISIS** | XBRL 보험계약 OCI 잔액; fallback: FISIS SH151/SI147 F64 | ✅ 대부분 OK |
| 27 | 재보험계약 금융손익 | XBRL→원문XML→**FISIS** | XBRL 재보험계약 OCI 잔액; fallback: FISIS F65 | ✅ 12/12 |
| 28 | 현금흐름위험회피 | XBRL→원문XML→**FISIS**(생보) | XBRL; 생보사 fallback: FISIS F67 | ✅ 대부분 OK |
| 29 | 재평가잉여금 | XBRL→원문XML | XBRL 전기말 재평가잉여금; fallback: 자본 세부 테이블 | ⚠️ 교보·한화 오차 |
| 30 | 확정급여부채 재측정 | XBRL→원문XML→**FISIS** | XBRL; fallback: FISIS F69 | ✅ 대부분 OK |

---

### 손익계산서 (P/L)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 32 | 보험손익 | XBRL | `ifrs-full_InsuranceServiceResult` | ✅ 12/12 |
| 33 | 보험서비스수익 | XBRL | `dart_OperatingIncomeInsurance` | ✅ 12/12 |
| 34 | 보험서비스비용 | XBRL | `dart_OperatingExpenseInsurance` | ✅ 12/12 |
| 35 | 투자손익 | XBRL | `dart_InvestmentIncomeExpenses` | ✅ 12/12 |
| 36 | 보험금융손익 | XBRL | Insurance finance income − expense | ✅ 12/12 (반기 손보사=0) |
| 37 | 재보험금융손익 | XBRL | Reinsurance finance income − expense | ✅ 12/12 (반기 손보사=0) |
| 38 | 금융손익 | 계산 | Col35 − Col36 − Col37 + Col39 − Col40 | ✅ 12/12 (반기 손보사=0) |
| 39 | 재산관리비 | XBRL | `ifrs-full_SellingGeneralAndAdministrativeExpense` | ✅ 11/11 (반기 손보사=0) |
| 40 | 기타투자손익 | XBRL | 수수료+임대료+기타수익−기타비용 합산 (상각후원가·처분·평가·종속기업 태그 제외) | ⚠️ 일부 소오차 (반기 손보사=0) |
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
| 73 | BEL | XBRL → 원문XML | XBRL 보험계약부채 구성요소 BEL. 다중 패턴(DisaggregationAxis·ComponentsAxis·InsuranceContractsAxis) 순차 시도 | ✅ 12/12 (연간·분기) |
| 74 | RA | XBRL → 원문XML → **FISIS** | 동일; 분기는 doc_3q/doc_1q "분기말 순장부금액" 행 → FISIS SH156/SI152 A112+A122 | ✅ 12/12 (연간·분기) |
| 75 | CSM | XBRL → 원문XML → **FISIS** | 동일; XBRL/원문XML 실패 시 FISIS SH156/SI152 A113 | ✅ 12/12 |
| 76 | 신계약CSM | XBRL → 원문XML | XBRL CSM 변동 주석; fallback: 사업보고서 원문XML CSM 변동표 | ⚠️ 동양 MISS (연간만 추출, 분기 없음) |
| 77 | CSM 상각 | XBRL → 원문XML | 동일 | ⚠️ 동양 MISS (연간만) |
| 78 | CSM 조정 | XBRL → 원문XML | 동일 | ✅ 12/12 (연간) |
| 79 | 보험금융손익(CSM) | XBRL → 원문XML | 동일 | ⚠️ DB손보 MISS (연간만) |
| 80 | 전기말 CSM | XBRL → 원문XML | XBRL 전기말 CSM; fallback: 원문XML 변동표 `기초` 행 | ⚠️ 삼성화재·DB손보 MISS (연간만) |

> **RA·CSM FISIS fallback**: 분기/반기는 doc_3q/doc_1q XML 우선, 없으면 FISIS. 사업보고서는 XBRL/원문XML 우선.  
> **분기 BEL**: 추가 패턴(InsuranceContractsAxis 기반) 추가로 2509 기준 12/12 추출.  
> **Col76~80**: 사업보고서 전용. 3분기/반기 XBRL에 CSM변동 태그 없어 분기에서는 추출 불가.

---

### CSM 기간별 기대수익인식금액

| Col | 항목 | 소스 | 결과 |
|-----|------|------|------|
| 83~87 | 1년이하~10년초과 | **XBRL MaturityAxis** → 원문XML | XBRL `ContractualServiceMargin+MaturityAxis` 우선; fallback: doc.xml 테이블 패턴A/B | ⚠️ 연간 9/12 OK (미래에셋·KB손보 MISS), 분기 일부 MISS |
| 88 | 합계 | 계산 | = Col75 ✅ 11/11 |

> **XBRL MaturityAxis 지원 회사 (2025 연간)**: 삼성생명(4버킷)·한화생명·동양생명·미래에셋 일부.  
> **doc.xml 지원 회사**: 교보생명·현대해상("초과/이하" 7버킷)·메리츠화재(15버킷 상세)·DB손보·삼성화재.  
> **분기(2509)**: 삼성화재·현대해상·KB손보는 Q3 MaturityAxis 없어 MISS.

---

### K-ICS (지급여력비율)

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 92~94 | 잠정치 (비율·가용·요구) | 원문XML | 사업보고서 지급여력비율 테이블 | ⚠️ 일부 회사 MISS (산출중) |
| 95~97 | 최종치 (비율·가용·요구) | 원문XML → **FISIS** | 사업보고서는 1분기보고서 원문XML 전년도 열; 분기/반기는 FISIS SH021/SI021 | ✅ 2512 12/12, 분기 12/12 |

> - **사업보고서(연간)**: 다음 해 1분기보고서(`doc_1q_{year+1}.xml`)의 "전기" 열에서 확정치 추출.  
> - **분기/반기**: FISIS API 해당 분기 직접 조회. 잠정치(Col92-94)는 검증 제외.

---

### 해약환급금준비금

| Col | 항목 | 소스 | 태그 / 방법 | 결과 |
|-----|------|------|-----------|------|
| 99 | 해약환급금준비금 | XBRL → **FISIS** | `dart_SurrenderValueReserve`; 연간은 FISIS F46 우선, 분기는 XBRL 우선 | ✅ 12/12 |

> 연간 사업보고서는 FISIS F46이 XBRL보다 안정적이어서 우선 사용. 분기는 엑셀이 전년도 연말 확정치 기준이라 XBRL 값을 사용.

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

## 검증 방법

```bash
# 특정 기간 정확도 확인 + FAIL 셀 주황색 표시
python verify_2512.py --period 2512
python verify_2512.py --period 2509

# 출력 예시
# TOTAL: OK=964, FAIL=17, MISS=132 (total=1113)
# [색칠] 149셀 주황색 표시 완료 → DATA_작업_빈칸.xlsx
```

- **OK**: 절대차 ≤ 1 이거나 상대오차 < 0.1%
- **FAIL**: 값을 추출했으나 정답과 차이 발생 → 주황색 표시
- **MISS**: 값을 추출하지 못함 (None) → 주황색 표시
- CHECK 컬럼(Col31·45·46·47·54·64·72·82·89·90·92~94·159)은 수식으로 계산되므로 검증 제외

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
