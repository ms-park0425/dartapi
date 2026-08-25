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

| 기간 | 보고서 종류 | OK | FAIL | MISS | 비교대상 | **OK율** |
|------|-----------|---:|-----:|-----:|--------:|--------:|
| 2312 | 2023년 사업보고서 | 517 | 201 | 237 | 955 | **54%** |
| 2412 | 2024년 사업보고서 | 775 | 73 | 248 | 1,096 | **71%** |
| 2503 | 2025년 1분기 | 701 | 34 | 195 | 930 | **75%** |
| 2506 | 2025년 반기 | 718 | 52 | 186 | 956 | **75%** |
| 2509 | 2025년 3분기 | 812 | 30 | 110 | 952 | **85%** |
| 2512 | 2025년 사업보고서 | 1,043 | 19 | 51 | 1,113 | **94%** |
| 2603 | 2026년 1분기 | — | — | — | — | **미검증** |
| 2606 | 2026년 반기 | — | — | — | — | **미검증** |

> 2503~2512는 `parse_annual.py` + `ir_fill.py` 적용 후 최종 수치 (2026-08-25 기준). 2312·2412는 DART 기준.

> **2603·2606**: 정답 데이터 없어 정확도 미검증. 추출 완료: 2603 73개 컬럼·744개 값, 2606 69개 컬럼·670개 값 (12개 회사). 미래에셋 CSM 2배 버그(`DIVIDEND_CLASSIFICATION_PREFIXES` 미적용) 및 Col80 NaN(`parse_prior_csm` 5-dim 누락) 수정 후 재추출 완료(2026-08-24). `2026_동업사 공시비교_DATA_작업_260824.xlsx`에 1,258셀 채움.

**기간별 특이사항**

- **2312 (53%)**: IFRS17 도입 첫 해로 XBRL 태그 구조가 정착되지 않아 패턴 불일치 다수. 전기말(2022년말) XBRL이 IFRS17 재작성 기준이라 이전기 비교수치 불일치.
- **2412 (70%)**: XBRL에 BEL/RA/CSM 구성요소 태그가 없는 회사 다수(2025년부터 확대 공시). FISIS로 보완. OCI세부(Col26-30)도 FISIS 추가.
- **2503 (85%)**: 1분기 보고서. K-ICS·RA·CSM·해약환급금·OCI세부는 FISIS로 채움. 삼성생명·메리츠화재 Col113~117 매핑 추가(2026-08-25)로 비교대상 확대.
- **2506 (90%)**: 반기 보고서. 손보사 Col36~40 반기 기재 방식 차이 처리. 삼성생명·메리츠화재 Col113~117 매핑 추가.
- **2509 (90%)**: 3분기 보고서. doc_3q XML에서 메리츠·KB손보 RA 직접 추출. 한화·동양·미래에셋 분기 BEL 추가 패턴. XBRL MaturityAxis CSM기간별 적용.
- **2512 (95%)**: 사업보고서 전용 항목 포함 최고 정확도. XBRL MaturityAxis에서 삼성생명·한화생명·동양생명·미래에셋 CSM기간별 직접 추출. 삼성생명·메리츠화재·한화·동양·삼성화재 Col108~117 IR 매핑 추가(2026-08-24~25). 잔여 MISS 51개: CSM기간별(미래에셋·KB손보), 계리가정변경(일부), Col23 자산운용률(3개사).
- **2603 (미검증)**: 2026년 1분기. DART 745셀 + IR fill +65셀 = 810셀 (12개 회사). 예실차 Col111-117 7~8개사 커버.
- **2606 (미검증)**: 2026년 반기. DART 671셀 + IR fill +75셀 = 746셀 (12개 회사). 2026년 신규 XBRL 구조(eHYA·eHYQ 분리, 전기말CSM 5-dim) 대응.

> **재검증 (2026-08-24)**: `parse_csm_movement` DIVIDEND 필터 및 `parse_prior_csm` 5-dim 패턴 수정 후 2503·2506·2509·2512 전체 재추출. 2503 수치 동일. 2506·2509·2512는 OK/FAIL/MISS 배분이 1~3셀 내 소폭 이동했으나 OK율은 변동 없음(70%·80%·88%).

> **2412 BEL/CSM MISS 원인**: 삼성생명 기준 2024년 XBRL 태그 수 245개(BEL/RA/CSM 분리 없음) → 2025년 988개(구성요소 분리 태그 신설). 2025년부터 대부분 회사가 XBRL 공시 수준을 크게 확대해 해결됨.

> **2026년 XBRL 구조 변경 (3가지)**  
> ① BS context `eHYA`(반기누계)·`eHYQ`(분기단독) 분리 → B/S 항목은 `eHYA` 우선 사용.  
> ② 전기말 CSM(Col80) 3-dim 구조 변경 (`SeparateMember + IssuedMember + ContractualServiceMarginMember`, tag=`ContractualServiceMargin`) → `parse_prior_csm` 패턴C 추가.  
> ③ 미래에셋 2026 XBRL: `TypesOfContractsAxis` 안에 **보험종류별**과 **배당유무별** 두 분류가 공존 → `DIVIDEND_CLASSIFICATION_PREFIXES` 필터를 `parse_csm_movement`(`target_ctxs`·`all_sep_ctxs`)에도 적용, `parse_prior_csm`에 5-dim 패턴A2 추가(tag=`InsuranceContractsThatAreLiabilities`). 이 수정 전에는 Col79(이자부리)·Col78(CSM조정)·Col80(전기말CSM)이 2배 또는 NaN이었다.

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

---

## IR 팩트시트 연계 (`ir_*.py`)

예실차·CSM 변동·손해율 같은 항목은 XBRL에 아예 없다. 애초에 DATA 시트의 그 칸들이
각 사 IR 팩트시트에서 온 값이기 때문이다. 미래에셋생명 FY2025 팩트시트 값에 1000을
곱하면 정답 파일과 소수점까지 일치하는 것으로 확인했다.

| 파일 | 역할 |
|---|---|
| `ir_parse.py` | 팩트시트(xlsx) 추출 엔진 + 회사별 매핑 `SPECS` |
| `ir_discover.py` | 정답 시트를 열쇠로 매핑을 역탐색 |
| `ir_fill.py` | 기간 지정 → `data/data_sheet_{기간}_ir.csv` 로 병합 |
| `verify_ir.py` | 추출값 ↔ 정답 대조 |
| `ir_store.py` | 받아온 IR 자료를 `data/{회사}/IR/{기간}/` 로 정리 |
| `ir_pdf.py` | PDF 파싱 — 회사별 파서 + **정기경영공시 공용 파서** |
| `ir_identity.py` | DATA 시트 안 컬럼 간 항등식을 찾아 빈칸을 계산으로 메움 |
| `sweep_miss.py` | 남은 빈칸을 팩트시트 전체에 대해 넓게 역탐색 |
| `sweep_pdf.py` | 남은 빈칸을 PDF 전체에 대해 역탐색 |
| `score_miss.py` | 작업 파일 ↔ 정답 시트 채움률 채점 |
| `recalc.py` | LibreOffice 로 수식을 계산시킨 사본 생성(채점 전처리) |

### 자료 보관 구조

IR 자료는 XBRL 캐시와 같은 자리에 회사별로 둔다.

```
data/{회사명}/IR/{기간}/{회사명}_{기간}_{종류}.{확장자}
   예) data/미래에셋생명/IR/2606/미래에셋생명_2606_팩트시트.xlsx
       data/교보생명/IR/2512/교보생명_2512_경영공시.pdf
```

종류는 팩트시트 / 팩트북 / 발표자료 / 경영공시.
브라우저로 받으면 파일명이 제각각이고 현대해상·메리츠는 아예 해시(`d5df3fdc….xlsx`)라
그대로 두면 나중에 알아볼 수 없다. `ir_store.py` 가 **시트 구성을 보고 회사를 판별해**
위 이름으로 바꿔 넣는다.

```bash
python ir_store.py --dry                    # 어디로 갈지 확인
python ir_store.py --period 2606            # 다운로드 폴더에서 정리 (기간 지정)
python ir_store.py --from ~/Downloads       # 소스 폴더 지정
```

기간은 `--period` 가 최우선이고, 없으면 파일명에서 추정한다. 파일명이 해시면
`--period` 를 반드시 줘야 한다. (본문에서 기간을 추정해봤으나 시트 안의 숫자를
헤더로 오인해 엉뚱한 값이 나오는 일이 잦아 뺐다.)

### 사용법

```bash
python ir_fill.py --period 2606 --dry      # 확인만
python ir_fill.py --period 2606            # data/data_sheet_2606_ir.csv 저장
python ir_fill.py --period 2606 --excel    # DATA_작업_빈칸.xlsx 에도 직접 기록

# 3) 정답이 있는 과거 기간으로 매핑 검증
python verify_ir.py 삼성생명 "$IR_DIR/SLI FY25Factsheet_KOR.xlsx" --periods 2503,2506,2509,2512

# 4) 새 회사·새 컬럼 매핑은 손으로 찾지 말고 역탐색부터
python ir_discover.py 삼성생명 "$IR_DIR/SLI FY25Factsheet_KOR.xlsx" 2512 --only 107,108,109
```

팩트시트는 `data/{회사}/IR/{기간}/` 에서 자동으로 찾는다. 해당 기간 파일이 없으면
**같은 연도의 더 나중 분기 파일**을 쓴다 — 분기 팩트시트에는 그 연도의 앞선 분기가
함께 들어 있기 때문이다(2Q 팩트시트 하나로 1분기·반기가 같이 채워진다).
대신 읽은 경우 실행 로그에 표시된다.

`--excel` 을 주면 `DATA_작업_빈칸.xlsx` 의 DATA 시트에 직접 써넣는다.
`parse_annual.py` 가 1차로 채운 뒤 남은 빈칸을 메우는 용도라 **빈칸만** 채우고,
IR에서 온 값은 **연한 파랑**으로 칠해 출처가 눈에 보이게 한다
(FAIL/MISS 를 표시하는 주황색과 구분된다).

`ir_fill.py` 는 **빈칸만 채운다.** 기존 값은 건드리지 않는다.
IR 값이 기존 값과 다르면 그건 정보다 — 실제로 이 대조로 아래 2배 버그를 잡았다.

### 검증 결과 (2025년 4개 기간 × 7개사, 292셀)

| 회사 | OK | FAIL | 채우는 컬럼 |
|---|---:|---:|---|
| 미래에셋생명 | 89 | 7 | 22·25~30·75~80·99·107~109·111~117 |
| 삼성화재 | 42 | 0 | 25·73~80·92~97·99·112·115 |
| 한화생명 | 40 | 0 | 21·75~80·111·112·115 |
| 삼성생명 | 38 | 2 | 75~80·107~109·111 |
| 동양생명 | 27 | 0 | 75~80·92~94 |
| 메리츠화재 | 24 | 0 | 75~80 |
| 현대해상 | 23 | 0 | 13·18·21~23·55·63·65·71·99·111~117 |
| **합계** | **283** | **9** → **96.9%** |

FAIL 7건은 팩트시트가 **과거 분기를 재작성(restate)** 한 탓이다.
각 기간은 그 기간에 공시된 팩트시트를 쓰면 해소된다.
(Col29 재평가잉여금은 값 자체가 100만원 안팎이라 상대오차가 크게 잡힌다)

### FISIS 대체

미래에셋생명 팩트시트의 **`SAP-BS`(보험회계기준 재무상태표)** 시트에 FISIS 에서 받던
항목이 그대로 있다 — 운용자산(Col22)·OCI 세부(Col26~30)·해약환급금준비금(Col99).
이 시트는 **백만원 단위**라 다른 시트(십억원)와 달라 `Item(..., unit=1)` 로 지정한다.

삼성화재는 `Capital Ratio` 시트에서 K-ICS 가용·요구자본(Col96·97),
`Liability structure` 시트에서 해약환급금준비금(Col99)이 나온다.

다만 회사마다 편차가 크다. 한화생명 팩트시트에는 운용자산·OCI 세부가 없다.
**FISIS 를 완전히 걷어내는 것보다 이중 출처로 두고 서로 검산하는 쪽이 낫다.**

### IR 보완 효과 (2025년 4개 기간, 정답 대조)

`parse_annual.py` 로 1차를 채운 뒤 `ir_fill.py` 를 돌린 최종 수치. 비교대상이 기간마다 다른 이유는 IR 매핑이 늘어나면서 정답이 있는 셀 수도 함께 늘기 때문이다.

| 기간 | DART만 OK율 | **IR 포함 OK율** | OK | FAIL | MISS | 비교대상 |
|---|---:|---:|---:|---:|---:|---:|
| 2503 | 67% | **75%** | 701 | 34 | 195 | 930 |
| 2506 | 70% | **75%** | 718 | 52 | 186 | 956 |
| 2509 | 80% | **85%** | 812 | 30 | 110 | 952 |
| 2512 | 88% | **94%** | 1,043 | 19 | 51 | 1,113 |

(2026-08-25 기준. 삼성생명·메리츠화재·동양생명 Col113~117 매핑 추가 포함)

### 잔여 MISS 분석 (2512 기준, 51개)

| 성격 | 셀 | 상황 |
|---|---:|---|
| CSM 기간별(83~87) | 10 | 미래에셋·KB손보 — XBRL/팩트시트 모두 없음 |
| 예실차 세부(113~117) | 15 | 삼성생명·한화·메리츠 일부 — 시트 구조 불일치 또는 항목 없음 |
| 계리가정변경(146~157) | 12 | 동양·삼성화재 — 연간 경영공시 PDF만 있음 |
| 자산운용률(Col23) | 3 | 삼성생명·삼성화재·KB손보 — doc.xml 테이블 없음 |
| 기타 분산 | 11 | Col21·29·59·60·62·70·79·80 등 |

팩트시트를 내는 8개사(현대해상 포함)는 사실상 짜낼 만큼 짜냈다.
남은 최대 덩어리는 **미매핑 3개사(KB손보·KB라이프·신한라이프)** 이며 PDF 파싱으로 접근 가능.

> **주의**: 매핑을 늘려도 MISS 가 그만큼 줄지는 않는다. 새로 매핑한 컬럼 상당수가
> 이미 DART·FISIS 로 채워져 있던 칸이기 때문이다(현대해상 재무상태표, 삼성화재 K-ICS 등).
> 그 매핑들은 빈칸을 메우기보다 **기존 값을 교차검증하는** 데 값어치가 있다.

### 미매핑 4개사 — 왜 못 붙였나

자료를 다 받아놓고 `ir_discover.py` 로 전수 대조한 결과다. **자료가 없어서가 아니라
기준이 달라서** 쓸 수 없는 경우가 대부분이다.

| 회사 | 가진 자료 | 판정 |
|---|---|---|
| KB손해보험 · KB라이프 | KB금융 팩트북 `I_*`/`L_*` 시트 | **기준 불일치.** 총자산이 정답과 0.05~0.16% 어긋난다(연결 vs 별도). 연간 열은 `Dec. 25(E)` 로 **잠정치**라 1.08% 차이 |
| 신한라이프 | 신한지주 팩트북 `Shinhan Life` 시트 | K-ICS·총자산·운용자산·요약손익까지만. **CSM·예실차가 없다** |
| 교보생명 | 정기경영공시 PDF | **매핑 완료.** `ir_pdf.parse_disclosure()` 가 19개 컬럼을 뽑는다(아래 참조). Col83~87 버킷은 여전히 정의가 달라 쓰지 않는다 |
| DB손해보험 | 분기 팩트시트(xlsx) | **매핑 완료.** 사이트가 복구돼 2503~2606 6개 분기를 모두 받았다. 36셀 전부 정확히 일치 |

KB 팩트북은 `Mar. 25` 형태 헤더를 읽도록 엔진을 고쳐 접근은 되게 해뒀다.
값을 쓰지 않을 뿐, 나중에 기준이 맞는 표가 추가되면 바로 붙일 수 있다.

### 검증 실패 9건의 성격

`verify_ir.py` 기준 292셀 중 283셀 일치(96.9%). 실패 9건은 세 종류다.

**① 팩트시트 재작성 (5건)** — 미래에셋 2503 예실차 계열 3건, 삼성생명 2509 손해율 2건.
FY2025 Q4 팩트시트로 과거 분기를 읽은 결과다. 같은 표의 예상보험금은 정확히 일치하는데
실제보험금만 어긋나는 것이 재작성의 증거다. `data/{회사}/IR/2503/` 에 그 분기 팩트시트를
넣으면 자동으로 그쪽을 먼저 읽는다.

**② 값이 너무 작아 상대오차만 큰 것 (2건)** — 미래에셋 재평가잉여금(Col29).
값 자체가 100만원 안팎이라 3.5% / 27.9% 로 잡히지만 절대차는 미미하다.

**③ 정답 파일 쪽 오류로 보이는 것 (2건)** — 미래에셋 **2506** 위험보험료(Col108)와
사고보험금(Col109)이 **서로 뒤바뀌어 있다.**

| | 팩트시트 | 정답 파일 |
|---|---:|---:|
| Col108 위험보험료 | 246,452.917 | 220,173.300 |
| Col109 사고보험금 | 220,173.300 | 246,452.917 |

두 값이 정확히 교차한다. 2503·2509·2512 는 소수점까지 일치하므로 코드 문제가 아니고,
2506 정답만 손해율이 100%를 넘게 되어 있다. **확인이 필요하다.**

### 값을 고를 때의 규칙 (정답 대조로 확정)

- **단위가 회사마다 다르다.** 십억원(미래에셋·한화·동양·삼성생명)=×1000,
  억원(삼성화재·메리츠·교보)=×100, 백만원(현대해상)=×1
- **흐름 항목은 연간 누계.** 예실차·신계약CSM·상각·조정·이자부리·위험보험료·사고보험금.
  FY/YTD 열이 있으면 그대로, 없으면 Q1~해당 분기 합산
- **잔액 항목은 시점값.** 기말CSM(Col75)은 그 분기말. 연간이면 4분기말
- **Col80(전기CSM)은 직전 연도 팩트시트의 기말CSM.** 당해 팩트시트에 전년 '연간' 열이
  없고 동기 분기만 실리는 경우가 있다(삼성화재 26.2Q)
- **Col77(CSM상각)은 부호 반전.** 팩트시트는 음수, DATA는 양수
- **Col95(K-ICS 비율)는 분수로 저장.** 166%면 1.66

### IR 자료로 못 채우는 항목

CSM 기간별 기대수익인식(Col83~88), 가정변경 세부(Col146~157), OCI 세부(Col24~30),
자본변동표(Col55~71)는 IR 자료에 항목 자체가 없다. 사업보고서 주석·경영공시·FISIS 가
정공법이다.

### 회사별 IR 자료실

| 회사 | 경로 | 형태 |
|---|---|---|
| 삼성생명 | samsunglife.com/individual/display/invest/PDK-IRIVI015220M | PDF + XLS |
| 삼성화재 | samsungfire.com/vh/page/VH.HPMK0201.do | PDF + XLSX |
| 한화생명 | company.hanwhalife.com/ko/investment/investor/earnings-release | PDF + 팩트시트 |
| 미래에셋생명 | life.miraeasset.com/micro/company/PC-HO-060401-000000.do | 연도 select 후 다운로드 |
| 동양생명 | myangel.co.kr/Company/Ir/CoIrData | PDF + XLS |
| 현대해상 | hi.co.kr → IR > IR자료실 > 실적발표자료 | onclick 에 `doBizFileDownload('/data/…')` 노출 |
| 메리츠금융 | meritzgroup.com/web/ko/ir/ir1.do | href 에 `/commfiles/…` 직접 노출 |
| KB금융 | kbfg.com/kor/ir/report/factbook/list.jsp | Fact Book. 계열사 상세는 IR PT |
| 신한지주 | shinhangroup.com/kr/ir/finance/factBook | XLS. 'Shinhan Life' 시트는 요약만 |
| 교보생명 | kyobo.com/dgt/web/notice-management/fixed-term/last-fiveYear | 정기경영공시 PDF |
| DB손해보험 | idbins.com/pc/bizxpress/cmy/inv/ir/FWCOMV1705_{YYMMDD}.shtm | PDF |

대부분 JavaScript 렌더링이라 정적 HTML만으로는 링크를 못 얻는다. 브라우저로 열어야 한다.
분기 실적발표는 분기 종료 후 6주 안팎(1분기 5월 중순, 반기 8월 중순, 3분기 11월 중순,
연간 다음 해 2월 중순)에 나온다.

---

## 2차 확장 — 페어합·느슨한 허용오차·PDF 전수 훑기

"붙일 수 있는 게 그게 끝이냐"는 물음에 대한 답이다. 세 갈래로 더 팠다.

### ① 두 행의 합 / 느슨한 허용오차 재탐색 (`sweep_miss.py`)

`ir_discover.discover_pairs()` 로 **같은 열 위아래 두 행의 합**을, `--loose` 로
**반올림 자릿수가 다른 근사치**를 각각 훑었다. 우연의 일치가 압도적이라
`_line_matches()` 로 "그 줄의 글자가 컬럼 이름과 실제로 겹칠 때"만 후보로 남겼다.

건진 것과 버린 것:

| 회사 | 새로 붙인 컬럼 | 버린 후보와 이유 |
|---|---|---|
| 삼성생명 | Col73(기말 BEL)·Col74(기말 RA) — `Ⅰ-3` 시트 | Col49(투자손익) 별도/연결 기준 차이로 2~3배 어긋남 |
| 한화생명 | Col21·44·48·74·108·**109**·**107** — `(별도) 요약 손익/재무상태표`, `효율지표` | Col109(사고보험금) 효율지표 추가 확인(일치). Col73(BEL) 16% 차이(발생사고요소 제외), Col22(운용자산) 특별계정 포함 여부는 여전히 제외 |
| 동양생명 | Col13·48·111·112·115·**108**·**109**·**107** — `FS`·`Productivity` 시트 | Col18·21·24·44·49·50은 `FS`가 연결 기준이라 DATA(별도)와 3% 차이 → 제외 |
| 미래에셋·삼성화재·현대해상·메리츠 | 없음 | 페어합 후보가 전부 우연의 일치였다 |

> 결론: 페어합 자체로 새로 붙은 컬럼은 없었다. 실효는 **느슨한 허용오차 + 의미 필터**
> 조합에서 나왔고, 그중 절반은 기준(연결/별도, 잔여보장/발생사고)이 달라 버렸다.

### ② DB손해보험 — 분기 팩트시트 6개 확보

`https://www.idbins.com/pcweb/bizxpress/cmy/inv/ir/__etc/YYYY.MM_FactSheet_DB Insurance_Kor.xlsx`
로 2503·2506·2509·2512·2603·2606 을 모두 받아 `data/DB손해보험/IR/{기간}/` 에 정리했다.

DB손보 팩트시트의 특징은 **movement 를 분기 단독으로 싣는다**는 점이다.
`BEL,CSM변동` 시트의 `'25.1Q | 2Q | 3Q | 4Q` 열은 누계가 아니라 각 분기값이라,
연간누계 컬럼(Col76~79)은 `CUM` 모드가 Q1..Qn 을 더해 만든다. `CSM(기시)` 는
1분기 열에서만 읽어야 전기말 잔액이 되므로 `OPEN` 모드를 새로 넣었다.

| 시트 | 행 | → | 모드 |
|---|---|---|---|
| `BEL,CSM변동` | CSM(기말) | Col75 | POINT |
| | 신계약 유입 | Col76 | CUM(분기합) |
| | 상각 | Col77 | CUM, 부호반전 |
| | CSM 조정 | Col78 | CUM |
| | 이자부리 | Col79 | CUM |
| | CSM(기시) | Col80 | **OPEN**(1Q 열) |
| `보험손익` | 보험금 예실차(발생사고요소조정 포함) | Col112 | CUM |
| | 사업비 예실차 | Col115 | CUM |
| (파생) | Col112 + Col115 | Col111 | |

**2503~2512 4개 기간 × 9컬럼 = 36셀 전부 정답과 일치(오차 0.002% 이내).**

Col73(BEL)·Col74(RA)는 붙이지 않았다. DB손보 팩트시트의 BEL 은 잔여보장부채만
담아 DATA 기준(순부채)과 30% 넘게 다르다.

### ③ 정기경영공시 PDF 공용 파서 (`ir_pdf.parse_disclosure()`)

이게 이번 확장에서 제일 큰 수확이다. **금감원 정기경영공시는 모든 보험사가
같은 목차·같은 표 제목을 쓴다.** 회사별 매핑 없이 표 제목만 보고 읽으면 된다.

교보생명 2512 경영공시로 검증한 결과 — **19/21 일치**:

| 근거 표 | 컬럼 |
|---|---|
| `회계모형별, 포트폴리오별 보험부채 현황` 합계행 | Col75 (일반모형 CSM + VFA CSM) |
| `일반회계와 감독회계의 차이` | Col48·Col49·Col50 |
| `보험부채 변동내역 > 계리적 가정` | **Col146~150(BEL+RA기준)**, **Col153~157(CSM기준)** |
| `최적가정 > 보험금 예실차비율` | Col101·Col102·Col103 |
| `해약환급금준비금 등의 적립` | Col18·Col71·Col99 |
| 보험수익 주석 `가. CSM 상각액` | Col77 |

계리적 가정 표는 `[최선추정부채 | 위험조정 | 보험계약마진]` 3열이고,
DATA 의 `BEL+RA기준`은 **앞 두 열의 합**, `CSM기준`은 마지막 열이다.
값이 0 인 열은 PDF 에서 아예 빠져 숫자가 2개만 잡히므로 `sum(v[:-1]) / v[-1]` 로 읽는다.

> **이전 판정 정정**: README 앞부분에서 "가정변경(146~157) 22셀 — IR 자료에 없음"
> 이라고 적었는데, **틀렸다.** 정기경영공시에 그대로 있다. 팩트시트에 없을 뿐이다.

쓰지 않기로 한 것:

- **지급여력비율 총괄** — 경영공시 값은 잠정치·억원 반올림이라 DATA(FISIS 최종치)와
  최대 1%p 어긋난다(교보 2512 요구자본 86,689억 vs 정답 85,825억).
- **보험계약마진 상각 만기표(Col83~87)** — 이 표는 *순상각액*(이자부리 차감) 기준이라
  버킷을 어떻게 묶어도 DATA 의 *기대수익인식금액*과 맞지 않는다.
  교보 2512: 표 1년 3,758억 vs DATA 1년이하 5,708억. 계(65,110억)만 Col88 과 일치한다.

`_nums()` 도 고쳤다. `'5,200 8 -5,208'` 처럼 부호가 붙은 뒤 토큰까지 앞 숫자에
붙여버리는 버그가 있었다 — 뒤 토큰이 `-` 로 시작하면 붙이지 않는다.

### ④ 시트 내부 항등식으로 메우기 (`ir_identity.py`)

IR 자료를 아무리 뒤져도 안 나오는 칸이 **같은 행의 다른 칸으로 계산되는** 경우가 많다.
정답 시트 84행 전체에서 예외 없이 성립하는 관계만 채택했다(상대오차 0.2% 이내,
20행 이상 검증, 기여도 2% 미만인 항은 잡음으로 배제).

찾은 항등식 30개 중 핵심:

| 종류 | 관계 |
|---|---|
| 같은 값의 중복 컬럼 | Col17=Col63, Col18=Col71, Col32=Col48, Col35=Col49, Col42=Col50, Col44=Col68, **Col75=Col88** |
| 부분합 | Col14=Col15−Col16, Col41=Col32+Col35, Col51=Col41+Col42 |
| 예실차 블록 | Col111=Col112+Col115, Col113=Col112+Col114, Col116=Col115+Col117 |
| 비율 | Col21=Col22÷Col23, Col104=Col102÷Col101 |

**Col75 = Col88 이 69행에서 예외 없이 성립**한다 — CSM 기간별 기대수익인식금액의
`합계` 는 CSM 잔액 그 자체다. 위에서 "CSM 기간별 157셀은 IR 로 불가"라고 했는데,
적어도 `합계`(Col88)는 CSM 잔액만 있으면 공짜로 채워진다.

전 기간 작업 파일에 적용해 **657칸**을 계산으로 메웠고, 그중 정답이 존재하는
**349칸을 대조해 전부 일치(오차 0건)** 했다. 나머지 308칸은 2603·2606 등
정답 자체가 아직 없는 기간이다.

```bash
python ir_identity.py                     # 항등식 목록만 출력
python -c "import ir_identity as I; I.fill_blanks('DATA_작업_빈칸.xlsx', dry=False)"
```

### 채점할 때 주의 — 수식 칸

DATA 작업 파일에는 사용자가 손으로 넣은 수식(`=-EP28`, `=19487+500373+68136`)이 섞여 있다.
openpyxl 로 저장하면 수식의 **계산값이 사라지므로**, `data_only=True` 로 읽으면
멀쩡한 칸이 전부 빈칸으로 보이고 `data_only=False` 로 읽으면 빈 칸을 참조하는
수식까지 "채워짐"으로 세어 반대로 과대평가된다.

정직하게 세려면 한 번 계산시켜야 한다.

```bash
python recalc.py DATA_작업_빈칸.xlsx /tmp/after_calc.xlsx   # LibreOffice headless
python score_miss.py /tmp/after_calc.xlsx --periods 2503,2506,2509,2512 --detail
```

### 현재 매핑 상태 (2503~2512, `verify_ir.py`)

| 회사 | OK | FAIL | 추가된 컬럼 (2026-08-24~25) |
|---|---:|---:|---|
| 미래에셋생명 | 89 | 7 | 팩트시트 재작성·정답파일 오류(기존 분석 그대로) |
| 한화생명 | 60 | 0 | Col109·Col107 추가 |
| 삼성생명 | 58 | 2 | Col113·114·116·117·112·115 추가(Ⅴ-4 손익계산서). 실패 2건은 2509 손해율 재작성 |
| 동양생명 | 53 | 2 | Col108·109·107·113·114·116·117 추가(FS+IS K-IFRS) |
| 삼성화재 | 50 | 0 | Col108·111·113·114·116·117 추가 |
| 메리츠화재 | 36 | 0 | Col113·114·116·117·112·115 추가(Insurance_PL) |
| DB손해보험 | 36 | 0 | |
| 현대해상 | 23 | 0 | |
| **합계** | **405** | **11** | 97.3% |

교보생명은 경영공시 PDF 로 별도 검증(19/21). 남은 미매핑은
**KB손해보험 · KB라이프 · 신한라이프 3개사**뿐이다.

### 다음에 할 일

1. **교보 2503·2506·2509 정기경영공시 PDF 수집** — 파서는 이미 있으므로 파일만 넣으면 된다.
   `data/교보생명/IR/{기간}/교보생명_{기간}_경영공시.pdf`
2. **다른 회사 경영공시도 같은 파서로** — 팩트시트에 없는 가정변경(146~157)·
   예실차비율(101~103)이 전 회사 공통으로 들어 있다. KB손보·KB라이프·신한라이프도
   경영공시를 받으면 매핑 없이 바로 읽힌다.
3. **분기 팩트시트 확보** — 현대해상 `(6) BS`, 삼성화재 `Capital Ratio` 처럼 한 기간만
   싣는 시트는 그 분기 파일이 있어야 Q1~Q3 을 채울 수 있다. 지금은 2512·2606 파일만 있어
   해당 컬럼이 연말값만 나온다.

---

## ⚠️ 작업 파일 주의 — `DATA_작업_빈칸.xlsx` 는 쓰지 말 것

2026-08-24 확인. 저장소에 있는 엑셀 3개는 상태가 서로 다르다.

| 파일 | 행 | 수식 | 다른 회사 행을 가리키는 수식 | 판정 |
|---|---:|---:|---:|---|
| `2026_동업사 공시비교_DATA_작업_260820.xlsx` | 95 | 2,318 | **0** | ✅ **이게 진짜 작업 파일** |
| `DATA_작업_빈칸_2512backup.xlsx` | 84 | 1,014 | 0 | 옛 레이아웃 백업 |
| `DATA_작업_빈칸.xlsx` | 95 | 1,809 | **632** | ❌ **망가짐** |

`DATA_작업_빈칸.xlsx` 는 84행 레이아웃에서 쓰던 수식이 95행 레이아웃으로 옮겨오면서
**참조 행이 4칸씩 어긋난 채로 남았다.** 예를 들어 교보 2512 는 32행인데
`Col153` 수식이 `=-EP28` 로 **28행(다른 회사)** 을 가리킨다. 632칸이 이런 상태다.

이 파일을 기준으로 채움률을 세면 수식이 엉뚱한 행에서 0 을 계산해 와서
"채워졌는데 값이 틀린" 칸이 1,000개 넘게 잡힌다. 파이프라인 성과가 아니라 파일 결함이다.

`ir_fill.py` 의 대상 엑셀은 하드코딩을 없애고 바꿀 수 있게 했다.

```bash
python ir_fill.py --period 2606 --excel --excel-path "2026_동업사 공시비교_DATA_작업_260820.xlsx"
# 또는  set WORK_XLSX=2026_동업사 공시비교_DATA_작업_260820.xlsx
```

### 진짜 작업 파일의 현재 상태

`2026_동업사 공시비교_DATA_작업_260820.xlsx` 를 LibreOffice 로 계산시켜 정답과 대조한 결과
**2212~2512 는 빈칸 0 · 불일치 0** 이다. 이미 완성돼 있다.

남은 일은 정답이 아직 없는 **2603 · 2606** 뿐이다.

| 기간 | 값 있는 칸 | 빈칸 |
|---|---:|---:|
| 2603 | 768 | 1,164 |
| 2606 | 694 | 1,238 |

여기에 팩트시트·경영공시·항등식을 돌려 **115칸**(2603 57 · 2606 58)을 새로 채웠다.
과거 기간 대조는 그대로 불일치 0 을 유지한다.

| 컬럼 | 칸 | 컬럼 | 칸 |
|---|---:|---|---:|
| Col111 예실차 | 14 | Col113·114·116·117 | 각 4 |
| Col112 보험금예실차 | 12 | Col77 CSM상각 | 4 |
| Col115 사업비예실차 | 12 | Col92·93·94 K-ICS | 각 3 |
| Col79·80·88·108·23 | 각 6 | Col107·109 손해율 | 각 4 |

> `ir_identity.fill_blanks()` 에 `periods=` 를 주면 그 기간만 손댄다. 이미 완성된 과거
> 기간까지 채우면 검산(CHECK) 컬럼이 한쪽만 채워져 0 이 아니게 되는 부작용이 있다.
> 검산 전용 컬럼(31·45·46·47·64·72·89·90·91)은 아예 채우지 않도록 막아뒀다.

## 경영공시 수집 — 실제로 받을 수 있는 것

교보생명 홈페이지(`kyobo.com/dgt/web/notice-management/fixed-term/last-fiveYear`)를
확인한 결과:

- **연간 정기경영공시는 PDF 로 제공**된다(직전 5개년, 2021~2025).
- **분기 경영공시는 PDF 가 아니다.** 당해년도 분기는 HTML 표로만 보여주고
  다운로드는 "준비 중입니다" 로 막혀 있으며, 지난해 분기 자료는 내려가 있다.

즉 **2503·2506·2509 같은 과거 분기 경영공시 PDF 는 구할 수 없다.**
경영공시로 얻을 수 있는 것은 12월(연간) 기간뿐이다.

그런데 2212·2312·2412 는 정답 시트가 이미 다 채워져 있어 새로 얻을 게 없다.
`ir_pdf.parse_disclosure()` 의 실전 가치는 **앞으로 매년 12월 기간**에 있다.

2024년 경영공시로 파서를 한 번 더 검증했다(`data/교보생명/IR/2412/`) — Col18·48·49·50·
71·75·77·101·102 전부 일치. 그 과정에서 두 가지를 고쳤다.

- `_row_after()` 가 항목번호까지 숫자로 읽던 문제 — `(3) 당기손익에…` 의 `3` 이 값으로 잡혔다.
- **Col77 은 `가. CSM 상각액` 이 아니라 `(3) 당기손익에 인식한 보험계약마진 금액`** 이다.
  취소계약분이 빠져 2024년에 153억 어긋났다. 2025년은 취소계약이 0 이라 우연히 같았다.

### 2603·2606 은 같은 매핑이 그대로 먹힌다

"25년 기준으로 값을 가져올 수 있으면 2603·2606 도 똑같이 되는 것 아니냐" — 맞다.
회사별 `SPECS` 는 기간에 종속되지 않으므로 팩트시트만 있으면 그대로 돌아간다.
실제 추출 컬럼 수를 세어보면:

| 회사 | 2512 | 2603 | 2606 | 비고 |
|---|---:|---:|---:|---|
| 미래에셋 | 24 | 24 | 24 | |
| 삼성화재 | 18 | **8** | 18 | 2603 팩트시트가 없어 2606 파일로 대신 읽음 |
| 현대해상 | 17 | **9** | 17 | 〃 |
| 한화 | 14 | 14 | 14 | |
| 동양 | 14 | 14 | 14 | |
| 삼성생명 | 12 | 12 | 12 | |
| DB손보 | 9 | 9 | 9 | |
| 메리츠 | 6 | 6 | 6 | |
| 교보 | 21 | **0** | **0** | 경영공시는 연간만 나온다(분기 PDF 없음) |

2606 은 8개사 전부 2512 와 **동일한 컬럼 수**를 뽑는다. 2603 만 두 회사가 모자란데,
`(6) BS`·`Capital Ratio` 처럼 **한 기간만 싣는 시트** 때문이고 2603 팩트시트를 받으면 해결된다.
다만 실익은 3칸(삼성화재 Col92·93·94)뿐이라 급하지 않다.

**IR 로 뽑히는 컬럼 대부분은 이미 DART·FISIS 가 채워둔 자리다.** 그래서 실제로 새로
들어가는 칸은 기간당 50여 개다. 이건 매핑이 부족해서가 아니라 겹치기 때문이다.

### 고친 버그 — 회사명이 수식인 행을 통째로 건너뛰고 있었다

작업 파일에서 같은 회사의 뒷 기간 행은 회사명을 `=C108` 같은 **위 행 참조 수식**으로
넣어둔다. openpyxl 은 수식 문자열을 그대로 돌려주므로 `str(...)== '_DB손보'` 비교가
실패해 **그 행 전체가 채워지지 않았다.** KB손보·DB손보의 2503~2512, 삼성화재 2506,
현대해상·메리츠 2503·2506 이 여기 해당한다.

`ir_fill.resolve_short()` 를 만들어 참조를 따라가 실제 이름을 읽도록 고쳤고,
`ir_identity` · `score_miss` 도 같은 함수를 쓴다.

### 최종 결과 (`2026_동업사 공시비교_DATA_작업_260820.xlsx` 기준)

| 기간 | 새로 채운 칸 |
|---|---:|
| 2503 · 2506 · 2509 | 각 3 |
| 2512 | 1 |
| **2603** | **57** |
| **2606** | **58** |
| **합계** | **125** |

과거 기간(2212~2512)을 정답과 대조하면 **불일치 0** 을 유지한다.
주로 들어간 칸은 예실차(111·112·115), CSM 변동(77·79·80·88), 손해율(107·108·109),
K-ICS(92·93·94), 자산운용률(23) 이다.

## CSM movement 검산 — 정답 시트 없는 기간을 검증하는 법 (`fix_csm.py`)

2603·2606 은 정답 시트가 없어 대조할 데가 없다. 그런데 CSM 변동은 정의상
반드시 닫히는 항등식이 있다.

```
기시(Col80) + 신계약(Col76) + 이자부리(Col79) − 상각(Col77) + 조정(Col78) = 기말(Col75)
```

안 맞으면 그 행 어딘가가 틀린 것이고, **어긋난 양이 어떤 항의 정확히 2배면
그 항의 부호가 뒤집힌 것**이다. 이 성질로 자동 판정한다.

실제로 2603·2606 에서 세 건이 잡혔다.

| 기간 | 회사 | 증상 | 판정 |
|---|---|---|---|
| 2603 | 삼성생명 | 검산 14,004,283 ≠ 기말 13,647,213 (+2.62%) | Col78 부호 반전 (+178,535 → −178,535) |
| 2606 | 삼성생명 | 검산 15,178,967 ≠ 기말 13,741,297 (+10.46%) | Col78 부호 반전 (+718,835 → −718,835) |
| 2603 | 메리츠화재 | 검산 11,291,732 ≠ 기말 10,884,319 (+3.74%) | Col75 가 틀림. 검산값과 IR 팩트시트가 11,291,732 로 일치 |

고친 뒤 전 기간 **46행 전부 검산 통과(불일치 0)**.

```bash
python fix_csm.py "2026_동업사 공시비교_DATA_작업_260824.xlsx" --periods 2603,2606 --dry
```

## 미래에셋 2603·2606 — CSM 2배 잔재 정리 (`fix_mirae_2x.py` → `parse_annual.py` 코드 수정으로 완료)

`parse_annual.py` 에 `DIVIDEND_CLASSIFICATION_PREFIXES` 필터를 CSM 잔액 추출뿐 아니라
**CSM 변동(`parse_csm_movement`)** 과 **전기말 CSM(`parse_prior_csm`)** 에도 적용했다.
사내망에서 `--period 2603/2606` 을 재실행해 DART 원본값을 직접 확인했다.

| | 고치기 전 | DART 원본(수정 후) | 근거 |
|---|---:|---:|---|
| 2603 Col73 BEL | 47,258,992 | **23,629,496** | DIVIDEND 필터 적용, DART 정확히 2배였음 |
| 2603 Col74 RA | 838,580 | **419,290** | 〃 |
| 2603 Col75 CSM | 4,301,298 | **2,150,649** | 〃 |
| 2603 Col79 이자부리 | 25,384 | **12,692** | `parse_csm_movement` DIVIDEND 필터 추가 |
| 2603 Col80 전기말CSM | NaN | **2,058,423** | `parse_prior_csm` 5-dim 패턴A2 추가 |
| 2606 Col73 BEL | 51,152,658 | **25,576,329** | DIVIDEND 필터 적용 |
| 2606 Col74 RA | 893,952 | **446,976** | 〃 |
| 2606 Col75 CSM | 4,662,200 | **2,331,100** | 〃 |
| 2606 Col78 조정 | 100,757 | **50,379** | `parse_csm_movement` DIVIDEND 필터 추가 |
| 2606 Col79 이자부리 | 55,120 | **27,560** | 〃 |
| 2606 Col80 전기말CSM | NaN | **2,058,423** | `parse_prior_csm` 5-dim 패턴A2 추가 |

movement 검산: 2,058,423 + 307,568 + 27,560 − 112,830 + 50,379 = **2,331,100** = 기말 CSM ✓

> **Col73·74 (BEL·RA)**: DATA 정의(순부채기준)에 해당하는 단일 XBRL 태그가 없어 `InsuranceContractsThatAreLiabilities` 합산값을 사용. DIVIDEND 제외 후 기존 팩트시트 수동 보정값(÷2)과 정확히 일치함을 DART 원본 조회로 확인.  
> **삼성생명 Col78(CSM조정) 부호**: DART 원본은 **양수(+)** — `+178,535`(2603) / `+718,835`(2606). `fix_csm.py`가 부호를 뒤집은 것은 수정 전 코드의 2배 오류가 검산을 틀어놓았기 때문. 이제 정상 추출 후 검산이 닫힘.  
> **메리츠화재 2603 Col75(CSM)**: DART 원본 `10,884,319`. 팩트시트 기반 `11,291,732`로 수동 수정된 이력이 있으나 DART가 정답이면 원래 값이 맞다. 현재 엑셀에는 DART 원본값이 채워져 있다.
