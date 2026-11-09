# 보험사 분기 공시 → Excel 입력

12개 보험사의 분기 공시를 모아 기존 Excel 양식에 입력하는 프로그램입니다.
**코드가 1차로 채우고, ChatGPT가 PDF 페이지를 읽어 빈칸을 보충하거나 입증된 오류를 고칩니다.**
담당자는 이 폴더를 연결한 ChatGPT 프로젝트 채팅에서 **“202609 분기보고서 채워줘”**라고 요청합니다.
ChatGPT가 코드 실행부터 PDF 검토·보정·최종 파일 저장까지 이어서 진행합니다.

## 한눈에 보는 흐름

![담당자 요청 → 공식 자료 수집 → 코드 1차 입력 → GPT PDF 2차 검토 → 검산 및 최종 Excel 전달](docs/quarterly_workflow.png)

| 누가 | 하는 일 |
|---|---|
| 담당자 | 사용할 Excel과 분기를 정하고 결과·보류 항목 확인 |
| 코드 | 공식 자료 수집, 숫자 추출, 단위 환산, Excel 입력, 보호·검산 |
| ChatGPT | 실제 PDF의 회사·표·기간·단위·회계 기준을 확인하고 보정 근거 작성 |

PDF와 숫자가 다르다고 바로 바꾸지는 않습니다. 같은 항목의 정밀값은 유지하고,
단위·부호·기간·계정 매핑·추출 오류가 확인됐을 때 수정합니다.

## 처음 받는 담당자: 한 번만 준비

1. **폴더를 PC에 저장합니다.** 전달받은 압축파일을 풀거나 GitHub 저장소를 내려받습니다. 이 `README.md`와 `quarter_pipeline.py`가 들어 있는 폴더를 사용합니다.
2. **실행 환경을 준비합니다.** Windows에 Python 3.10 이상을 설치하고 `처음설정.cmd`를 더블클릭합니다. 설치가 어려우면 이 단계만 IT 담당자에게 요청합니다.
3. **DART 키를 저장합니다.** 폴더의 `.env`를 메모장으로 열어 아래 항목에 본인의 인증키를 넣고 저장합니다. 키를 채팅에 붙여 넣을 필요는 없습니다.

   ```dotenv
   DART_API_KEY=발급받은_인증키
   ```

4. **본인 계정의 ChatGPT에서 폴더를 연결합니다.** 로컬 프로젝트를 만들거나 프로젝트 메뉴의 `Edit project → Add folder`에서 위 폴더를 추가합니다. 메뉴 이름은 앱 버전에 따라 다를 수 있습니다. 로컬 파일 읽기·명령 실행·PDF 이미지 열기가 지원되는 환경에서 진행합니다.
5. **그 프로젝트 안에서 새 채팅을 엽니다.** 아래 첫 설정 요청을 보냅니다. 전달용 ZIP에는 `작업양식.xlsx`가 들어 있습니다. 다른 Excel을 사용할 때는 실제 파일 경로를 함께 알려줍니다.

   > 이 폴더의 README.md와 AGENTS.md를 읽고 분기 입력 준비 상태를 확인해줘.
   > 작업 파일은 작업양식.xlsx야. DART 키가 설정됐는지 값은 출력하지 말고 확인하고,
   > 사용할 Excel 경로를 알려줘. 아직 데이터 입력은 시작하지 마.

웹 ChatGPT 프로젝트에 파일만 업로드한 경우에는 PC의 이 폴더나 실행 환경이 자동으로 연결되지 않습니다.
ChatGPT가 로컬 파일을 읽거나 명령을 실행할 수 없다고 하면 초기 설정 담당자에게 **로컬 프로젝트 연결과 실행 권한**을 확인해 달라고 요청합니다.
[프로젝트와 로컬 폴더에 관한 공식 안내](https://learn.chatgpt.com/docs/projects)

**`.env`, `quarterly.local.json`, `.venv`와 로그인 정보는 이전 담당자의 것을 그대로 넘기지 않습니다.**
전달용 압축 도구는 이를 제외합니다. 받는 담당자가 자기 PC와 계정으로 설정합니다.

## 매 분기: 프로젝트 채팅에서 이렇게 요청합니다

### 1. 등록한 프로젝트의 채팅을 엽니다

위에서 폴더를 연결한 프로젝트를 선택하고 새 채팅을 엽니다. 이전 담당자의 채팅을 옮길 필요는 없습니다.
ChatGPT는 폴더에 저장된 실행 지침·회사별 출처 힌트·과거 오류 메모를 읽고 작업합니다.

### 2. 채울 분기를 말합니다

작업 파일이 설정되어 있다면 아래 한 문장으로 요청합니다.

> 202609 분기보고서 채워줘.

**202609는 2026년 9월 말, 즉 2026년 3분기입니다.** ChatGPT는 이를 프로그램의 내부 분기 코드 `2609`로 바꿔 실행합니다.
2026년 6월은 `202606`, 2026년 12월은 `202612`처럼 말하면 됩니다.

작업 Excel까지 지정하거나 전체 절차를 명시하려면 다음 예시를 사용합니다.

> 작업 파일은 C:/업무/공시비교/작업양식.xlsx야. 202609 분기보고서 채워줘.
> README.md와 AGENTS.md, docs의 회사별 출처 힌트와 오류 메모를 참고해줘.
> 코드로 1차 입력한 다음 GPT가 PDF 페이지 이미지를 직접 보고 빈칸과 잘못된 값을 검토해줘.
> 같은 항목의 정밀값은 유지하고, 기존 기간과 수식·CALC 셀은 보존해줘. FISIS는 쓰지 마.
> 검산과 최종 저장까지 진행하고, 결과 Excel 링크와 남은 미확인 항목을 알려줘.

### 3. ChatGPT가 수집부터 최종 저장까지 진행합니다

| 순서 | ChatGPT가 하는 일 | 담당자가 확인할 내용 |
|---|---|---|
| ① 자료 수집 | 해당 분기의 DART 공시·공식 IR 자료와 PDF 확보 | 공시가 아직 안 올라온 회사가 있는지 |
| ② 코드 1차 입력 | API·XBRL·XML·팩트시트 값을 추출해 `stage1.xlsx` 저장 | 사용할 작업 파일과 대상 분기가 맞는지 |
| ③ GPT 2차 검토 | 회사별 힌트를 참고해 PDF 페이지 이미지를 직접 확인, 빈칸 보충·입증된 오류 보정 | 기준 차이로 보류한 항목이 있는지 |
| ④ 검산·최종 저장 | 기존 데이터·수식 보호 확인, 결과와 변경 근거 저장 | 최종 파일, 미해결 목록, 재계산 상태 |

담당자가 검토용 JSON을 작성하거나 다음 명령을 직접 실행할 필요는 없습니다.
**1차 입력이 끝났다는 안내만으로 작업이 완료된 것은 아닙니다.** PDF 직접 검토와 최종 저장까지 진행해야 합니다.
공시 미게시·수집 실패·기준 충돌은 그대로 보고하며 추정값을 넣지 않습니다.

### 4. 결과 파일과 보류 항목을 확인합니다

ChatGPT의 완료 답변에서 **최종 Excel 링크**를 열고, 함께 안내한 미해결 항목과 재계산 상태를 확인합니다.
202609 작업 결과는 `runs/2609/실행ID/`에 저장됩니다. 원본 Excel은 덮어쓰지 않습니다.

추가 확인이 필요하면 같은 프로젝트 채팅에서 이렇게 요청합니다.

> 방금 작업한 결과 파일 위치와 회사별로 남은 빈칸·보류 이유를 보여줘.

> 방금 실행에서 미확인으로 남은 항목을 공식 PDF에서 추가 확인하고, 확인된 것만 결과 파일에 반영해줘.

> 중단된 202609 작업의 실행 폴더를 확인해서 이어서 진행해줘. 중복 실행 전에 기존 진행 상태부터 확인해줘.

### 선택 사항: 분기 선택 화면으로 실행

채팅 대신 `분기입력.cmd`에서 **연도·분기·작업 Excel**을 고를 수도 있습니다.
이 경로는 PDF 검토용 **Codex CLI 설치와 본인 계정 로그인**이 별도로 필요합니다. 초기 설정 담당자가 지원하면 됩니다.
ChatGPT 프로젝트 채팅에서 직접 진행하는 위 절차와 구분합니다.
[Codex CLI 공식 안내](https://learn.chatgpt.com/docs/codex/cli) · [로그인 안내](https://learn.chatgpt.com/docs/auth)

## 결과는 어디에 있나

원본을 덮어쓰지 않고 `runs/분기/실행ID/`에 저장합니다.

| 파일 | 확인할 내용 |
|---|---|
| `final.xlsx` | PDF 검토·보정을 반영한 결과 Excel |
| `summary.md` / `summary.json` | 입력·검토·수정 수, 검산과 재계산 상태 |
| `unresolved.csv` | 아직 못 찾은 값, 기준 충돌, 검토하지 못한 항목 |
| `changes.json` | 수정 전후값과 PDF 페이지·사유 |
| `stage1.xlsx` | 코드가 채운 1차 결과 |
| `collection_issues.json` | 공시 미게시·다운로드 실패 등 수집 문제 |

`completed_with_unresolved`는 최종 파일은 저장됐지만 확인할 항목이 남았다는 뜻입니다.
`pending_excel_recalculation`이면 수식 재계산이 남았습니다. LibreOffice가 있으면 자동 재계산하고,
Windows Excel에서는 아래 개발 명령으로 실행 사본만 재계산할 수 있습니다.

## 과거에 정리한 힌트는 어디에 있나

**아래 세 JSON은 다음 PDF 검토 프롬프트에 전달됩니다. 과거 정답 숫자를 입력값으로 재사용하는 파일이 아닙니다.**

| 파일 | 내용 |
|---|---|
| [quarterly_lessons.json](docs/quarterly_lessons.json) | 누계·잠정/최종·연결/별도·부호 등 과거 실수와 재발 방지 기준 |
| [review_navigation_2025.json](docs/review_navigation_2025.json) | 직접 읽었던 회사·칼럼별 PDF 표·행·열·단위·페이지 예시 |
| [company_field_sources.json](docs/company_field_sources.json) | 회사별 출처, 팩트시트 위치, 항목 해석 안내 |

사람이 읽을 때는 [회사별 출처 안내](docs/회사별_칼럼별_출처안내.md),
[오류 원인 메모](docs/2025_오류원인메모.md), [2025년 검증 현황](docs/2025_검증현황.md)을 봅니다.
ChatGPT가 새 대화에서 지켜야 할 실행 지침은 [AGENTS.md](AGENTS.md),
2609 작업 절차는 [2609 실행 지침](docs/2609_실행지침.md)에 있습니다.
과거 페이지 번호·표 구조는 힌트이며 현재 분기의 원문에서 다시 확인합니다.

## 입력과 보호 기준

- FISIS는 사용자 요청으로 제외했습니다. `use_fisis=false`를 유지합니다.
- `DATA_작업_빈칸_2512backup.xlsx`는 정답 기준입니다. 작업 양식으로 사용하지 않습니다.
- 일반 실행에서는 2512까지 보호하고, 선택 분기 이전의 기존 데이터도 보호합니다.
- CALC·CHECK 셀을 채우지 않으며 기존 수식과 과거 기간을 보존합니다.
- 미게시·미확인 값은 빈칸/미해결로 남깁니다. 다른 분기 값이나 역산값으로 대신하지 않습니다.
- 범위·발표 시점이 다른 값은 충돌로 보류합니다. 반올림값으로 기존 정밀값을 덮지 않습니다.

## 설정·명령어: 설정 담당자용

설정 우선순위는 **명령줄 `--workbook` → `quarterly.local.json` → `quarterly.json`**입니다.
개인 경로는 `quarterly.local.json`에 저장합니다.

```json
{"workbook":"C:/업무/작업양식.xlsx","use_fisis":false}
```

```powershell
# 현재 대화의 ChatGPT가 직접 PDF 검토를 이어서 수행
.venv\Scripts\python.exe quarter_pipeline.py --period 2609 --reviewer assistant

# 다른 Excel 지정
.venv\Scripts\python.exe quarter_pipeline.py --period 2609 --reviewer assistant --workbook "C:/업무/작업양식.xlsx"

# Codex CLI가 PDF 검토까지 실행
.venv\Scripts\python.exe quarter_pipeline.py --period 2609

# 중단된 CLI 검토 재개 / 현재 ChatGPT가 작성한 검토 결과 반영
.venv\Scripts\python.exe quarter_pipeline.py --resume "runs/2609/실행ID"
.venv\Scripts\python.exe quarter_pipeline.py --finalize "runs/2609/실행ID"

# Windows Excel로 실행 사본 재계산
powershell -NoProfile -File recalculate_excel.ps1 -Path "runs/2609/실행ID/final.xlsx"

# 정답본 대조는 읽기 전용 / 자동 테스트
.venv\Scripts\python.exe quarter_pipeline.py --check-baseline --period 2512
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m pytest tests -q
```

**`--reviewer assistant` 명령만으로는 완료되지 않습니다.** 자료 수집·1차 입력·검토 준비 후
현재 ChatGPT가 PDF 이미지를 확인하고 `--finalize`까지 수행해야 합니다.
공식 사이트 수집 실패 시에는 ChatGPT/운영 담당자의 추가 확인이 필요할 수 있습니다.

## 폴더 전달과 유지보수

전달용 묶음을 만들려면 아래 명령을 사용합니다. 코드·지침·작업 양식·정답본을 묶고,
개인 키·개인 경로·로그인·가상환경·과거 실행 로그는 제외합니다. 원본 파일은 수정하지 않습니다.

```powershell
.venv\Scripts\python.exe tools/export_handoff.py
```

공유 폴더에는 코드뿐 아니라 **실제로 사용할 작업 Excel과 정답본**도 있어야 합니다.
GitHub에 전달하려면 이 로컬 변경을 별도로 커밋·업로드해야 합니다. 폴더 정리만으로 원격 저장소가 갱신되지는 않습니다.

| 경로 | 역할 |
|---|---|
| `quarter_pipeline.py`, `quarterly/` | 실제 분기 수집·입력·PDF 검토·최종 저장 |
| `dart_api.py`, `parse_annual.py`, `ir_parse.py`, `ir_pdf.py` | 수집·추출 규칙 |
| `quarter_picker.py`, `분기입력.cmd`, `처음설정.cmd` | 담당자용 실행·초기 설정 |
| `docs/` | 출처 힌트·오류 메모·검증 현황 |
| `tools/`, `backtest_2512.py`, `tests/` | 개발·검증·인수인계 도구 |
| `data/` | 받은 공시 자료 |
| `runs/` | 실행 결과·원문 사본·검증 근거·정리 파일의 복구용 사본, Git 제외 |

과거 분기 재현 검증은 실제 정답본을 수정하지 않고 별도 사본에서 수행합니다.
검증 결과를 확정한 뒤 채점하며, 기준 차이·미해결 원인은 그대로 기록합니다.
