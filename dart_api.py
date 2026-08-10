import os
import io
import csv
import zipfile
import xml.etree.ElementTree as ET
import requests
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("DART_API_KEY")
BASE_URL = "https://opendart.fss.or.kr/api"
TAXONOMY_CSV = os.path.join("data", "taxonomy.csv")

_corp_map = None  # 기업명 -> 기업코드 캐시


def build_taxonomy_from_lab(lab_ko_path: str) -> dict:
    """_lab-ko.xml 파싱해서 account_id -> 한글명 매핑 반환"""
    tree = ET.parse(lab_ko_path)
    root = tree.getroot()
    ns = {
        "link": "http://www.xbrl.org/2003/linkbase",
        "xlink": "http://www.w3.org/1999/xlink",
    }

    # loc: xlink:label -> account_id 매핑
    loc_to_id = {}
    for loc in root.findall(".//link:loc", ns):
        href = loc.get("{http://www.w3.org/1999/xlink}href", "")
        label = loc.get("{http://www.w3.org/1999/xlink}label", "")
        if "#" in href:
            account_id = href.split("#")[-1]
            loc_to_id[label] = account_id

    # labelArc: loc label -> label label 연결
    loc_to_label = {}
    for arc in root.findall(".//link:labelArc", ns):
        frm = arc.get("{http://www.w3.org/1999/xlink}from", "")
        to = arc.get("{http://www.w3.org/1999/xlink}to", "")
        role = arc.get("{http://www.w3.org/1999/xlink}arcrole", "")
        if "concept-label" in role:
            loc_to_label.setdefault(frm, []).append(to)

    # label: xlink:label -> 한글 텍스트
    label_to_text = {}
    for lbl in root.findall(".//link:label", ns):
        lbl_key = lbl.get("{http://www.w3.org/1999/xlink}label", "")
        role = lbl.get("{http://www.w3.org/1999/xlink}role", "")
        lang = lbl.get("{http://www.w3.org/XML/1998/namespace}lang", "")
        if lang == "ko" and role == "http://www.xbrl.org/2003/role/label" and lbl.text:
            label_to_text[lbl_key] = lbl.text.strip()

    # 조합
    taxonomy = {}
    for loc_label, account_id in loc_to_id.items():
        for lbl_label in loc_to_label.get(loc_label, []):
            if lbl_label in label_to_text:
                taxonomy[account_id] = label_to_text[lbl_label]
                break

    return taxonomy


def save_taxonomy_csv(lab_ko_path: str) -> str:
    """_lab-ko.xml 파싱 후 기존 taxonomy.csv에 머지"""
    print(f"파싱 중: {lab_ko_path}")
    new_data = build_taxonomy_from_lab(lab_ko_path)
    os.makedirs("data", exist_ok=True)

    # 기존 CSV 읽기
    existing = {}
    if os.path.exists(TAXONOMY_CSV):
        with open(TAXONOMY_CSV, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                existing[row["account_id"]] = row["label_kor"]

    before = len(existing)
    existing.update(new_data)
    added = len(existing) - before

    with open(TAXONOMY_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["account_id", "label_kor"])
        for account_id, label in sorted(existing.items()):
            writer.writerow([account_id, label])

    print(f"  추가: {added}개 / 누적: {len(existing)}개")
    return TAXONOMY_CSV


def update_taxonomy_from_xbrl_dir(xbrl_dir: str):
    """XBRL 폴더에서 _lab-ko.xml 찾아서 taxonomy 머지"""
    for fname in os.listdir(xbrl_dir):
        if fname.endswith("_lab-ko.xml"):
            save_taxonomy_csv(os.path.join(xbrl_dir, fname))
            return
    print(f"  _lab-ko.xml 없음: {xbrl_dir}")


def parse_xbrl(xbrl_path: str) -> dict:
    """
    XBRL 파일에서 단순 context(연결/별도 축 하나만 있는)의 수치 추출
    반환: {account_id: {"연결_당기": val, "연결_전기": val, "별도_당기": val, ...}}
    """
    tree = ET.parse(xbrl_path)
    root = tree.getroot()
    ns_xbrli = "http://www.xbrl.org/2003/instance"
    ns_xbrldi = "http://xbrl.org/2006/xbrldi"

    # context 분류: id -> {fs, kind, ...}
    ctx_info = {}
    for ctx in root.findall(f"{{{ns_xbrli}}}context"):
        members = ctx.findall(f".//{{{ns_xbrldi}}}explicitMember")
        if len(members) != 1:
            continue
        dim = members[0].get("dimension", "")
        val = (members[0].text or "").strip()
        if "ConsolidatedAndSeparateFinancialStatements" not in dim:
            continue

        is_con = "Consolidated" in val
        period = ctx.find(f"{{{ns_xbrli}}}period")
        instant = period.find(f"{{{ns_xbrli}}}instant")
        start = period.find(f"{{{ns_xbrli}}}startDate")
        end = period.find(f"{{{ns_xbrli}}}endDate")

        fs = "연결" if is_con else "별도"
        ctx_id = ctx.get("id", "")

        if instant is not None:
            ctx_info[ctx_id] = {"fs": fs, "kind": "시점", "date": instant.text}
        elif start is not None and end is not None:
            # FQQ=단일분기, FQA=누적, 연간 등 구분
            duration_type = "누적" if "FQA" in ctx_id or "FYA" in ctx_id else "단일"
            ctx_info[ctx_id] = {"fs": fs, "kind": "기간", "start": start.text, "end": end.text, "dtype": duration_type}

    if not ctx_info:
        return {}

    # 당기/전기 레이블 결정 (날짜 기준)
    instant_dates = sorted(set(v["date"] for v in ctx_info.values() if v["kind"] == "시점"), reverse=True)
    instant_label = {d: ("당기말" if i == 0 else "전기말") for i, d in enumerate(instant_dates[:2])}

    # 기간: (end, dtype) 조합으로 레이블
    period_keys = sorted(set((v["end"], v["dtype"]) for v in ctx_info.values() if v["kind"] == "기간"), reverse=True)
    period_label = {}
    cur_cnt, prev_cnt = 0, 0
    for end, dtype in period_keys:
        if cur_cnt == 0:
            period_label[(end, dtype)] = f"당기_{dtype}"
            cur_cnt += 1
        elif prev_cnt == 0:
            period_label[(end, dtype)] = f"전기_{dtype}"
            prev_cnt += 1
        else:
            period_label[(end, dtype)] = f"{end}_{dtype}"

    def col_name(info):
        fs = info["fs"]
        if info["kind"] == "시점":
            lbl = instant_label.get(info["date"], info["date"])
        else:
            lbl = period_label.get((info["end"], info["dtype"]), info["end"])
        return f"{fs}_{lbl}"

    # 수치 추출
    result = {}
    for elem in root:
        tag = elem.tag
        ctx_ref = elem.get("contextRef", "")
        if ctx_ref not in ctx_info or not elem.text:
            continue
        # account_id 구성
        if "}" in tag:
            ns_uri, local = tag[1:].split("}", 1)
            # 네임스페이스 -> prefix
            if "ifrs-full" in ns_uri or "ifrs_full" in ns_uri:
                prefix = "ifrs-full"
            elif "dart" in ns_uri and "dart-gcd" not in ns_uri:
                prefix = "dart"
            elif "dart-gcd" in ns_uri:
                prefix = "dart-gcd"
            else:
                continue
            account_id = f"{prefix}_{local}"
        else:
            continue

        col = col_name(ctx_info[ctx_ref])
        val = elem.text.strip()
        # 숫자만 저장 (텍스트블록 제외)
        try:
            float(val)
        except ValueError:
            continue

        if account_id not in result:
            result[account_id] = {}
        result[account_id][col] = val

    return result


def build_financials_csv(companies: list, output_path: str):
    """
    companies: [{"name": "삼성생명", "xbrl_path": "..."}]
    taxonomy.csv의 account_id 기준으로 행, 회사×기간 기준으로 열 구성
    """
    # taxonomy 읽기
    taxonomy = {}
    if os.path.exists(TAXONOMY_CSV):
        with open(TAXONOMY_CSV, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                taxonomy[row["account_id"]] = row["label_kor"]

    # 각 회사 데이터 파싱
    all_data = {}
    all_cols = []
    for comp in companies:
        name = comp["name"]
        data = parse_xbrl(comp["xbrl_path"])
        all_data[name] = data
        for cols in data.values():
            for col in cols:
                full_col = f"{name}_{col}"
                if full_col not in all_cols:
                    all_cols.append(full_col)

    # 전체 account_id 수집 (taxonomy 순서 우선)
    all_ids = list(taxonomy.keys())
    extra = [aid for comp_data in all_data.values() for aid in comp_data if aid not in taxonomy]
    for aid in extra:
        if aid not in all_ids:
            all_ids.append(aid)

    # CSV 작성
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["account_id", "label_kor"] + all_cols)
        for aid in all_ids:
            row_vals = [aid, taxonomy.get(aid, "")]
            has_data = False
            for full_col in all_cols:
                # full_col = "{회사명}_{fs}_{기간}" 에서 회사명 분리
                comp_name = full_col.split("_")[0]
                col_key = full_col[len(comp_name)+1:]
                val = all_data.get(comp_name, {}).get(aid, {}).get(col_key, "")
                if val:
                    has_data = True
                row_vals.append(val)
            if has_data:
                writer.writerow(row_vals)

    print(f"저장 완료: {output_path}")


def _load_corp_map():
    global _corp_map
    if _corp_map is not None:
        return _corp_map

    response = requests.get(f"{BASE_URL}/corpCode.xml", params={"crtfc_key": API_KEY})
    response.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        with z.open("CORPCODE.xml") as f:
            tree = ET.parse(f)

    _corp_map = {}
    for item in tree.getroot().findall("list"):
        name = item.findtext("corp_name", "").strip()
        code = item.findtext("corp_code", "").strip()
        if name and code:
            _corp_map[name] = code

    return _corp_map


def find_corp_code(corp_name: str) -> str:
    """기업명으로 기업코드 반환. 정확히 일치하지 않으면 부분 일치 후보 출력."""
    corp_map = _load_corp_map()

    if corp_name in corp_map:
        return corp_map[corp_name]

    candidates = [name for name in corp_map if corp_name in name]
    if not candidates:
        raise ValueError(f"기업을 찾을 수 없습니다: {corp_name}")
    if len(candidates) == 1:
        print(f"[{corp_name}] -> [{candidates[0]}] 으로 검색합니다.")
        return corp_map[candidates[0]]

    print(f"여러 기업이 검색됩니다. 하나를 선택하세요:")
    for i, name in enumerate(candidates[:10]):
        print(f"  {i+1}. {name} ({corp_map[name]})")
    idx = int(input("번호 입력: ")) - 1
    return corp_map[candidates[idx]]


def get_disclosures(corp_code: str = None, bgn_de: str = None, end_de: str = None, page_no: int = 1, page_count: int = 10, last_reprt_at: str = "N"):
    """경영공시 목록 조회"""
    params = {
        "crtfc_key": API_KEY,
        "page_no": page_no,
        "page_count": page_count,
        "last_reprt_at": last_reprt_at,
    }
    if corp_code:
        params["corp_code"] = corp_code
    if bgn_de:
        params["bgn_de"] = bgn_de
    if end_de:
        params["end_de"] = end_de

    response = requests.get(f"{BASE_URL}/list.json", params=params)
    response.raise_for_status()
    return response.json()


def get_company_info(corp_code: str):
    """기업 기본정보 조회"""
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
    }
    response = requests.get(f"{BASE_URL}/company.json", params=params)
    response.raise_for_status()
    return response.json()


# 보고서 코드
REPRT_CODE = {
    "사업보고서": "11011",
    "반기보고서": "11012",
    "1분기보고서": "11013",
    "3분기보고서": "11014",
}

# 재무제표 구분
FS_DIV = {
    "연결": "CFS",  # 연결재무제표
    "별도": "OFS",  # 별도재무제표
}


def get_financial_statements(corp_code: str, bsns_year: str, reprt_code: str = "11012", fs_div: str = "CFS"):
    """XBRL 파싱된 재무제표 JSON 조회"""
    params = {
        "crtfc_key": API_KEY,
        "corp_code": corp_code,
        "bsns_year": bsns_year,
        "reprt_code": reprt_code,
        "fs_div": fs_div,
    }
    response = requests.get(f"{BASE_URL}/fnlttSinglAcntAll.json", params=params)
    response.raise_for_status()
    return response.json()


def find_rcept_no(corp_code: str, bsns_year: str, reprt_code: str = "11012") -> str:
    """공시 목록에서 해당 보고서의 접수번호(rcept_no) 찾기"""
    bgn_de = f"{bsns_year}0101"
    end_de = f"{bsns_year}1231"

    report_keyword_map = {"11011": "사업보고서", "11012": "반기보고서", "11013": "분기보고서", "11014": "분기보고서"}
    target_keyword = report_keyword_map.get(reprt_code, "")

    # 전체 목록에서 원본(기재정정 아닌 것) 우선 탐색
    data = get_disclosures(corp_code=corp_code, bgn_de=bgn_de, end_de=end_de, page_count=100, last_reprt_at="N")
    items = data.get("list", [])
    candidates = [item for item in items if target_keyword in item.get("report_nm", "")]

    if not candidates:
        raise ValueError(f"해당 보고서를 찾을 수 없습니다. (연도: {bsns_year}, 코드: {reprt_code})")

    # 가장 최근 제출본 선택 (기재정정 포함)
    selected = max(candidates, key=lambda x: x.get("rcept_dt", ""))

    print(f"보고서 선택: {selected['report_nm']} ({selected['rcept_dt']})")
    return selected["rcept_no"]


def _fetch_document_content(rcept_no: str) -> str:
    """document.xml ZIP에서 본문 XML 문자열 반환"""
    import re as _re
    resp = requests.get(f"{BASE_URL}/document.xml", params={"crtfc_key": API_KEY, "rcept_no": rcept_no})
    if not resp.content.startswith(b"PK"):
        raise ValueError(f"document.xml 응답 오류: {resp.content[:100]}")
    with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
        return z.read(z.namelist()[0]).decode("utf-8")


def fetch_kics_ratio(rcept_no: str) -> dict:
    """공시서류에서 지급여력비율(킥스), 가용자본, 요구자본 추출.
    반환: {킥스비율_당기, 킥스비율_전기, 가용자본_당기, 가용자본_전기, 요구자본_당기, 요구자본_전기}
    단위: 비율=%, 자본=백만원
    """
    import re
    content = _fetch_document_content(rcept_no)

    def _extract_best(pattern):
        matches = list(re.finditer(pattern, content, re.DOTALL))
        best = []
        for m in matches:
            nums = re.findall(r">\s*([\d,.]+)\s*</T[DH]>", m.group())
            if len(nums) > len(best):
                best = nums
        return best

    ratio_best = _extract_best(r"지급여력비율\(A/B\).*?</TR>")
    avail_best = _extract_best(r"지급여력금액\s*\(A\).*?</TR>")
    req_best = _extract_best(r"지급여력기준[금액]*\s*\(B\).*?</TR>")

    def _get(lst, idx):
        return lst[idx].replace(",", "") if len(lst) > idx else ""

    return {
        "킥스비율_당기": _get(ratio_best, 0),
        "킥스비율_전기": _get(ratio_best, 1),
        "가용자본_당기": _get(avail_best, 0),
        "가용자본_전기": _get(avail_best, 1),
        "요구자본_당기": _get(req_best, 0),
        "요구자본_전기": _get(req_best, 1),
    }


def fetch_new_csm(rcept_no: str) -> dict:
    """공시서류에서 신계약 CSM(당기 최초 인식 계약, 연결 기준) 추출.
    반환: {신계약CSM_당기, 신계약CSM_전기} (단위: 백만원, ADECIMAL=-6 적용된 표시값)
    """
    import re
    content = _fetch_document_content(rcept_no)

    # ACODE=IncreaseDecreaseThroughEffectsOfContractsInitiallyRecognised
    # + ContractualServiceMarginMember + InsuranceContractsIssuedMember 조건
    pattern = re.compile(
        r'<TE[^>]*ACODE="ifrs-full_IncreaseDecreaseThroughEffectsOfContractsInitiallyRecognisedInPeriodInsuranceContractsLiabilityAsset"'
        r'[^>]*ACONTEXT="([^"]+)"[^>]*>([^<]*)</TE>',
        re.DOTALL,
    )

    cur_val, prev_val = None, None
    for m in pattern.finditer(content):
        ctx, val = m.group(1), m.group(2).strip()
        if "ContractualServiceMarginMember" not in ctx:
            continue
        if "InsuranceContractsIssuedMember" not in ctx:
            continue
        if "ConsolidatedMember" not in ctx:
            continue
        # 보험종류 축 없이 합계인 경우 or 있어도 합산 — 여기선 보험종류 축 없는 행만 선택
        # 보험종류 축 없는 context는 TypesOfContractsAxis가 없음
        if "TypesOfContractsAxis" in ctx:
            continue
        num = val.replace(",", "").replace("(", "-").replace(")", "").strip()
        if "CFY2026" in ctx and cur_val is None:
            cur_val = num
        elif "PFY2025" in ctx and prev_val is None:
            prev_val = num

    return {
        "신계약CSM_당기": cur_val or "",
        "신계약CSM_전기": prev_val or "",
    }


def download_xbrl(rcept_no: str, corp_name: str, reprt_name: str, year: str, reprt_code: str) -> str:
    """XBRL 원본 ZIP 다운로드 후 data/{기업명}/{보고서분류}/{연도}/ 에 저장"""
    params = {
        "crtfc_key": API_KEY,
        "rcept_no": rcept_no,
        "reprt_code": reprt_code,
    }
    response = requests.get(f"{BASE_URL}/fnlttXbrl.xml", params=params)
    response.raise_for_status()

    # ZIP이 아니면 API 에러 메시지 출력
    if not response.content.startswith(b"PK"):
        raise ValueError(f"XBRL 파일을 받지 못했습니다. API 응답: {response.text}")

    save_path = os.path.join("data", corp_name, reprt_name, year)
    os.makedirs(save_path, exist_ok=True)

    with zipfile.ZipFile(io.BytesIO(response.content)) as z:
        z.extractall(save_path)

    print(f"저장 완료: {save_path}")
    print(f"파일 목록: {os.listdir(save_path)}")
    return save_path


FISIS_API_KEY = "7b145068523cd3c2642e8b902a21999d"
FISIS_BASE = "http://fisis.fss.or.kr/openapi"

# 12개사 FISIS 코드 (finance_cd, lrgDiv)
FISIS_CORP_CODES = {
    "미래에셋생명": ("0010608", "H"),
    "삼성생명":    ("0010595", "H"),
    "교보생명":    ("0010597", "H"),
    "한화생명":    ("0010593", "H"),
    "신한라이프생명": ("0010599", "H"),
    "동양생명":    ("0010622", "H"),
    "KB생명":     ("0010616", "H"),
    "삼성화재":    ("0010633", "I"),
    "현대해상":    ("0010634", "I"),
    "메리츠화재":  ("0010626", "I"),
    "KB손해보험":  ("0010635", "I"),
    "DB손해보험":  ("0010636", "I"),
}


def fetch_fisis_kics(corp_name: str, base_month: str) -> dict:
    """FISIS API에서 K-ICS 지급여력비율/가용자본/요구자본 조회.
    base_month: 'YYYYMM' (예: '202503')
    반환: {95: 비율(소수), 96: 가용자본(백만원), 97: 요구자본(백만원)} 또는 {}
    단위: FISIS는 원 단위 → /1e6 → 백만원
    """
    import requests as _req
    entry = FISIS_CORP_CODES.get(corp_name)
    if not entry:
        return {}
    finance_cd, lrg_div = entry
    list_no = "SH021" if lrg_div == "H" else "SI021"
    try:
        r = _req.get(f"{FISIS_BASE}/statisticsInfoSearch.json", params={
            "lang": "kr", "auth": FISIS_API_KEY,
            "financeCd": finance_cd, "listNo": list_no,
            "term": "Q", "startBaseMm": base_month, "endBaseMm": base_month,
            "numOfRows": 10,
        }, timeout=10)
        items = r.json().get("result", {}).get("list", [])
    except Exception:
        return {}

    ratio = avail = req = None
    for it in items:
        v = it.get("a", "")
        if not v or v == "0":
            continue
        try:
            fv = float(v)
        except ValueError:
            continue
        if it["account_cd"] == "A" and fv > 0:
            ratio = fv / 100  # % → 소수
        elif it["account_cd"] == "B" and fv > 0:
            avail = fv / 1e6  # 원 → 백만원
        elif it["account_cd"] == "C" and fv > 0:
            req = fv / 1e6

    if ratio is None or avail is None or req is None:
        return {}
    return {95: ratio, 96: avail, 97: req}


def fetch_fisis_balance(corp_name: str, base_month: str) -> dict:
    """FISIS API에서 재무 항목 일괄 조회.
    커버 항목:
      Col16  신종자본증권       SH151/SI147 F3
      Col22  운용자산          SH150/SI146 A1
      Col73  BEL (생보만)      SH156/SI152 A111+A121+A3
      Col74  RA               SH156/SI152 A112+A122
      Col75  CSM              SH156/SI152 A113
      Col99  해약환급금준비금    SH151/SI147 F46
    반환: {col: value_백만원, ...} (값 없으면 해당 col 제외)
    """
    import requests as _req
    entry = FISIS_CORP_CODES.get(corp_name)
    if not entry:
        return {}
    fcd, lrg = entry
    is_life = (lrg == "H")

    def _fetch(list_no):
        try:
            r = _req.get(f"{FISIS_BASE}/statisticsInfoSearch.json", params={
                "lang": "kr", "auth": FISIS_API_KEY,
                "financeCd": fcd, "listNo": list_no,
                "term": "Q", "startBaseMm": base_month, "endBaseMm": base_month,
                "numOfRows": 200,
            }, timeout=10)
            return {it["account_cd"]: float(it["a"])
                    for it in r.json().get("result", {}).get("list", [])
                    if it.get("a") and it["a"] not in ("0", "")}
        except Exception:
            return {}

    result = {}

    # 책임준비금 (RA/CSM만 — BEL은 정의 불일치로 제외)
    r_data = _fetch("SH156" if is_life else "SI152")
    if r_data:
        ra = r_data.get("A112", 0) + r_data.get("A122", 0)
        csm = r_data.get("A113", 0)
        if ra:  result[74] = ra / 1e6
        if csm: result[75] = csm / 1e6

    # 재무상태표 자산 (운용자산, 총자산)
    # 한화생명/KB생명: FISIS A1 기준이 엑셀과 달라서 제외
    _NO_FISIS_COL22 = {"한화생명", "KB생명"}
    if corp_name not in _NO_FISIS_COL22 or not is_life:
        a_data = _fetch("SH150" if is_life else "SI146")
        if a_data:
            total_a = a_data.get("A", 0)
            if corp_name not in _NO_FISIS_COL22:
                run = a_data.get("A1", 0)
                if run:
                    if is_life:
                        if total_a > 0:
                            sa_data = _fetch("SH152")
                            sa_total = sa_data.get("A", 0)
                            if sa_total > 0 and abs(run + sa_total - total_a) / total_a < 0.015:
                                run += sa_total
                    result[22] = run / 1e6

    # 재무상태표 부채자본 (신종자본증권, 해약환급금, OCI세부)
    b_data = _fetch("SH151" if is_life else "SI147")
    if b_data:
        hyb = b_data.get("F3", 0)
        hyk = b_data.get("F46", 0)
        f64 = b_data.get("F64", 0)   # Col26 보험계약자산(부채)순금융손익
        f65 = b_data.get("F65", 0)   # Col27 재보험계약자산(부채)순금융손익
        f67 = b_data.get("F67", 0)   # Col28 현금흐름위험회피 (생보사만 정확)
        f69 = b_data.get("F69", 0)   # Col30 확정급여(보험수리적손익)
        if hyb: result[16] = hyb / 1e6
        if hyk: result[99] = hyk / 1e6
        if f64: result[26] = f64 / 1e6
        if f65: result[27] = f65 / 1e6
        if f67 and is_life: result[28] = f67 / 1e6   # 생보사만: 손보사 F67은 기준 불일치
        if f69: result[30] = f69 / 1e6

    return result


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="DART 경영공시 / 재무제표 조회")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # 공시 목록 조회
    p_disc = subparsers.add_parser("공시", help="경영공시 목록 조회")
    p_disc.add_argument("corp_name", help="기업명 (예: 삼성전자)")
    p_disc.add_argument("--bgn_de", default=None, help="시작일 (예: 20240101)")
    p_disc.add_argument("--end_de", default=None, help="종료일 (예: 20241231)")
    p_disc.add_argument("--page", type=int, default=1)
    p_disc.add_argument("--count", type=int, default=10)

    # 재무제표 JSON 조회
    p_fin = subparsers.add_parser("재무", help="XBRL 파싱 재무제표 JSON 조회")
    p_fin.add_argument("corp_name", help="기업명 (예: 삼성전자)")
    p_fin.add_argument("year", help="사업연도 (예: 2024)")
    p_fin.add_argument("--보고서", dest="reprt", default="반기보고서",
                       choices=list(REPRT_CODE.keys()), help="보고서 종류 (기본: 반기보고서)")
    p_fin.add_argument("--재무제표", dest="fs", default="연결",
                       choices=list(FS_DIV.keys()), help="재무제표 구분 (기본: 연결)")

    # 택사노미 CSV 생성
    p_tax = subparsers.add_parser("택사노미", help="_lab-ko.xml 파싱해서 taxonomy.csv 생성")
    p_tax.add_argument("lab_ko_path", help="_lab-ko.xml 파일 경로")

    # XBRL 원본 다운로드
    p_xbrl = subparsers.add_parser("xbrl", help="XBRL 원본 파일 다운로드")
    p_xbrl.add_argument("corp_name", help="기업명 (예: 삼성전자)")
    p_xbrl.add_argument("year", help="사업연도 (예: 2024)")
    p_xbrl.add_argument("--보고서", dest="reprt", default="반기보고서",
                        choices=list(REPRT_CODE.keys()), help="보고서 종류 (기본: 반기보고서)")

    args = parser.parse_args()

    if args.command == "택사노미":
        save_taxonomy_csv(args.lab_ko_path)
        exit(0)

    corp_code = find_corp_code(args.corp_name)

    if args.command == "공시":
        result = get_disclosures(
            corp_code=corp_code,
            bgn_de=args.bgn_de,
            end_de=args.end_de,
            page_no=args.page,
            page_count=args.count,
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.command == "재무":
        result = get_financial_statements(
            corp_code=corp_code,
            bsns_year=args.year,
            reprt_code=REPRT_CODE[args.reprt],
            fs_div=FS_DIV[args.fs],
        )
        print(json.dumps(result, ensure_ascii=False, indent=2))

    elif args.command == "xbrl":
        reprt_code = REPRT_CODE[args.reprt]
        print(f"접수번호 조회 중...")
        rcept_no = find_rcept_no(corp_code, args.year, reprt_code)
        print(f"접수번호: {rcept_no}")
        download_xbrl(rcept_no, corp_name=args.corp_name, reprt_name=args.reprt, year=args.year, reprt_code=reprt_code)
