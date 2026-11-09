from __future__ import annotations

from copy import copy
from pathlib import Path
import re
import math

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.styles import PatternFill
from openpyxl.workbook.properties import CalcProperties

from .core import Period, numeric, save_workbook


def resolve(ws, row, col=3):
    seen = set()
    while (row, col) not in seen:
        seen.add((row, col))
        v = ws.cell(row, col).value
        if not isinstance(v, str) or not v.startswith("="):
            return str(v).strip() if v is not None else None
        m = re.fullmatch(r"=\$?([A-Z]{1,3})\$?(\d+)", v.strip())
        if not m:
            return None
        row, col = int(m[2]), openpyxl.utils.column_index_from_string(m[1])
    raise ValueError("회사명 참조 수식에 순환이 있습니다.")


class DataBook:
    def __init__(self, path: Path, protect_through="2512"):
        self.path = Path(path)
        self.protect_through = str(protect_through)
        self.wb = openpyxl.load_workbook(path)
        self.ws = self.wb["DATA"]
        self.columns = {}
        for c in range(1, self.ws.max_column + 1):
            v = self.ws.cell(1, c).value
            if v is not None:
                try:
                    number = int(v)
                except (ValueError, TypeError):
                    continue
                if number in self.columns:
                    raise ValueError(f"중복 Col 번호: {number}")
                self.columns[number] = c
        if not {1, 2, 3} <= self.columns.keys():
            raise ValueError("DATA 첫 행에 Col1/2/3 번호가 필요합니다.")
        self.refresh()
        self.protected_snapshot = self.history_snapshot()

    def history_snapshot(self):
        return {f"{p}:{short}": {"row": row, "values": {str(c): self.ws.cell(row, ec).value for c, ec in self.columns.items()}}
                for (p, short), row in self.rows.items() if p.isdigit() and int(p) <= int(self.protect_through)}

    def assert_history_unchanged(self, snapshot=None):
        self.refresh()
        def same(a, b):
            if isinstance(a, dict) and isinstance(b, dict):
                return a.keys() == b.keys() and all(same(a[k], b[k]) for k in a)
            if a == b:
                return True
            # XLSX serialization rounds binary floats to decimal text. Accept
            # only a few machine ULPs, not a business/percentage tolerance.
            if isinstance(a, float) and isinstance(b, float) and math.isfinite(a) and math.isfinite(b):
                return abs(a-b) <= 4 * max(math.ulp(a), math.ulp(b))
            return False
        if not same(self.history_snapshot(), self.protected_snapshot if snapshot is None else snapshot):
            raise ValueError(f"{self.protect_through}까지의 과거 값·수식·행 위치가 변경됐습니다.")

    def refresh(self):
        self.rows = {}
        for row in range(6, self.ws.max_row + 1):
            p = self.ws.cell(row, self.columns[2]).value
            s = resolve(self.ws, row, self.columns[3])
            if p is None or not s:
                continue
            key = (str(p), s)
            if key in self.rows:
                raise ValueError(f"회사·기간 중복: {key}")
            self.rows[key] = row

    def ensure_period(self, period: Period, companies):
        """Use the spacer after each company's block. Never shift historical rows."""
        for _, short in companies:
            key = (period.code, short)
            if key in self.rows:
                continue
            candidates = [(int(p), r) for (p, s), r in self.rows.items()
                          if s == short and p.isdigit() and int(p) < int(period.code)]
            if not candidates:
                raise ValueError(f"{short}: 이전 분기 템플릿 행이 없습니다.")
            _, source = max(candidates)
            target = source + 1
            # Occupied spacer: append instead of inserting and damaging other-sheet references.
            if any(self.ws.cell(target, c).value is not None for c in self.columns.values()):
                target = max(self.rows.values()) + 2
                while any(self.ws.cell(target, c).value is not None for c in self.columns.values()):
                    target += 1
            for col, ec in self.columns.items():
                src, dst = self.ws.cell(source, ec), self.ws.cell(target, ec)
                dst._style = copy(src._style)
                dst.alignment = copy(src.alignment)
                dst.protection = copy(src.protection)
                if src.data_type == "f" and col > 3:
                    dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
                else:
                    dst.value = None
            self.ws.row_dimensions[target].height = self.ws.row_dimensions[source].height
            self.ws.cell(target, self.columns[1]).value = period.code + short
            self.ws.cell(target, self.columns[2]).value = int(period.code)
            self.ws.cell(target, self.columns[3]).value = short
            self.refresh()

    def fields(self):
        # Read each level with its own merged-cell anchor, not arbitrary left text.
        def text(row, ec):
            for merged in self.ws.merged_cells.ranges:
                if merged.min_row <= row <= merged.max_row and merged.min_col <= ec <= merged.max_col:
                    return self.ws.cell(merged.min_row, merged.min_col).value
            return self.ws.cell(row, ec).value
        out = {}
        for col, ec in self.columns.items():
            labels = list(dict.fromkeys(str(x).strip().replace("\n", " ")
                                        for r in (3, 4, 5) if (x := text(r, ec)) is not None))
            if col > 3 and labels:
                out[col] = " > ".join(labels)
        return out

    def values(self, period, short):
        row = self.rows[(str(period), short)]
        return {c: self.ws.cell(row, e).value for c, e in self.columns.items() if c > 3}

    def apply(self, period, short, col, value, color="D9EAF7", expected=...):
        if str(period).isdigit() and int(period) <= int(self.protect_through):
            raise ValueError(f"과거 정답 기간 보호: {period}. 정답본 대조 기능을 사용하세요.")
        if col <= 3 or col not in self.columns or not numeric(value):
            raise ValueError(f"입력 불가: Col{col}={value}")
        cell = self.ws.cell(self.rows[(str(period), short)], self.columns[col])
        if expected is not ... and cell.value != expected:
            raise ValueError(f"기존값 변경 충돌: {short} Col{col}")
        if cell.data_type == "f":
            raise ValueError(f"수식 보호: {short} Col{col} ({cell.coordinate})")
        old = cell.value
        cell.value = value
        cell.fill = PatternFill("solid", fgColor=color)
        return {"period": str(period), "company": short, "field": col,
                "cell": cell.coordinate, "old": old, "new": value}

    def save(self, path):
        self.assert_history_unchanged()
        self.wb.calculation = CalcProperties(calcId=0, fullCalcOnLoad=True, forceFullCalc=True)
        save_workbook(self.wb, Path(path))


def validate(book: DataBook, period, companies):
    results = []
    rules = [
        ("자산=부채+자본", (4, 12, 15), lambda d: d[4] - d[12] - d[15]),
        ("CSM 변동", (75, 76, 77, 78, 79, 80), lambda d: d[80]+d[76]-d[77]+d[78]+d[79]-d[75]),
        ("예실차 합계", (111, 112, 115), lambda d: d[111]-d[112]-d[115]),
        ("보험금 예실차", (112, 113, 114), lambda d: d[112]-d[113]+d[114]),
        ("사업비 예실차", (115, 116, 117), lambda d: d[115]-d[116]+d[117]),
        ("K-ICS 잠정", (92, 93, 94), lambda d: d[93]-d[92]*d[94]),
        ("K-ICS 확정", (95, 96, 97), lambda d: d[96]-d[95]*d[97]),
    ]
    for _, short in companies:
        d = book.values(period, short)
        for name, cols, calc in rules:
            if not all(numeric(d.get(c)) for c in cols):
                results.append({"company": short, "check": name, "fields": list(cols), "status": "not_computable"})
                continue
            error = calc(d)
            # Existing DATA is in million KRW; this is a screening tolerance, not proof.
            tolerance = max(1, max(abs(d[c]) for c in cols) * 0.001)
            results.append({"company": short, "check": name, "fields": list(cols), "status": "pass" if abs(error) <= tolerance else "fail",
                            "difference": error, "tolerance": tolerance})
    return results
