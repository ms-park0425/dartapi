from __future__ import annotations

import calendar
import hashlib
import json
import math
import os
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo


@dataclass(frozen=True)
class Period:
    year: int
    quarter: int

    def __post_init__(self):
        if not 2000 <= self.year <= 2099 or self.quarter not in range(1, 5):
            raise ValueError("연도는 2000~2099, 분기는 1~4여야 합니다.")

    @classmethod
    def parse(cls, code: str):
        if len(str(code)) != 4 or str(code)[2:] not in ("03", "06", "09", "12"):
            raise ValueError("기간코드는 YYMM 형식입니다. 예: 2609")
        return cls(2000 + int(str(code)[:2]), int(str(code)[2:]) // 3)

    @property
    def month(self):
        return self.quarter * 3

    @property
    def code(self):
        return f"{self.year % 100:02d}{self.month:02d}"

    @property
    def end(self):
        return date(self.year, self.month, calendar.monthrange(self.year, self.month)[1])

    @property
    def report_type(self):
        return {1: "1분기보고서", 2: "반기보고서", 3: "3분기보고서", 4: "사업보고서"}[self.quarter]

    @property
    def report_code(self):
        return {1: "11013", 2: "11012", 3: "11014", 4: "11011"}[self.quarter]

    def matches_report(self, title: str):
        import re
        label = "사업보고서" if self.quarter == 4 else "반기보고서" if self.quarter == 2 else "분기보고서"
        return label in title and bool(re.search(rf"\(\s*{self.year}\s*\.\s*{self.month:02d}\s*\)", title))


def korea_now():
    # Windows machines do not always include the IANA database.
    try:
        return datetime.now(ZoneInfo("Asia/Seoul"))
    except Exception:
        from datetime import timezone, timedelta
        return datetime.now(timezone(timedelta(hours=9)))


def numeric(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def digest(path: Path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def dump(path: Path, data: Any):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
    for attempt in range(8):
        try:
            os.replace(tmp, path)
            break
        except PermissionError:
            if attempt == 7:
                raise
            time.sleep(0.05 * (attempt + 1))


def load(path: Path, default=None):
    return json.loads(Path(path).read_text(encoding="utf-8")) if Path(path).exists() else default


def save_workbook(wb, path: Path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.stem + ".tmp" + path.suffix)
    wb.save(tmp)
    os.replace(tmp, path)


COMPANIES = [
    ("미래에셋생명", "미래"), ("삼성생명", "삼성"), ("교보생명", "교보"),
    ("한화생명", "한화"), ("신한라이프생명", "신한라이프"), ("동양생명", "동양"),
    ("KB생명", "KB라이프"), ("삼성화재", "삼성화재"), ("현대해상", "현대해상"),
    ("메리츠화재", "메리츠화재"), ("KB손해보험", "KB손보"), ("DB손해보험", "DB손보"),
]
