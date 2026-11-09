"""Small desktop selector for the complete quarterly workflow."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, ttk

ROOT = Path(__file__).resolve().parent


def main():
    window = tk.Tk()
    window.title("보험사 분기 공시 입력")
    window.geometry("780x540")
    window.minsize(600, 400)
    panel = ttk.Frame(window, padding=18)
    panel.pack(fill="both", expand=True)
    ttk.Label(panel, text="분기를 선택하면 자료 수집부터 PDF 검토·엑셀 수정까지 실행합니다.").pack(anchor="w", pady=(0, 16))
    controls = ttk.Frame(panel)
    controls.pack(fill="x")
    year, quarter = tk.StringVar(value=str(date.today().year)), tk.StringVar(value="3")
    ttk.Label(controls, text="연도").pack(side="left")
    ttk.Combobox(controls, textvariable=year, values=list(range(2023, date.today().year + 2)), width=8, state="readonly").pack(side="left", padx=(8, 18))
    ttk.Label(controls, text="분기").pack(side="left")
    ttk.Combobox(controls, textvariable=quarter, values=[1, 2, 3, 4], width=5, state="readonly").pack(side="left", padx=8)
    from quarterly.core import load, dump
    config_path = ROOT / "quarterly.local.json"
    from quarter_pipeline import read_config
    config = read_config(ROOT / "quarterly.json")
    local_config = load(config_path, {})
    workbook = tk.StringVar(value=str((ROOT / config["workbook"]).resolve()))
    ttk.Label(panel, text="기준 작업 엑셀 (최종 결과는 실행 폴더에 별도 저장)").pack(anchor="w", pady=(18, 4))
    path_row = ttk.Frame(panel)
    path_row.pack(fill="x")
    ttk.Entry(path_row, textvariable=workbook).pack(side="left", fill="x", expand=True)

    def browse():
        chosen = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if chosen:
            workbook.set(chosen)
    ttk.Button(path_row, text="찾기", command=browse).pack(side="left", padx=(8, 0))
    status = tk.StringVar(value="대기")
    ttk.Label(panel, textvariable=status).pack(anchor="w", pady=(16, 4))
    output = tk.Text(panel, height=16, wrap="word", state="disabled")
    output.pack(fill="both", expand=True)
    events = queue.Queue()

    def worker(command):
        try:
            p = subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                 text=True, encoding="utf-8", errors="replace", bufsize=1)
            for line in p.stdout:
                events.put(("line", line))
            events.put(("done", p.wait()))
        except Exception as exc:
            events.put(("line", str(exc) + "\n"))
            events.put(("done", 1))

    def start():
        local_config["workbook"] = workbook.get()
        dump(config_path, local_config)
        button.config(state="disabled")
        status.set("자료 수집 → 1차 입력 → PDF 직접 검토 → 2차 수정 진행 중")
        command = [sys.executable, "-u", str(ROOT / "quarter_pipeline.py"), "--year", year.get(), "--quarter", quarter.get(), "--workbook", workbook.get()]
        threading.Thread(target=worker, args=(command,), daemon=True).start()

    button = ttk.Button(panel, text="선택한 분기 실행", command=start)
    button.pack(anchor="e", pady=(12, 0))

    def poll():
        while not events.empty():
            kind, value = events.get()
            if kind == "line":
                output.config(state="normal")
                output.insert("end", value)
                output.see("end")
                output.config(state="disabled")
            else:
                status.set("실행 종료 — 아래 저장 경로와 미해결 항목 수를 확인하세요." if value == 0 else "실행 중단 — 아래 오류와 실행 폴더를 확인하세요.")
                button.config(state="normal")
        window.after(200, poll)
    poll()
    window.mainloop()


if __name__ == "__main__":
    main()
