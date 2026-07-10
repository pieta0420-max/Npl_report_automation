"""tkinter wizard: pick Data Disk file(s) -> confirm sheet-type detection ->
confirm column mapping -> generate the report.

Every step after file selection shows the auto-detected result and lets the
user override it before anything is written -- per spec, auto-detection is
always followed by a human confirmation step.
"""
from __future__ import annotations

import os
import tkinter as tk
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from aggregator import compute_all, load_config
from column_matcher import ColumnMatch, get_header_texts, match_columns
from excel_io import load_workbook_any
from normalizer import ConfirmedSheet, normalize_all
from schema import CANONICAL_SCHEMAS
from sheet_matcher import SheetMatch, detect_sheet_type, find_header_row
from word_writer import build_word_report
from writer import write_report

TYPE_LABELS = {
    "borrower": "차주 (Sheet A)",
    "collateral": "담보 (Sheet C-1)",
    "guarantee": "신용보증서 (Sheet E)",
    "rehab": "회생차주 (Sheet F)",
}
UNUSED_LABEL = "사용 안 함"
TYPE_CHOICES = [UNUSED_LABEL] + list(TYPE_LABELS.values())
LABEL_TO_TYPE = {v: k for k, v in TYPE_LABELS.items()}


class ScrollableFrame(ttk.Frame):
    """A frame with a vertical scrollbar; put widgets in .inner."""

    def __init__(self, parent, height=400):
        super().__init__(parent)
        canvas = tk.Canvas(self, height=height, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)


@dataclass
class SheetRow:
    """One (file, sheet) pair being confirmed in the wizard."""
    file_label: str
    sheet_name: str
    grid: list
    header_row_index: Optional[int]
    detected_type: Optional[str]
    confidence: float
    type_var: tk.StringVar = field(default=None)
    column_matches: List[ColumnMatch] = field(default_factory=list)
    column_vars: List[tk.StringVar] = field(default_factory=list)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NPL Data Disk → 투심보고서 자동화")
        self.geometry("1000x700")

        self.file_labels: List[str] = []
        self.workbooks: Dict[str, dict] = {}
        self.sheet_rows: List[SheetRow] = []

        self.container = ttk.Frame(self)
        self.container.pack(fill="both", expand=True, padx=12, pady=12)

        self.show_file_select()

    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    # ------------------------------------------------------------------
    # Step 1: file selection
    # ------------------------------------------------------------------
    def show_file_select(self):
        self._clear()
        ttk.Label(self.container, text="1단계: Data Disk 파일 선택", font=("", 14, "bold")).pack(anchor="w")
        ttk.Label(self.container, text="은행/매각자문사별 Data Disk 엑셀 파일을 하나 이상 선택하세요.").pack(anchor="w", pady=(0, 10))

        self.file_listbox = tk.Listbox(self.container, height=10)
        self.file_listbox.pack(fill="both", expand=True, pady=(0, 10))

        btn_row = ttk.Frame(self.container)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="파일 추가...", command=self._add_files).pack(side="left")
        ttk.Button(btn_row, text="선택 제거", command=self._remove_selected_file).pack(side="left", padx=6)
        ttk.Button(btn_row, text="다음 →", command=self._go_to_sheet_mapping).pack(side="right")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Data Disk 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        for p in paths:
            if p not in self.file_labels:
                self.file_labels.append(p)
                self.file_listbox.insert("end", p)

    def _remove_selected_file(self):
        sel = list(self.file_listbox.curselection())
        for idx in reversed(sel):
            del self.file_labels[idx]
            self.file_listbox.delete(idx)

    def _go_to_sheet_mapping(self):
        if not self.file_labels:
            messagebox.showwarning("파일 없음", "Data Disk 파일을 최소 1개 이상 선택하세요.")
            return

        self.sheet_rows = []
        try:
            for path in self.file_labels:
                label = Path(path).stem
                wb = load_workbook_any(path)
                self.workbooks[label] = wb
                for sheet_name, grid in wb.items():
                    if not grid:
                        continue
                    header_row = find_header_row(grid)
                    match = detect_sheet_type(sheet_name, grid)
                    self.sheet_rows.append(SheetRow(
                        file_label=label, sheet_name=sheet_name, grid=grid,
                        header_row_index=match.header_row_index,
                        detected_type=match.type_id, confidence=match.confidence,
                    ))
        except Exception as e:
            messagebox.showerror("파일 읽기 오류", f"{e}\n\n{traceback.format_exc()}")
            return

        self.show_sheet_mapping()

    # ------------------------------------------------------------------
    # Step 2: sheet-type confirmation
    # ------------------------------------------------------------------
    def show_sheet_mapping(self):
        self._clear()
        ttk.Label(self.container, text="2단계: 시트 판별 확인", font=("", 14, "bold")).pack(anchor="w")
        ttk.Label(self.container, text="자동 판별 결과를 확인하고 필요하면 수정하세요. (같은 유형이 여러 시트에 잡힐 수 있으니 중복 여부를 확인하세요.)").pack(anchor="w", pady=(0, 10))

        scroll = ScrollableFrame(self.container, height=480)
        scroll.pack(fill="both", expand=True)

        header = ttk.Frame(scroll.inner)
        header.pack(fill="x", pady=2)
        for text, width in [("파일", 22), ("시트명", 16), ("자동판별", 10), ("신뢰도", 8), ("확정", 22)]:
            ttk.Label(header, text=text, width=width, font=("", 9, "bold")).pack(side="left")

        for row in self.sheet_rows:
            line = ttk.Frame(scroll.inner)
            line.pack(fill="x", pady=1)
            ttk.Label(line, text=row.file_label, width=22).pack(side="left")
            ttk.Label(line, text=row.sheet_name, width=16).pack(side="left")
            ttk.Label(line, text=TYPE_LABELS.get(row.detected_type, "-"), width=10).pack(side="left")
            ttk.Label(line, text=f"{row.confidence:.2f}", width=8).pack(side="left")

            default_label = TYPE_LABELS.get(row.detected_type, UNUSED_LABEL)
            row.type_var = tk.StringVar(value=default_label)
            combo = ttk.Combobox(line, textvariable=row.type_var, values=TYPE_CHOICES, width=20, state="readonly")
            combo.pack(side="left")

        btn_row = ttk.Frame(self.container)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="← 이전", command=self.show_file_select).pack(side="left")
        ttk.Button(btn_row, text="다음 →", command=self._go_to_column_mapping).pack(side="right")

    def _go_to_column_mapping(self):
        confirmed_rows = [r for r in self.sheet_rows if LABEL_TO_TYPE.get(r.type_var.get())]
        if not confirmed_rows:
            messagebox.showwarning("선택 없음", "최소 하나의 시트에 유형을 지정하세요.")
            return

        for row in confirmed_rows:
            type_id = LABEL_TO_TYPE[row.type_var.get()]
            schema = CANONICAL_SCHEMAS[type_id]
            header_row_index = row.header_row_index if row.header_row_index is not None else 0
            header_texts = get_header_texts(row.grid, header_row_index)
            row.column_matches = match_columns(header_texts, schema)

        self.confirmed_rows = confirmed_rows
        self.show_column_mapping()

    # ------------------------------------------------------------------
    # Step 3: column mapping confirmation
    # ------------------------------------------------------------------
    def show_column_mapping(self):
        self._clear()
        ttk.Label(self.container, text="3단계: 컴럼 매핑 확인", font=("", 14, "bold")).pack(anchor="w")
        ttk.Label(self.container, text="소스 컴럼이 매핑된 표준 항목을 확인하고 필요하면 수정하세요.").pack(anchor="w", pady=(0, 10))

        notebook = ttk.Notebook(self.container)
        notebook.pack(fill="both", expand=True)

        for row in self.confirmed_rows:
            type_id = LABEL_TO_TYPE[row.type_var.get()]
            schema = CANONICAL_SCHEMAS[type_id]
            target_choices = [UNUSED_LABEL] + [c.key for c in schema.columns if not c.derived]

            tab = ScrollableFrame(notebook, height=450)
            notebook.add(tab, text=f"{row.file_label}/{row.sheet_name}")

            head = ttk.Frame(tab.inner)
            head.pack(fill="x", pady=2)
            for text, width in [("원본 헤더", 45), ("신뢰도", 8), ("매핑된 표준 항목", 26)]:
                ttk.Label(head, text=text, width=width, font=("", 9, "bold")).pack(side="left")

            row.column_vars = []
            for cm in row.column_matches:
                line = ttk.Frame(tab.inner)
                line.pack(fill="x", pady=1)
                ttk.Label(line, text=cm.source_header[:60], width=45).pack(side="left")
                ttk.Label(line, text=f"{cm.confidence:.2f}", width=8).pack(side="left")

                var = tk.StringVar(value=cm.matched_key if cm.matched_key else UNUSED_LABEL)
                combo = ttk.Combobox(line, textvariable=var, values=target_choices, width=24, state="readonly")
                combo.pack(side="left")
                row.column_vars.append(var)

        btn_row = ttk.Frame(self.container)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="← 이전", command=self.show_sheet_mapping).pack(side="left")
        ttk.Button(btn_row, text="다음 →", command=self.show_report_info).pack(side="right")

    # ------------------------------------------------------------------
    # Step 4: report metadata (bank/program name, date, optional Word template)
    # ------------------------------------------------------------------
    def show_report_info(self):
        self._clear()
        ttk.Label(self.container, text="4단계: 보고서 정보", font=("", 14, "bold")).pack(anchor="w")
        ttk.Label(self.container, text="DD 데이터만으로는 알 수 없는 항목입니다. 엑셀 보고서 생성에는 Program명만 쓰이고, 워드 보고서를 함께 만들려면 전부 입력하세요.").pack(anchor="w", pady=(0, 10))

        form = ttk.Frame(self.container)
        form.pack(fill="x", pady=6)

        default_program = " / ".join(sorted(set(r.file_label for r in self.confirmed_rows)))
        self.bank_name_var = tk.StringVar(value="")
        self.program_name_var = tk.StringVar(value=default_program)
        self.report_date_var = tk.StringVar(value=datetime.now().strftime("%Y. %m. %d"))
        self.word_template_var = tk.StringVar(value="")

        for label, var, width in [
            ("매각은행명 (예: 우리은행)", self.bank_name_var, 40),
            ("Program명 (예: WRB 2026-2 Program)", self.program_name_var, 40),
            ("보고서 날짜 (예: 2026. 07. 10)", self.report_date_var, 40),
        ]:
            row = ttk.Frame(form)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=label, width=32).pack(side="left")
            ttk.Entry(row, textvariable=var, width=width).pack(side="left")

        row = ttk.Frame(form)
        row.pack(fill="x", pady=(12, 4))
        ttk.Label(row, text="워드 보고서 템플릿 (.docx, 선택)", width=32).pack(side="left")
        ttk.Entry(row, textvariable=self.word_template_var, width=50).pack(side="left")
        ttk.Button(row, text="찾아보기...", command=self._pick_word_template).pack(side="left", padx=6)
        ttk.Label(self.container, text="템플릿을 지정하면 원본 파일을 복사해 1~9페이지(자산 소개) 표/문장을 채운 워드 파일도 함께 생성합니다. 비워두면 엑셀만 생성합니다.",
                  foreground="#666666", wraplength=900).pack(anchor="w", pady=(2, 0))

        btn_row = ttk.Frame(self.container)
        btn_row.pack(fill="x", pady=(20, 0))
        ttk.Button(btn_row, text="← 이전", command=self.show_column_mapping).pack(side="left")
        ttk.Button(btn_row, text="보고서 생성 →", command=self._generate).pack(side="right")

    def _pick_word_template(self):
        path = filedialog.askopenfilename(
            title="워드 보고서 템플릿 선택",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.word_template_var.set(path)

    # ------------------------------------------------------------------
    # Step 5: generate
    # ------------------------------------------------------------------
    def _generate(self):
        for row in self.confirmed_rows:
            for cm, var in zip(row.column_matches, row.column_vars):
                chosen = var.get()
                cm.matched_key = None if chosen == UNUSED_LABEL else chosen

        default_name = f"투심자료_{datetime.now().strftime('%Y%m%d')}.xlsx"
        out_path = filedialog.asksaveasfilename(
            title="엑셀 저장 파일 이름",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not out_path:
            return

        word_template = self.word_template_var.get().strip()
        word_out_path = None
        if word_template:
            default_word_name = f"투심보고서_{datetime.now().strftime('%Y%m%d')}.docx"
            word_out_path = filedialog.asksaveasfilename(
                title="워드 저장 파일 이름",
                defaultextension=".docx",
                initialfile=default_word_name,
                filetypes=[("Word files", "*.docx")],
            )
            if not word_out_path:
                return

        try:
            confirmed_sheets = [
                ConfirmedSheet(
                    file_label=row.file_label, sheet_name=row.sheet_name,
                    type_id=LABEL_TO_TYPE[row.type_var.get()],
                    header_row_index=row.header_row_index if row.header_row_index is not None else 0,
                    column_matches=row.column_matches, grid=row.grid,
                )
                for row in self.confirmed_rows
            ]
            frames = normalize_all(confirmed_sheets)
            config = load_config()
            agg_result = compute_all(frames, config)
            program_name = self.program_name_var.get().strip() or "NPL Data Disk 정리 결과"
            write_report(frames, agg_result, config, out_path, program_name=program_name)

            if word_template and word_out_path:
                build_word_report(
                    word_template, word_out_path, agg_result, frames, config,
                    bank_name=self.bank_name_var.get().strip() or "매각은행",
                    program_name=program_name,
                    report_date=self.report_date_var.get().strip() or datetime.now().strftime("%Y. %m. %d"),
                )
        except Exception as e:
            messagebox.showerror("생성 오류", f"{e}\n\n{traceback.format_exc()}")
            return

        self.show_done(out_path, word_out_path, agg_result)

    def show_done(self, out_path: str, word_out_path: Optional[str], agg_result):
        self._clear()
        ttk.Label(self.container, text="완료!", font=("", 16, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(self.container, text=f"생성된 엑셀 파일: {out_path}").pack(anchor="w")
        if word_out_path:
            ttk.Label(self.container, text=f"생성된 워드 파일: {word_out_path}").pack(anchor="w")

        unclassified = agg_result.unclassified_property_types | agg_result.unclassified_regions
        if unclassified:
            ttk.Label(self.container, text="").pack(pady=4)
            ttk.Label(
                self.container,
                text="⚠ classification_config.json에 없어 '기타'로 처리된 항목이 있습니다. 필요시 설정 파일을 보강하세요.",
                foreground="#B22222",
            ).pack(anchor="w")
            ttk.Label(self.container, text=", ".join(sorted(unclassified)), wraplength=900).pack(anchor="w")

        btn_row = ttk.Frame(self.container)
        btn_row.pack(fill="x", pady=(20, 0))
        ttk.Button(btn_row, text="폴더 열기", command=lambda: os.startfile(os.path.dirname(out_path))).pack(side="left")
        ttk.Button(btn_row, text="종료", command=self.destroy).pack(side="right")
        ttk.Button(btn_row, text="처음부터 다시", command=self.show_file_select).pack(side="right", padx=6)


def run():
    app = App()
    app.mainloop()
