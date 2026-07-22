from __future__ import annotations

from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook

APP_TITLE = "원우ENG CWT 생산계획 수립 프로그램"
APP_VERSION = "20260722_2135"
PLAN_COLUMNS = ["착수일", "라인", "생산번호(물류번호)", "자재번호", "품명", "모델명", "차대호기", "순번", "연번", "CWT", "RADAR", "설계모델", "소요량", "발주량", "발주번호", "납기일자", "비고"]
EXCEPTION_COLUMNS = ["예외유형", "라인", "생산번호/물류번호", "영업모델", "자재번호", "사유"]


def clean(value) -> str:
    return str(value or "").replace(" ", "").strip()


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def search_text(value) -> str:
    return str(value or "").replace(" ", "").upper()


def date_text(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime("%Y-%m-%d")
        except Exception:
            pass
    raw = text(value)
    if not raw:
        return ""
    for sep in ("-", ".", "/"):
        parts = raw.split(sep)
        if len(parts) >= 3 and len(parts[0]) == 4:
            return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2][:2].zfill(2)}"
    return raw


def safe_int(value) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return 999999


def score_header(values, required) -> int:
    cells = {clean(value).upper() for value in values}
    return sum(1 for header in required if clean(header).upper() in cells)


def detect_header_row(ws, preferred: int, required: list[str]) -> int:
    best_row = preferred
    best_score = score_header([cell.value for cell in ws[preferred]], required)
    for row_no in range(1, min(ws.max_row, 12) + 1):
        score = score_header([cell.value for cell in ws[row_no]], required)
        if score > best_score:
            best_row, best_score = row_no, score
    return best_row


def read_rows(path: Path, preferred_header_row: int, required: list[str]) -> tuple[list[str], list[dict], int]:
    if not path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(".xlsx 또는 .xlsm 파일만 지원합니다. .xls 파일은 xlsx로 저장 후 다시 선택하세요.")
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_row = detect_header_row(ws, preferred_header_row, required)
    headers = [clean(cell.value) for cell in ws[header_row]]
    rows = []
    for row_no, values in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
        values = list(values)
        if any(text(value) for value in values):
            rows.append({"row_number": row_no, "values": values})
    wb.close()
    return headers, rows, header_row


def col(headers: list[str], names: list[str], *, last: bool = False) -> int | None:
    targets = {clean(name).upper() for name in names}
    matches = [i for i, header in enumerate(headers) if clean(header).upper() in targets]
    if not matches:
        return None
    return matches[-1] if last else matches[0]


def val(row: dict, index: int | None) -> str:
    if index is None or index >= len(row["values"]):
        return ""
    return text(row["values"][index])


def dval(row: dict, index: int | None) -> str:
    if index is None or index >= len(row["values"]):
        return ""
    return date_text(row["values"][index])


def exception(kind, line, key, model, material_no, reason) -> dict:
    return {"예외유형": kind, "라인": line, "생산번호/물류번호": key, "영업모델": model, "자재번호": material_no, "사유": reason}


def monthly_map(headers: list[str]) -> dict:
    return {
        "production_no": col(headers, ["생산번호", "물류번호"]),
        "sales_model": col(headers, ["영업모델"]),
        "machine_no": col(headers, ["차대호기"]),
        "seq": col(headers, ["순번"], last=True),
        "serial": col(headers, ["연번"]),
        "line": col(headers, ["라인"]),
        "start_date": col(headers, ["착수일", "착수일자"]),
        "radar": col(headers, ["RADAR", "Radar"]),
        "cwt": col(headers, ["CWT"]),
        "design_model": col(headers, ["설계모델"]),
        "remark": col(headers, ["비고"]),
    }


def material_map(headers: list[str]) -> dict:
    mapping = {
        "logistics_no": col(headers, ["물류번호", "생산번호"]),
        "material_no": col(headers, ["자재번호"]),
        "material_name": col(headers, ["품명"]),
        "required_qty": col(headers, ["소요량"]),
        "order_qty": col(headers, ["발주량"]),
        "po_no": col(headers, ["발주번호"]),
        "due_date": col(headers, ["납기일자", "납품예정일"]),
    }
    if len(headers) >= 24:
        fallback = {"logistics_no": 3, "material_no": 4, "material_name": 5, "required_qty": 8, "order_qty": 9, "po_no": 20, "due_date": 23}
        for key, index in fallback.items():
            if mapping[key] is None:
                mapping[key] = index
    return mapping


def parse_monthly(path: Path, line_type: str, filter_hyundai: bool, exceptions: list[dict]) -> tuple[list[dict], int, int]:
    headers, rows, _ = read_rows(path, 2, ["생산번호", "영업모델", "착수일"])
    c = monthly_map(headers)
    required = {"production_no": "생산번호", "sales_model": "영업모델", "machine_no": "차대호기", "seq": "순번", "serial": "연번", "line": "라인", "start_date": "착수일", "cwt": "CWT", "design_model": "설계모델", "remark": "비고"}
    missing = [name for key, name in required.items() if c.get(key) is None]
    if missing:
        exceptions.append(exception("필수 컬럼 누락", line_type, "", "", "", f"{path.name}: {', '.join(missing)} 컬럼을 찾을 수 없습니다."))
        return [], len(rows), 0
    parsed, dx_count = [], 0
    for row in rows:
        item = {
            "production_no": clean(val(row, c["production_no"])),
            "sales_model": val(row, c["sales_model"]),
            "machine_no": val(row, c["machine_no"]),
            "seq": val(row, c["seq"]),
            "serial": val(row, c["serial"]),
            "line": val(row, c["line"]),
            "line_type": line_type,
            "start_date": dval(row, c["start_date"]),
            "radar": val(row, c["radar"]),
            "cwt": val(row, c["cwt"]),
            "design_model": val(row, c["design_model"]),
            "remark": val(row, c["remark"]),
        }
        if not item["production_no"]:
            continue
        model = item["sales_model"].strip()
        if filter_hyundai and model.startswith("DX"):
            dx_count += 1
            exceptions.append(exception("DX 제품 제외", line_type, item["production_no"], model, "", "통합2라인 디벨론 제품은 생산계획 I/F 대상에서 제외합니다."))
        elif filter_hyundai and not model.startswith(("HX", "R")):
            exceptions.append(exception("모델 필터 제외", line_type, item["production_no"], model, "", "통합2라인 현대 제품 기준 HX~/R~에 해당하지 않습니다."))
        else:
            parsed.append(item)
    return parsed, len(rows), dx_count


def line_from_logistics(logistics_no: str) -> str:
    if logistics_no.startswith("KPA10"):
        return "통합1라인"
    if logistics_no.startswith("KPA20"):
        return "통합2라인"
    return "라인 미분류"


def parse_material(path: Path, exceptions: list[dict]) -> tuple[list[dict], int]:
    headers, rows, _ = read_rows(path, 1, ["물류번호", "자재번호", "품명"])
    c = material_map(headers)
    required = {"logistics_no": "물류번호", "material_no": "자재번호", "material_name": "품명", "required_qty": "소요량", "order_qty": "발주량", "po_no": "발주번호", "due_date": "납기일자"}
    missing = [name for key, name in required.items() if c.get(key) is None]
    if missing:
        exceptions.append(exception("필수 컬럼 누락", "물류번호별 자재소요현황", "", "", "", f"{path.name}: {', '.join(missing)} 컬럼을 찾을 수 없습니다."))
    parsed = []
    for row in rows:
        logistics_no = clean(val(row, c["logistics_no"]))
        if not logistics_no:
            continue
        parsed.append({
            "logistics_no": logistics_no,
            "material_no": val(row, c["material_no"]),
            "material_name": val(row, c["material_name"]),
            "required_qty": val(row, c["required_qty"]),
            "order_qty": val(row, c["order_qty"]),
            "po_no": val(row, c["po_no"]),
            "due_date": dval(row, c["due_date"]),
            "line_type": line_from_logistics(logistics_no),
        })
    return parsed, len(rows)


def plan_row(monthly, material) -> dict:
    return {
        "착수일": monthly["start_date"], "라인": monthly["line_type"], "생산번호(물류번호)": monthly["production_no"],
        "자재번호": material["material_no"], "품명": material["material_name"], "모델명": monthly["sales_model"],
        "차대호기": monthly["machine_no"], "순번": monthly["seq"], "연번": monthly["serial"], "CWT": monthly["cwt"],
        "RADAR": monthly["radar"], "설계모델": monthly["design_model"], "소요량": material["required_qty"],
        "발주량": material["order_qty"], "발주번호": material["po_no"], "납기일자": material["due_date"], "비고": monthly["remark"],
    }


def next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 1
    while True:
        candidate = parent / f"{stem}_{index:03d}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def write_sheet(wb, title, rows, columns=None):
    ws = wb.create_sheet(title)
    if isinstance(rows[0], dict):
        headers = columns or list(rows[0].keys())
        ws.append(headers)
        for row in rows:
            ws.append([row.get(header, "") for header in headers])
    else:
        for row in rows:
            ws.append(row)
    ws.freeze_panes = "A2"
    for cells in ws.columns:
        ws.column_dimensions[cells[0].column_letter].width = min(max(len(str(cell.value or "")) for cell in cells) + 2, 44)


def process(line1_path: Path | None, line2_path: Path | None, material_path: Path, output_dir: Path) -> Path:
    if not line1_path and not line2_path:
        raise ValueError("월확정 통합1라인 또는 통합2라인 파일 중 하나 이상을 선택하세요.")
    exceptions = []
    line1, line1_total, _ = parse_monthly(line1_path, "통합1라인", False, exceptions) if line1_path else ([], 0, 0)
    line2, line2_total, dx_count = parse_monthly(line2_path, "통합2라인", True, exceptions) if line2_path else ([], 0, 0)
    material_all, material_total = parse_material(material_path, exceptions)
    material_line1 = [m for m in material_all if m["logistics_no"].startswith("KPA10") and "카운터웨이트" in search_text(m["material_name"])]
    material_line2 = [m for m in material_all if m["logistics_no"].startswith("KPA20") and "카운터웨이트" in search_text(m["material_name"])]
    selected_materials = []
    excluded_materials = []
    if line1_path:
        selected_materials.extend(material_line1)
    else:
        excluded_materials.extend(material_line1)
    if line2_path:
        selected_materials.extend(material_line2)
    else:
        excluded_materials.extend(material_line2)
    material_index = defaultdict(list)
    all_material_index = defaultdict(list)
    for item in material_all:
        all_material_index[item["logistics_no"]].append(item)
    for item in selected_materials:
        material_index[item["logistics_no"]].append(item)
    monthly = [*line1, *line2]
    plans = []
    for item in monthly:
        matches = material_index.get(item["production_no"], [])
        if len(matches) == 1:
            plans.append(plan_row(item, matches[0]))
        elif len(matches) == 0:
            candidates = all_material_index.get(item["production_no"], [])
            if candidates:
                nos = ", ".join(m["material_no"] for m in candidates if m["material_no"])
                names = " / ".join(sorted({m["material_name"] for m in candidates if m["material_name"]}))
                exceptions.append(exception("카운터웨이트 품명 미확인", item["line_type"], item["production_no"], item["sales_model"], nos, f"동일 물류번호 자재는 있으나 품명에서 카운터웨이트를 확인하지 못했습니다. 품명: {names}"))
            else:
                exceptions.append(exception("자재소요 미매칭", item["line_type"], item["production_no"], item["sales_model"], "", "생산번호와 일치하는 카운터웨이트 자재소요가 없습니다."))
        else:
            exceptions.append(exception("자재소요 중복", item["line_type"], item["production_no"], item["sales_model"], ", ".join(m["material_no"] for m in matches), "동일 생산번호에 카운터웨이트 자재소요가 2건 이상입니다."))
    monthly_keys = {m["production_no"] for m in monthly}
    for item in selected_materials:
        if item["logistics_no"] not in monthly_keys:
            exceptions.append(exception("월확정 미매칭", item["line_type"], item["logistics_no"], "", item["material_no"], "물류번호와 일치하는 월확정 생산번호가 없습니다."))
    for item in excluded_materials:
        exceptions.append(exception("선택 라인 제외", item["line_type"], item["logistics_no"], "", item["material_no"], f"{item['line_type']} 월확정 파일이 선택되지 않아 생산계획 생성 대상에서 제외했습니다."))
    plans.sort(key=lambda r: (str(r["착수일"]), str(r["라인"]), safe_int(r["순번"]), str(r["연번"]), str(r["생산번호(물류번호)"])))
    dates = sorted(row["착수일"] for row in plans if row["착수일"])
    period = "START-END" if not dates else f"{dates[0].replace('-', '')}-{dates[-1].replace('-', '')}"
    period_label = "착수일 없음" if not dates else f"{dates[0]} ~ {dates[-1]}"
    output_path = next_available_path(output_dir / f"원우ENG_CWT_생산계획_{period}.xlsx")
    summary = [["구분", "건수", "비고"], ["월확정 통합1라인 전체", line1_total, line1_path.name if line1_path else "파일 미선택"], ["월확정 통합1라인 사용", len(line1), "선택 시 전체 사용"], ["월확정 통합2라인 전체", line2_total, line2_path.name if line2_path else "파일 미선택"], ["월확정 통합2라인 현대 제품", len(line2), "HX~/R~"], ["월확정 통합2라인 DX 제외", dx_count, "예외목록 기록"], ["자재소요 전체", material_total, material_path.name], ["자재소요 KPA10 카운터웨이트", len(material_line1), "통합1라인 매칭 대상"], ["자재소요 KPA20 카운터웨이트", len(material_line2), "통합2라인 매칭 대상"], ["선택 라인 제외 자재", len(excluded_materials), "선택되지 않은 월확정 라인의 자재"], ["정상 생산계획", len(plans), "생산계획 시트"], ["예외", len(exceptions), "예외목록 시트"], ["착수기간", period_label, "월확정 착수일 기준"], ["앱 버전", APP_VERSION, "Python + openpyxl + tkinter"]]
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "생산계획", plans or [{c: "" for c in PLAN_COLUMNS}], PLAN_COLUMNS)
    write_sheet(wb, "예외목록", exceptions or [exception("예외 없음", "", "", "", "", "예외가 없습니다.")], EXCEPTION_COLUMNS)
    write_sheet(wb, "검증요약", summary)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def open_folder(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        pass


class PlannerGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x740")
        self.minsize(1040, 660)
        self.configure(bg="#f5f7fb")
        self.files: dict[str, Path | None] = {
            "line1": None,
            "line2": None,
            "material": None,
            "output_dir": Path.home() / "Desktop",
        }
        self.last_output: Path | None = None
        self._build_styles()
        self._build_ui()
        self._update_state()

    def _build_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Root.TFrame", background="#f5f7fb")
        style.configure("Panel.TFrame", background="#ffffff", relief="solid", borderwidth=1)
        style.configure("Title.TLabel", background="#f5f7fb", foreground="#172033", font=("Malgun Gothic", 22, "bold"))
        style.configure("Sub.TLabel", background="#f5f7fb", foreground="#64748b", font=("Malgun Gothic", 10))
        style.configure("PanelTitle.TLabel", background="#ffffff", foreground="#172033", font=("Malgun Gothic", 11, "bold"))
        style.configure("Hint.TLabel", background="#ffffff", foreground="#64748b", font=("Malgun Gothic", 9))
        style.configure("Path.TLabel", background="#ffffff", foreground="#2563eb", font=("Malgun Gothic", 10, "bold"))
        style.configure("Metric.TLabel", background="#ffffff", foreground="#166534", font=("Malgun Gothic", 18, "bold"))
        style.configure("Primary.TButton", font=("Malgun Gothic", 11, "bold"), padding=(16, 8))
        style.configure("TButton", font=("Malgun Gothic", 10), padding=(12, 6))

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=24)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(root, text="월확정 통합1/2라인과 물류번호별 자재소요현황을 생산번호 기준으로 매칭해 생산계획 엑셀을 생성합니다.", style="Sub.TLabel").pack(anchor="w", pady=(6, 18))

        self.file_labels: dict[str, ttk.Label] = {}
        self._file_panel(root, "line1", "1. 월확정조립서열계획 통합1라인", "선택 사항 · 선택 시 KPA10~ 카운터웨이트 자재와 매칭")
        self._file_panel(root, "line2", "2. 월확정조립서열계획 통합2라인", "선택 사항 · HX~/R~ 현대 제품만 사용하고 DX~는 예외 처리")
        self._file_panel(root, "material", "3. 물류번호별 자재소요현황", "필수 · KPA10~/KPA20~ 및 품명 카운터웨이트 포함 건만 사용")
        self._output_panel(root)

        action = ttk.Frame(root, style="Root.TFrame")
        action.pack(fill="x", pady=(14, 12))
        self.generate_btn = ttk.Button(action, text="생산계획 생성", style="Primary.TButton", command=self.generate)
        self.generate_btn.pack(side="left")
        ttk.Button(action, text="초기화", command=self.reset).pack(side="left", padx=(8, 0))
        self.open_btn = ttk.Button(action, text="결과 폴더 열기", command=self.open_result_folder, state="disabled")
        self.open_btn.pack(side="left", padx=(8, 0))
        self.status_var = tk.StringVar(value="월확정 1개 이상과 자재소요 파일을 선택하세요.")
        ttk.Label(action, textvariable=self.status_var, style="Sub.TLabel").pack(side="right")

        metrics = ttk.Frame(root, style="Root.TFrame")
        metrics.pack(fill="x", pady=(0, 12))
        self.metric_vars = {
            "plans": tk.StringVar(value="-"),
            "exceptions": tk.StringVar(value="-"),
            "period": tk.StringVar(value="-"),
            "output": tk.StringVar(value="-"),
        }
        self._metric_panel(metrics, "정상 생산계획", self.metric_vars["plans"], 0)
        self._metric_panel(metrics, "예외", self.metric_vars["exceptions"], 1)
        self._metric_panel(metrics, "착수기간", self.metric_vars["period"], 2)
        self._metric_panel(metrics, "출력 파일", self.metric_vars["output"], 3)

        table_box = ttk.Frame(root, style="Panel.TFrame", padding=12)
        table_box.pack(fill="both", expand=True)
        ttk.Label(table_box, text="생산계획 미리보기", style="PanelTitle.TLabel").pack(anchor="w", pady=(0, 8))
        columns = ["착수일", "라인", "생산번호(물류번호)", "자재번호", "모델명", "차대호기", "CWT", "RADAR"]
        self.preview = ttk.Treeview(table_box, columns=columns, show="headings", height=12)
        for column in columns:
            self.preview.heading(column, text=column)
            self.preview.column(column, width=132, anchor="w")
        ybar = ttk.Scrollbar(table_box, orient="vertical", command=self.preview.yview)
        xbar = ttk.Scrollbar(table_box, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.preview.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        xbar.pack(side="bottom", fill="x")

    def _file_panel(self, parent: ttk.Frame, key: str, title: str, hint: str) -> None:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=14)
        panel.pack(fill="x", pady=5)
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text=title, style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(panel, text=hint, style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        label = ttk.Label(panel, text="파일을 선택하세요", style="Path.TLabel")
        label.grid(row=0, column=1, rowspan=2, sticky="ew", padx=16)
        self.file_labels[key] = label
        ttk.Button(panel, text="파일 선택", command=lambda: self.select_file(key)).grid(row=0, column=2, rowspan=2)

    def _output_panel(self, parent: ttk.Frame) -> None:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=14)
        panel.pack(fill="x", pady=5)
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text="4. 결과 저장 폴더", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(panel, text="생성된 생산계획 엑셀을 저장할 위치", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.output_label = ttk.Label(panel, text=str(self.files["output_dir"]), style="Path.TLabel")
        self.output_label.grid(row=0, column=1, rowspan=2, sticky="ew", padx=16)
        ttk.Button(panel, text="폴더 선택", command=self.select_output_dir).grid(row=0, column=2, rowspan=2)

    def _metric_panel(self, parent: ttk.Frame, title: str, var: tk.StringVar, column: int) -> None:
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=14)
        panel.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        parent.columnconfigure(column, weight=1)
        ttk.Label(panel, text=title, style="PanelTitle.TLabel").pack(anchor="w")
        ttk.Label(panel, textvariable=var, style="Metric.TLabel").pack(anchor="w", pady=(6, 0))

    def select_file(self, key: str) -> None:
        selected = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel Workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if not selected:
            return
        path = Path(selected)
        self.files[key] = path
        self.file_labels[key].configure(text=path.name)
        self._update_state()

    def select_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="결과 저장 폴더 선택")
        if not selected:
            return
        self.files["output_dir"] = Path(selected)
        self.output_label.configure(text=selected)
        self._update_state()

    def _update_state(self) -> None:
        ready = bool((self.files["line1"] or self.files["line2"]) and self.files["material"] and self.files["output_dir"])
        if hasattr(self, "generate_btn"):
            self.generate_btn.configure(state="normal" if ready else "disabled")

    def generate(self) -> None:
        try:
            self.status_var.set("엑셀 파일을 처리하고 있습니다...")
            self.update_idletasks()
            output_path = process(self.files["line1"], self.files["line2"], self.files["material"], self.files["output_dir"])
        except Exception as exc:
            self.status_var.set("처리 중 오류가 발생했습니다.")
            messagebox.showerror(APP_TITLE, str(exc))
            return

        self.last_output = output_path
        self.open_btn.configure(state="normal")
        self._load_result(output_path)
        self.status_var.set(f"완료: {output_path.name}")
        messagebox.showinfo(APP_TITLE, f"생산계획 생성 완료\n\n{output_path}")

    def _load_result(self, output_path: Path) -> None:
        wb = load_workbook(output_path, data_only=True, read_only=True)
        plan_ws = wb["생산계획"]
        exception_ws = wb["예외목록"]
        summary_ws = wb["검증요약"]
        self.metric_vars["plans"].set(f"{max(plan_ws.max_row - 1, 0)}건")
        self.metric_vars["exceptions"].set(f"{max(exception_ws.max_row - 1, 0)}건")
        period = "-"
        for row in summary_ws.iter_rows(values_only=True):
            if row and row[0] == "착수기간":
                period = str(row[1] or "-")
                break
        self.metric_vars["period"].set(period)
        self.metric_vars["output"].set(output_path.name)

        for item in self.preview.get_children():
            self.preview.delete(item)
        headers = [cell.value for cell in next(plan_ws.iter_rows(min_row=1, max_row=1))]
        header_index = {header: i for i, header in enumerate(headers)}
        columns = list(self.preview["columns"])
        for row in plan_ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            self.preview.insert("", "end", values=[row[header_index.get(column, -1)] if header_index.get(column, -1) >= 0 else "" for column in columns])
        wb.close()

    def open_result_folder(self) -> None:
        if self.last_output:
            open_folder(self.last_output.parent)

    def reset(self) -> None:
        for key in ("line1", "line2", "material"):
            self.files[key] = None
            self.file_labels[key].configure(text="파일을 선택하세요")
        self.last_output = None
        self.open_btn.configure(state="disabled")
        for var in self.metric_vars.values():
            var.set("-")
        for item in self.preview.get_children():
            self.preview.delete(item)
        self.status_var.set("월확정 1개 이상과 자재소요 파일을 선택하세요.")
        self._update_state()


def main() -> None:
    app = PlannerGui()
    app.mainloop()


if __name__ == "__main__":
    main()
