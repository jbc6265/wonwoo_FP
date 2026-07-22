from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook, load_workbook


APP_TITLE = "원우ENG CWT 생산계획 수립 프로그램"


def norm(value) -> str:
    return str(value or "").replace(" ", "").strip()


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def date_text(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    raw = text(value)
    if not raw:
        return ""
    for sep in ("-", ".", "/"):
        parts = raw.split(sep)
        if len(parts) >= 3 and len(parts[0]) == 4:
            return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2][:2].zfill(2)}"
    return raw


def header_map(ws, row_no: int) -> dict[str, list[int]]:
    found: dict[str, list[int]] = defaultdict(list)
    for index, cell in enumerate(ws[row_no], 1):
        name = norm(cell.value)
        if name:
            found[name].append(index)
    return found


def col(headers: dict[str, list[int]], names: list[str], *, last: bool = False, fallback: int | None = None) -> int | None:
    for name in names:
        values = headers.get(norm(name))
        if values:
            return values[-1] if last else values[0]
    return fallback


def read_monthly(path: Path, line_type: str, *, filter_hyundai: bool, exceptions: list[dict]) -> list[dict]:
    ws = load_workbook(path, data_only=True).active
    headers = header_map(ws, 2)
    c = {
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
    required = ["production_no", "sales_model", "machine_no", "seq", "serial", "line", "start_date", "cwt", "design_model", "remark"]
    missing = [key for key in required if c[key] is None]
    if missing:
        exceptions.append(to_exception("필수 컬럼 누락", line_type, "", "", "", f"{path.name}: {', '.join(missing)} 컬럼을 찾을 수 없습니다."))
        return []

    rows = []
    for row_no in range(3, ws.max_row + 1):
        production_no = norm(ws.cell(row_no, c["production_no"]).value)
        if not production_no:
            continue
        sales_model = text(ws.cell(row_no, c["sales_model"]).value)
        item = {
            "production_no": production_no,
            "sales_model": sales_model,
            "machine_no": text(ws.cell(row_no, c["machine_no"]).value),
            "seq": text(ws.cell(row_no, c["seq"]).value),
            "serial": text(ws.cell(row_no, c["serial"]).value),
            "line": text(ws.cell(row_no, c["line"]).value),
            "line_type": line_type,
            "start_date": date_text(ws.cell(row_no, c["start_date"]).value),
            "radar": text(ws.cell(row_no, c["radar"]).value) if c["radar"] else "",
            "cwt": text(ws.cell(row_no, c["cwt"]).value),
            "design_model": text(ws.cell(row_no, c["design_model"]).value),
            "remark": text(ws.cell(row_no, c["remark"]).value),
        }
        if filter_hyundai:
            if sales_model.startswith(("HX", "R")):
                rows.append(item)
            elif sales_model.startswith("DX"):
                exceptions.append(to_exception("DX 제품 제외", line_type, production_no, sales_model, "", "통합2라인 디벨론 제품은 생산계획 I/F 대상에서 제외합니다."))
            else:
                exceptions.append(to_exception("모델 필터 제외", line_type, production_no, sales_model, "", "통합2라인 현대 제품 기준 HX~/R~에 해당하지 않습니다."))
        else:
            rows.append(item)
    return rows


def read_material(path: Path, exceptions: list[dict]) -> list[dict]:
    ws = load_workbook(path, data_only=True).active
    headers = header_map(ws, 1)
    c = {
        "logistics_no": col(headers, ["물류번호", "생산번호"], fallback=4),
        "material_no": col(headers, ["자재번호"], fallback=5),
        "material_name": col(headers, ["품명"], fallback=6),
        "required_qty": col(headers, ["소요량"], fallback=9),
        "order_qty": col(headers, ["발주량"], fallback=10),
        "po_no": col(headers, ["발주번호"], fallback=21),
        "due_date": col(headers, ["납기일자", "납품예정일"], fallback=24),
    }
    rows = []
    for row_no in range(2, ws.max_row + 1):
        logistics_no = norm(ws.cell(row_no, c["logistics_no"]).value)
        material_name = text(ws.cell(row_no, c["material_name"]).value)
        if not logistics_no or "카운터웨이트" not in norm(material_name):
            continue
        if not logistics_no.startswith(("KPA10", "KPA20")):
            continue
        rows.append({
            "logistics_no": logistics_no,
            "material_no": text(ws.cell(row_no, c["material_no"]).value),
            "material_name": material_name,
            "required_qty": text(ws.cell(row_no, c["required_qty"]).value),
            "order_qty": text(ws.cell(row_no, c["order_qty"]).value),
            "po_no": text(ws.cell(row_no, c["po_no"]).value),
            "due_date": date_text(ws.cell(row_no, c["due_date"]).value),
            "line_type": "통합1라인" if logistics_no.startswith("KPA10") else "통합2라인",
        })
    return rows


def to_exception(kind, line, key, model, material_no, reason) -> dict:
    return {
        "예외유형": kind,
        "라인": line,
        "생산번호/물류번호": key,
        "영업모델": model,
        "자재번호": material_no,
        "사유": reason,
    }


def to_plan(monthly: dict, material: dict) -> dict:
    return {
        "착수일": monthly["start_date"],
        "라인": monthly["line_type"],
        "생산번호(물류번호)": monthly["production_no"],
        "자재번호": material["material_no"],
        "품명": material["material_name"],
        "모델명": monthly["sales_model"],
        "차대호기": monthly["machine_no"],
        "순번": monthly["seq"],
        "연번": monthly["serial"],
        "CWT": monthly["cwt"],
        "RADAR": monthly["radar"],
        "설계모델": monthly["design_model"],
        "소요량": material["required_qty"],
        "발주량": material["order_qty"],
        "발주번호": material["po_no"],
        "납기일자": material["due_date"],
        "비고": monthly["remark"],
    }


def write_sheet(wb: Workbook, title: str, rows: list[dict] | list[list]) -> None:
    ws = wb.create_sheet(title)
    if not rows:
        return
    if isinstance(rows[0], dict):
        headers = list(rows[0].keys())
        ws.append(headers)
        for row in rows:
            ws.append([row.get(header, "") for header in headers])
    else:
        for row in rows:
            ws.append(row)
    for column_cells in ws.columns:
        width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 42)
        ws.column_dimensions[column_cells[0].column_letter].width = width


def create_plan(line1_path: Path, line2_path: Path, material_path: Path, output_path: Path) -> tuple[int, int]:
    exceptions: list[dict] = []
    line1 = read_monthly(line1_path, "통합1라인", filter_hyundai=False, exceptions=exceptions)
    line2 = read_monthly(line2_path, "통합2라인", filter_hyundai=True, exceptions=exceptions)
    materials = read_material(material_path, exceptions)

    material_index: dict[str, list[dict]] = defaultdict(list)
    for item in materials:
        material_index[item["logistics_no"]].append(item)

    plans = []
    for monthly in [*line1, *line2]:
        matches = material_index.get(monthly["production_no"], [])
        if len(matches) == 1:
            plans.append(to_plan(monthly, matches[0]))
        elif len(matches) == 0:
            exceptions.append(to_exception("자재소요 미매칭", monthly["line_type"], monthly["production_no"], monthly["sales_model"], "", "생산번호와 일치하는 카운터웨이트 자재소요가 없습니다."))
        else:
            exceptions.append(to_exception("자재소요 중복", monthly["line_type"], monthly["production_no"], monthly["sales_model"], ", ".join(m["material_no"] for m in matches), "동일 생산번호에 카운터웨이트 자재소요가 2건 이상입니다."))

    monthly_keys = {row["production_no"] for row in [*line1, *line2]}
    for material in materials:
        if material["logistics_no"] not in monthly_keys:
            exceptions.append(to_exception("월확정 미매칭", material["line_type"], material["logistics_no"], "", material["material_no"], "물류번호와 일치하는 월확정 생산번호가 없습니다."))

    plans.sort(key=lambda r: (r["착수일"], r["라인"], safe_int(r["순번"]), r["연번"], r["생산번호(물류번호)"]))
    dates = [row["착수일"] for row in plans if row["착수일"]]
    period = f"{min(dates).replace('-', '')}-{max(dates).replace('-', '')}" if dates else "START-END"
    if output_path.is_dir():
        output_path = output_path / f"원우ENG_CWT_생산계획_{period}.xlsx"

    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "생산계획", plans)
    write_sheet(wb, "예외목록", exceptions or [to_exception("예외 없음", "", "", "", "", "예외가 없습니다.")])
    write_sheet(wb, "검증요약", [
        ["구분", "건수", "비고"],
        ["월확정 통합1라인 사용", len(line1), line1_path.name],
        ["월확정 통합2라인 현대 제품", len(line2), line2_path.name],
        ["자재소요 카운터웨이트", len(materials), material_path.name],
        ["정상 생산계획", len(plans), "생산계획 시트"],
        ["예외", len(exceptions), "예외목록 시트"],
        ["착수기간", period, "월확정 착수일 기준"],
    ])
    wb.save(output_path)
    return len(plans), len(exceptions)


def safe_int(value) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return 999999


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(APP_TITLE, "월확정 통합1라인 파일을 선택하세요.")
    line1 = filedialog.askopenfilename(title="월확정 통합1라인 선택", filetypes=[("Excel", "*.xlsx *.xls")])
    if not line1:
        return
    messagebox.showinfo(APP_TITLE, "월확정 통합2라인 파일을 선택하세요.")
    line2 = filedialog.askopenfilename(title="월확정 통합2라인 선택", filetypes=[("Excel", "*.xlsx *.xls")])
    if not line2:
        return
    messagebox.showinfo(APP_TITLE, "물류번호별 자재소요현황 파일을 선택하세요.")
    material = filedialog.askopenfilename(title="물류번호별 자재소요현황 선택", filetypes=[("Excel", "*.xlsx *.xls")])
    if not material:
        return
    output_dir = filedialog.askdirectory(title="결과 엑셀을 저장할 폴더 선택")
    if not output_dir:
        return
    plan_count, exception_count = create_plan(Path(line1), Path(line2), Path(material), Path(output_dir))
    messagebox.showinfo(APP_TITLE, f"생산계획 생성 완료\n정상 {plan_count}건, 예외 {exception_count}건")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        messagebox.showerror(APP_TITLE, str(exc))
        raise
