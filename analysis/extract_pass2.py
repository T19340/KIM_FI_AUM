# -*- coding: utf-8 -*-
"""2차 추출: 파생열 수식 전체, ListObject 정의, PivotTables_EoM 수식, VBA 여부"""
import json, os
import win32com.client

SRC = r"D:\복호화\FI운용본부_수탁고양식_new.xlsb"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted")

def log(m): print(m, flush=True)
def save_json(name, obj):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)
def save_tsv(name, rows):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        for r in rows:
            f.write("\t".join("" if c is None else str(c).replace("\t", " ").replace("\n", "\\n") for c in r) + "\n")
def cell2d(v):
    if v is None: return [[None]]
    if not isinstance(v, tuple): return [[v]]
    return [list(r) for r in v]

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
try: excel.Calculation = -4135
except Exception: pass

log("opening...")
wb = excel.Workbooks.Open(SRC, UpdateLinks=0, ReadOnly=True)
log("opened. HasVBProject=" + str(wb.HasVBProject))

try:
    ws = wb.Worksheets("MOS2101")
    nrows = ws.UsedRange.Rows.Count

    # ListObject 정의
    lots = []
    for lo in ws.ListObjects:
        lots.append({"name": lo.Name, "range": lo.Range.Address,
                     "header_row": lo.HeaderRowRange.Row,
                     "data_rows": lo.DataBodyRange.Rows.Count,
                     "cols": lo.Range.Columns.Count})
    save_json("mos2101_listobjects.json", lots)
    log("listobjects: " + json.dumps(lots, ensure_ascii=False))

    # 파생열 258~274 마지막 행 수식 + 헤더
    cols = []
    for c in range(258, 275):
        cols.append({"col": c, "header": ws.Cells(1, c).Value,
                     "formula_last": ws.Cells(nrows, c).Formula,
                     "formula_r1c1": ws.Cells(nrows, c).FormulaR1C1})
    save_json("mos2101_derived_cols_full.json", cols)

    # PivotTables_EoM 전체 수식
    ws = wb.Worksheets("PivotTables_EoM")
    ur = ws.UsedRange
    save_tsv("pivottables_eom_formulas.tsv", cell2d(ur.Formula))
    log(f"eom formulas saved {ur.Rows.Count}x{ur.Columns.Count}")

    # 수익자 시트 전체 수식
    ws = wb.Worksheets("수익자")
    save_tsv("수익자_formulas.tsv", cell2d(ws.UsedRange.Formula))

    # BOS3218 1행 헤더
    ws = wb.Worksheets("BOS3218")
    save_tsv("bos3218_header_rows1-3.tsv", cell2d(ws.Range(ws.Cells(1, 1), ws.Cells(3, ws.UsedRange.Columns.Count)).Value))

    # 월말기준수탁고 rows1-6 (주소 확인용 좁은 덤프)
    ws = wb.Worksheets("월말기준수탁고")
    save_tsv("월말기준수탁고_rows1-6.tsv", cell2d(ws.Range("A1:N6").Formula))

    log("DONE")
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
