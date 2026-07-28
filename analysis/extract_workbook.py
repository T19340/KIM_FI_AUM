# -*- coding: utf-8 -*-
"""FI운용본부_수탁고양식_new.xlsb 구조 추출 (읽기 전용 COM)"""
import json, os, sys, traceback
import win32com.client

SRC = r"D:\복호화\FI운용본부_수탁고양식_new.xlsb"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted")
os.makedirs(OUT, exist_ok=True)

def log(msg):
    print(msg, flush=True)

def save_json(name, obj):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1, default=str)

def save_tsv(name, rows):
    with open(os.path.join(OUT, name), "w", encoding="utf-8") as f:
        for r in rows:
            f.write("\t".join("" if c is None else str(c) for c in r) + "\n")

def cell2d(v):
    # COM Range.Value 는 단일셀이면 스칼라, 아니면 tuple of tuples
    if v is None: return [[None]]
    if not isinstance(v, tuple): return [[v]]
    return [list(r) for r in v]

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
excel.EnableEvents = False
try:
    excel.Calculation = -4135  # xlCalculationManual
except Exception:
    pass

log("opening workbook (90MB, may take a while)...")
wb = excel.Workbooks.Open(SRC, UpdateLinks=0, ReadOnly=True)
log("opened.")

try:
    # ---- 1. 시트 목록 ----
    sheets = []
    for ws in wb.Worksheets:
        try:
            ur = ws.UsedRange
            info = {"name": ws.Name, "visible": ws.Visible,
                    "rows": ur.Rows.Count, "cols": ur.Columns.Count,
                    "first_cell": ur.Cells(1, 1).Address}
        except Exception as e:
            info = {"name": ws.Name, "error": str(e)}
        sheets.append(info)
    save_json("sheets.json", sheets)
    log("sheets: " + ", ".join(s["name"] for s in sheets))

    # ---- 2. MOS2101 ----
    ws = wb.Worksheets("MOS2101")
    ur = ws.UsedRange
    nrows, ncols = ur.Rows.Count, ur.Columns.Count
    log(f"MOS2101 used range: {nrows} rows x {ncols} cols")

    hdr = cell2d(ws.Range(ws.Cells(1, 1), ws.Cells(3, ncols)).Value)
    save_tsv("mos2101_header_rows1-3.tsv", hdr)

    # IY(259)열 수식 전 구간 → R1C1 로 읽어 연속 동일 구간(regime) 압축
    iy_col = 259
    log("reading IY column formulas (R1C1)...")
    f_iy = cell2d(ws.Range(ws.Cells(1, iy_col), ws.Cells(nrows, iy_col)).FormulaR1C1)
    log("reading 영업일 column values...")
    d_col = cell2d(ws.Range(ws.Cells(1, 2), ws.Cells(nrows, 2)).Value)
    regimes = []
    cur = None
    for i in range(nrows):
        f = f_iy[i][0]
        f = "" if f is None else str(f)
        if cur is None or f != cur["formula_r1c1"]:
            if cur is not None:
                regimes.append(cur)
            cur = {"formula_r1c1": f, "row_start": i + 1, "row_end": i + 1,
                   "date_start": d_col[i][0], "date_end": d_col[i][0]}
        else:
            cur["row_end"] = i + 1
            cur["date_end"] = d_col[i][0]
    if cur is not None:
        regimes.append(cur)
    save_json("mos2101_iy_regimes.json", regimes)
    log(f"IY regimes: {len(regimes)}")

    # A1 스타일 수식 샘플 (마지막 데이터 행 + 각 regime 대표행)
    samples = {}
    for rg in regimes[:200]:
        r = rg["row_end"]
        samples[str(r)] = ws.Cells(r, iy_col).Formula
    save_json("mos2101_iy_samples_a1.json", samples)

    # 주변 열(IX~JB)에 다른 수기/수식 열이 있는지 확인: 1행 헤더 + 마지막 행 수식
    around = []
    for c in range(min(ncols, 255), min(ncols + 3, 262) + 1):
        try:
            around.append({"col": c, "header1": ws.Cells(1, c).Value,
                           "header2": ws.Cells(2, c).Value,
                           "header3": ws.Cells(3, c).Value,
                           "last_formula": ws.Cells(nrows, c).Formula})
        except Exception as e:
            around.append({"col": c, "error": str(e)})
    save_json("mos2101_tail_cols.json", around)

    # 최근 5행 값 샘플
    tail = cell2d(ws.Range(ws.Cells(nrows - 4, 1), ws.Cells(nrows, min(ncols, 60))).Value)
    save_tsv("mos2101_tail_rows_first60cols.tsv", tail)

    # ---- 3. BOS3218 ----
    try:
        ws = wb.Worksheets("BOS3218")
        ur = ws.UsedRange
        nr, nc = ur.Rows.Count, ur.Columns.Count
        log(f"BOS3218: {nr} x {nc}")
        head = cell2d(ws.Range(ws.Cells(1, 1), ws.Cells(min(nr, 8), nc)).Value)
        save_tsv("bos3218_head.tsv", head)
        tailr = cell2d(ws.Range(ws.Cells(max(1, nr - 4), 1), ws.Cells(nr, nc)).Value)
        save_tsv("bos3218_tail.tsv", tailr)
        save_json("bos3218_info.json", {"rows": nr, "cols": nc})
    except Exception:
        log("BOS3218 error:\n" + traceback.format_exc())

    # ---- 4. 수익자 ----
    try:
        ws = wb.Worksheets("수익자")
        ur = ws.UsedRange
        nr, nc = ur.Rows.Count, ur.Columns.Count
        log(f"수익자: {nr} x {nc}")
        if nr <= 5000:
            allv = cell2d(ur.Value)
            save_tsv("수익자_full.tsv", allv)
        else:
            save_tsv("수익자_head.tsv", cell2d(ws.Range(ws.Cells(1, 1), ws.Cells(50, nc)).Value))
        # 수식 여부 확인 (첫 데이터 행)
        save_json("수익자_row2_formulas.json",
                  {"formulas": cell2d(ws.Range(ws.Cells(2, 1), ws.Cells(2, nc)).Formula)})
    except Exception:
        log("수익자 error:\n" + traceback.format_exc())

    # ---- 5. PivotTables_EoM 피벗 정의 ----
    ORIENT = {1: "Row", 2: "Column", 3: "Page", 4: "Data", 0: "Hidden"}
    FUNC = {-4157: "Sum", -4112: "Count", -4106: "Average", -4136: "Max",
            -4139: "Min", -4149: "Product", -4155: "CountNums",
            -4158: "StdDev", -4159: "StdDevp", -4164: "Var", -4165: "Varp"}
    try:
        ws = wb.Worksheets("PivotTables_EoM")
        ur = ws.UsedRange
        log(f"PivotTables_EoM: {ur.Rows.Count} x {ur.Columns.Count}, pivots={ws.PivotTables().Count}")
        pivots = []
        for i in range(1, ws.PivotTables().Count + 1):
            pt = ws.PivotTables(i)
            p = {"name": pt.Name, "location": pt.TableRange2.Address}
            try:
                p["source"] = pt.PivotCache().SourceData
            except Exception as e:
                p["source_error"] = str(e)
            for coll, key in [(pt.RowFields, "row_fields"), (pt.ColumnFields, "column_fields"),
                              (pt.PageFields, "page_fields")]:
                items = []
                for j in range(1, coll.Count + 1):
                    fld = coll.Item(j)
                    d = {"name": fld.Name, "source_name": fld.SourceName}
                    try:
                        d["current_page"] = str(fld.CurrentPage)
                    except Exception:
                        pass
                    try:
                        vis = fld.VisibleItems
                        if vis.Count <= 60:
                            d["visible_items"] = [vis.Item(k).Name for k in range(1, vis.Count + 1)]
                        else:
                            d["visible_items_count"] = vis.Count
                    except Exception:
                        pass
                    items.append(d)
                p[key] = items
            dfs = []
            for j in range(1, pt.DataFields.Count + 1):
                fld = pt.DataFields.Item(j)
                dfs.append({"name": fld.Name, "source_name": fld.SourceName,
                            "function": FUNC.get(fld.Function, fld.Function),
                            "number_format": fld.NumberFormat})
            p["data_fields"] = dfs
            try:
                cfs = pt.CalculatedFields()
                p["calculated_fields"] = [{"name": cfs.Item(k).Name, "formula": cfs.Item(k).Formula}
                                          for k in range(1, cfs.Count + 1)]
            except Exception:
                pass
            pivots.append(p)
            log(f"  pivot: {pt.Name} @ {p['location']}")
        save_json("pivots_eom.json", pivots)
        # 시트 자체 값도 저장 (피벗 배치 파악)
        vals = cell2d(ws.UsedRange.Value)
        save_tsv("pivottables_eom_values.tsv", vals)
    except Exception:
        log("PivotTables_EoM error:\n" + traceback.format_exc())

    # ---- 6. 월말기준수탁고 ----
    try:
        ws = wb.Worksheets("월말기준수탁고")
        ur = ws.UsedRange
        nr, nc = ur.Rows.Count, ur.Columns.Count
        log(f"월말기준수탁고: {nr} x {nc}")
        save_tsv("월말기준수탁고_values.tsv", cell2d(ur.Value))
        save_tsv("월말기준수탁고_formulas.tsv", cell2d(ur.Formula))
    except Exception:
        log("월말기준수탁고 error:\n" + traceback.format_exc())

    # ---- 7. 다른 시트 중 피벗 있는 시트 / 이름정의 ----
    try:
        pv_sheets = []
        for ws in wb.Worksheets:
            try:
                c = ws.PivotTables().Count
                if c:
                    pv_sheets.append({"sheet": ws.Name, "pivot_count": c})
            except Exception:
                pass
        save_json("pivot_sheets.json", pv_sheets)
        names = []
        for nm in wb.Names:
            try:
                names.append({"name": nm.Name, "refers_to": nm.RefersTo})
            except Exception:
                pass
        save_json("defined_names.json", names)
    except Exception:
        log("names error:\n" + traceback.format_exc())

    log("DONE")
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
