# -*- coding: utf-8 -*-
"""3차 추출: 9개 피벗 페이지 필터의 아이템별 Visible 상태"""
import json, os
import win32com.client

SRC = r"D:\복호화\FI운용본부_수탁고양식_new.xlsb"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted")

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try: excel.Calculation = -4135
except Exception: pass

print("opening...", flush=True)
wb = excel.Workbooks.Open(SRC, UpdateLinks=0, ReadOnly=True)
print("opened.", flush=True)

try:
    ws = wb.Worksheets("PivotTables_EoM")
    out = []
    for i in range(1, ws.PivotTables().Count + 1):
        pt = ws.PivotTables(i)
        p = {"name": pt.Name, "page_fields": []}
        for j in range(1, pt.PageFields.Count + 1):
            fld = pt.PageFields.Item(j)
            d = {"name": fld.Name}
            try: d["multi"] = fld.EnableMultiplePageItems
            except Exception as e: d["multi_err"] = str(e)
            try: d["current_page"] = str(fld.CurrentPageName)
            except Exception:
                try: d["current_page"] = str(fld.CurrentPage)
                except Exception: pass
            items = []
            try:
                pis = fld.PivotItems()
                for k in range(1, pis.Count + 1):
                    it = pis.Item(k)
                    try: vis = it.Visible
                    except Exception: vis = None
                    items.append({"name": it.Name, "visible": vis})
            except Exception as e:
                d["items_err"] = str(e)
            d["items"] = items
            p["page_fields"].append(d)
        out.append(p)
        print(f"  {pt.Name}: {[f['name'] for f in p['page_fields']]}", flush=True)
    with open(os.path.join(OUT, "pivot_page_filters.json"), "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1, default=str)
    print("DONE", flush=True)
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
