# -*- coding: utf-8 -*-
"""원본 양식 복원 후 ABL생명 행을 COM으로 삽입 (조건부 서식 보존)."""
import shutil
import win32com.client

SRC = r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\output\FI운용본부수탁고_양식.xlsx'
DST = r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\template\FI운용본부수탁고_양식.xlsx'

shutil.copy2(SRC, DST)
print('원본 복원 완료')

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(DST, UpdateLinks=0)
try:
    ws = wb.Worksheets('월말기준수탁고')
    vals = ws.Range('F64:F100').Value
    names = []
    for (v,) in vals:
        if v is None or str(v).strip() == '':
            break
        names.append(str(v).strip())
    print('기존 MID:', len(names), '명')
    assert 'ABL생명' not in names and 'KDB생명' in names
    names.insert(names.index('KDB생명') + 1, 'ABL생명')
    last = 64 + len(names) - 1          # 새 마지막 행 (84)
    ws.Range('F83:I83').Copy()
    ws.Range(f'F{last}:I{last}').PasteSpecial(Paste=-4122)  # xlPasteFormats
    excel.CutCopyMode = False
    ws.Range(f'F64:F{last}').Value = [[n] for n in names]
    wb.Save()
    print(f'ABL생명 삽입 완료 (F{64 + names.index("ABL생명")}), MID {len(names)}명')
finally:
    wb.Close(False)
    excel.Quit()
