# -*- coding: utf-8 -*-
"""수익자 표 MID열: 83행(내부가 된 옛 마지막 행)의 굵은 아래 테두리 제거.

굵은 테두리는 표 바깥 경계 — '가장 아래 행'의 아래에만 있어야 한다.
내부 행(82행) 서식을 83행에 복사해 되돌린다.
"""
import win32com.client

DST = r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\template\FI운용본부수탁고_양식.xlsx'
xlEdgeBottom = 9

excel = win32com.client.DispatchEx('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(DST, UpdateLinks=0)
try:
    ws = wb.Worksheets('월말기준수탁고')

    def w(addr):
        b = ws.Range(addr).Borders(xlEdgeBottom)
        return (b.LineStyle, b.Weight)

    print('수정 전 아래테두리 (LineStyle, Weight):')
    for a in ('F82', 'F83', 'F84'):
        print(' ', a, ws.Range(a).Value, w(a))

    ws.Range('F82:I82').Copy()
    ws.Range('F83:I83').PasteSpecial(Paste=-4122)  # xlPasteFormats
    excel.CutCopyMode = False

    print('수정 후:')
    for a in ('F82', 'F83', 'F84'):
        print(' ', a, ws.Range(a).Value, w(a))
    wb.Save()
    print('저장 완료')
finally:
    wb.Close(SaveChanges=False)
    excel.Quit()
