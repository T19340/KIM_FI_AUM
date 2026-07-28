# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화')
from update_bos import read_bos_xlsx, derive, COLS
import pandas as pd

raw = read_bos_xlsx(r'K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata\BOS3218_20260721.xlsx')
print('shape:', raw.shape)
print('dtypes:', raw.dtypes.astype(str).unique().tolist())

b = raw.copy()
b.columns = COLS[:len(b.columns)]
col = b['모펀드']
print('모펀드: None', col.isna().sum(), '| 빈문자', (col == '').sum(),
      '| 값', ((col.notna()) & (col != '')).sum())
f = col.ffill()
print('ffill후: None', f.isna().sum(), '| 빈문자', (f == '').sum())
print('행 2394-2401:', [repr(v) for v in f.iloc[2393:2401].tolist()])

d = derive(raw)
print('derive후 모펀드 빈값:', (d['모펀드'].fillna('').astype(str).str.strip() == '').sum())
print('derive후 J NaN:', d['최종모펀드'].isna().sum())
print('derive 행수:', len(d), 'vs raw', len(raw))
