# -*- coding: utf-8 -*-
import sys
import pandas as pd

sys.path.insert(0, r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화')
import sutakgo as sg

D = r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\data'
new = pd.read_csv(D + r'\bos3218.csv', dtype=str).fillna('')
old = pd.read_csv(D + r'\bos3218_prev.csv', dtype=str).fillna('')

# 1) pd.read_excel 이 '' 마커를 보존했는지: 새 CSV에서 최종모펀드가 빈 행 수
print('new: J빈값 행', (new['최종모펀드'] == '').sum(),
      '/ 모펀드 빈값 행', (new['모펀드'] == '').sum())
print('old: J빈값 행', (old['최종모펀드'] == '').sum(),
      '/ 모펀드 빈값 행', (old['모펀드'] == '').sum())

# 2) 부재 구역 대표 펀드의 J 비교 (01179 등)
for code in ['01179', '01Z18', '07211', '07203']:
    n_rows = new[new['종류형모펀드_k'] == code]
    o_rows = old[old['종류형모펀드_k'] == code]
    nj = n_rows['최종모펀드'].unique().tolist()[:3]
    oj = o_rows['최종모펀드'].unique().tolist()[:3]
    print(f'{code}: old J={oj}  new J={nj}')

# 3) 법인용MMF2호(국공채)(모) 코드 찾기 + 7/21 vs 7/14 파생값
hist = sg.load_history()
hist['처리일'] = pd.to_datetime(hist['처리일'])
mmf2 = hist[hist['펀드명'].astype(str).str.contains('법인용MMF2호', na=False)]
print()
print('법인용MMF2호 관련 펀드:',
      mmf2[['펀드', '펀드명']].drop_duplicates().to_string())
target = mmf2[mmf2['펀드명'].astype(str).str.contains(r'\(모\)', na=False)]
code = target['펀드약칭'].iloc[-1] if len(target) else None
print('모펀드 코드:', code)
if code:
    h14 = hist[(hist['처리일'] == pd.Timestamp('2026-07-14'))
               & (hist['펀드약칭'] == code)]
    print('7/14 저장값:', h14[['종류형구분2', '최종모펀드코드', '최종모펀드유무',
                              '기획실수탁고분류']].to_string())
    # 새 BOS 에서 이 코드가 어디 등장?
    for col in ['모펀드', '종류형모펀드_k', '종류형자펀드_l']:
        o = (old[col] == code).sum(); n = (new[col] == code).sum()
        print(f'  BOS {col}: old {o}회 / new {n}회')
    nk = new[new['종류형모펀드_k'] == code]
    if len(nk):
        print('  new BOS 해당행 J:', nk['최종모펀드'].unique().tolist())
