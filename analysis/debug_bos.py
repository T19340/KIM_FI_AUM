# -*- coding: utf-8 -*-
import pandas as pd

bos = pd.read_csv(r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화\data\bos3218.csv',
                  dtype=str).fillna('')
b = bos.copy()
b['ffillB'] = b['모펀드'].replace('', pd.NA).ffill()
bad = b[b['최종모펀드'] != b['ffillB']]
print('불일치', len(bad), '행. 사례:')
cols = ['순번', '모펀드', '종류형모펀드', '종류형자펀드', '최종모펀드',
        '실질최종모펀드', 'ffillB']
print(bad[cols].head(12).to_string())
print()
print('최종모펀드 값 상위:', bad['최종모펀드'].value_counts().head(5).to_dict())

# 불일치 행의 ffillB(현재 블록 모펀드)가 다른 행에서 모/자로 등장하는가
child_set = set(bos['종류형자펀드']) - {''}
kmo_set = set(bos['종류형모펀드']) - {''}
mo_set = set(b['ffillB'].dropna())
print('bad행 ffillB가 타행 종류형자/모로 등장:',
      bad['ffillB'].isin(child_set | kmo_set).sum(), '/', len(bad))

# 반대 방향: 최종모펀드 값이 자기 블록 모펀드의 "모펀드"인가 —
# 즉 모펀드(ffillB)가 다른 블록에서 종류형자/모로 나오고 그 블록의 모펀드가 J인가
lookup_child = bos[bos['종류형자펀드'] != ''].drop_duplicates('종류형자펀드')
lookup_child = lookup_child.set_index('종류형자펀드')['최종모펀드']
lookup_kmo = bos[bos['종류형모펀드'] != ''].drop_duplicates('종류형모펀드')
lookup_kmo = lookup_kmo.set_index('종류형모펀드')['최종모펀드']
hop = bad['ffillB'].map(lambda c: lookup_child.get(c, lookup_kmo.get(c, None)))
match_hop = (hop == bad['최종모펀드']).sum()
print('1홉 상향(모펀드의 최종모) 일치:', match_hop, '/', len(bad))

# 혹시 J가 "순번 리셋" 블록 기준? 순번이 1로 리셋되는 지점 확인
b['순번n'] = pd.to_numeric(b['순번'], errors='coerce')
resets = (b['순번n'] == 1).sum()
print('순번=1 등장 횟수(블록 수):', resets)

# 순번=1 리셋 기준 블록의 첫 모펀드로 J를 만들면?
b['blk'] = (b['순번n'] == 1).cumsum()
blk_first_mo = b.groupby('blk')['모펀드'].transform(
    lambda s: s.replace('', pd.NA).ffill())
alt = b['모펀드'].replace('', pd.NA).groupby(b['blk']).ffill()
badalt = (b['최종모펀드'] != alt.fillna('')).sum()
print('블록(순번리셋) 내 ffill 기준 불일치:', badalt)
