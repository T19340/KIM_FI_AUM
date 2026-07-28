# -*- coding: utf-8 -*-
"""미매핑 사모 17건을 BOS 가족관계·기존 매핑·원본 수익자 힌트로 분류."""
import sys
sys.path.insert(0, r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화')
import pandas as pd
import sutakgo as sg

RAW = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"
df = sg.build_dataset(RAW, verbose=False)
D21 = pd.Timestamp("2026-07-21")
bos = sg.load_bos().fillna("")
ben = sg.load_beneficiary()

UNMAPPED = ["06W41", "06W42", "06W57", "08M61", "08M62", "08M65", "08N56",
            "08N57", "08N58", "08P26", "08P39", "08P41", "08P52", "08P55",
            "08Q13", "08Q14", "08Q58"]

# 펀드코드 → 최근 원본 수익자 힌트 (전 기간, 최신값)
h = df[df["공백"].notna() & df["공백"].astype(str).str.strip().ne("")
       & df["공백"].astype(str).ne("nan")]
h = h.sort_values("처리일")
hint = h.groupby("펀드약칭")["공백"].last().to_dict()

# BOS 가족: 같은 K-그룹 (종류형모펀드 기준) 구성원
fam_by_k = {}
for _, r in bos.iterrows():
    k = r["종류형모펀드_k"]
    if not k:
        continue
    fam = fam_by_k.setdefault(k, set())
    fam.add(k)
    if r["종류형자펀드_l"]:
        fam.add(r["종류형자펀드_l"])

d21 = df[df["처리일"] == D21].set_index("펀드약칭")

print(f"{'코드':7s}{'유형':8s}{'수탁고':>9s}  가족 / 기존매핑 / 힌트")
print("-" * 100)
for code in UNMAPPED:
    row = d21.loc[code]
    jong = row["종류형구분"]          # 일반펀드/판매펀드/운용펀드
    amt = float(row["기획실수탁고분류"])
    # 가족 찾기: 자신이 K이거나, 자신이 속한 K그룹
    fams = set()
    for k, members in fam_by_k.items():
        if code in members:
            fams |= members
    fams.discard(code)
    mapped_fam = sorted(f for f in fams if f in ben.index)
    fam_hints = {f: hint[f] for f in sorted(fams) if f in hint}
    my_hint = hint.get(code, "")
    fam_str = ",".join(sorted(fams)) if fams else "(단독)"
    print(f"{code:7s}{str(jong):8s}{amt:9,.1f}  가족[{fam_str}]"
          f"  기존매핑{mapped_fam}  내힌트'{my_hint}'  가족힌트{fam_hints}")
