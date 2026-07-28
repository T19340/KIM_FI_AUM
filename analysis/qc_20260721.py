# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화')
import pandas as pd
import sutakgo as sg

RAW = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"
df = sg.build_dataset(RAW, verbose=False)
D21 = pd.Timestamp("2026-07-21")
D14 = pd.Timestamp("2026-07-14")

d21 = df[df["처리일"] == D21]
d14 = df[df["처리일"] == D14]
print("[QC1] 행수: 7/21", len(d21), "vs 7/14", len(d14))

# 신규 펀드 (7/14에 없던 코드)
new_funds = set(d21["펀드약칭"]) - set(d14["펀드약칭"])
print("[QC2] 7/21 신규 펀드:", sorted(new_funds))
bos = sg.load_bos().fillna("")
in_bos = set(bos["종류형모펀드_k"]) | set(bos["종류형자펀드_l"]) | set(bos["모펀드"])
ben = sg.load_beneficiary()
for f in sorted(new_funds):
    r = d21[d21["펀드약칭"] == f].iloc[0]
    memo = []
    if f not in in_bos:
        memo.append("BOS미수록(단독펀드 처리)")
    if r["공모사모구분"] == "사모" and f not in ben.index:
        memo.append("수익자매핑 없음!")
    print(f"   {f} {str(r['펀드명'])[:30]} 사모={r['공모사모구분']} 부서={r['운용부서']}: {'; '.join(memo) or 'OK'}")

# 사모인데 수익자 매핑 없는 펀드 (FI 범위, 7/21)
fi21 = d21[d21["운용부서"].isin(["FI운용본부", "공동"])]
samo = fi21[fi21["공모사모구분"] == "사모"]
unmapped = samo[~samo["펀드약칭"].isin(ben.index)]
print("[QC3] 7/21 FI 사모 중 수익자 매핑 없음:", len(unmapped), "펀드")
if len(unmapped):
    print(unmapped[["펀드약칭", "펀드명", "공백"]].head(10).to_string())

# 팀분류 커버리지 (FI 필터 범위에서 3개 팀 외 값)
base = sg._base(df, [D21])
others = base[~base["팀분류"].isin(sg.TEAMS)]
print("[QC4] FI 필터 범위에서 3팀 외 팀분류:", others["팀분류"].unique().tolist(),
      "(", len(others), "행 )")

# 펀드일임구분 FALSE 데이터
f21 = base[base["펀드일임구분"].astype(str) == "FALSE"]
print("[QC5] 펀드일임구분 FALSE 행:", len(f21))

# 연속성: 팀별 합계 7/14 vs 7/21
pv = sg.pivot_team(df, [D14, D21])
t = pv.groupby(level=0).sum()
t.columns = ["7/14", "7/21"]
t["증감"] = t["7/21"] - t["7/14"]
print("[QC6] 팀별 수탁고 (억원):")
print(t.round(1).to_string())
print("   본부계:", t["7/14"].sum().round(1), "→", t["7/21"].sum().round(1))
