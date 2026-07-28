# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'C:\Users\USER\Cowork\2026-07-23_수탁고-자동화')
import pandas as pd
import sutakgo as sg

RAW = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"
df = sg.build_dataset(RAW, verbose=False)
D21, D25 = pd.Timestamp("2026-07-21"), pd.Timestamp("2025-12-31")

# 1) 05B54 (법인용MMF2호(국공채)(모)) 7/21 파생값
r = df[(df["처리일"] == D21) & (df["펀드약칭"] == "05B54")]
print("[1] 05B54 7/21:", r[["종류형구분2", "최종모펀드코드", "최종모펀드유무",
                            "기획실수탁고분류"]].to_string())

# 2) 상/하위 5 (25YE→7/21) — 하위에 '전액 감소'가 남아 있는지
dates = sg.pick_dates(df, D21)
tb = sg.topbottom_block(df, dates, n=5)
for team in sg.TEAMS:
    print(f"[2] {team}")
    for key in ("top", "bottom"):
        for nm, base, cur, diff in tb[team][key]:
            wipe = " ←전액감소" if (pd.notna(base) and (pd.isna(cur) or cur == 0)
                                  and abs(diff) > 100) else ""
            print(f"   {key:6s} {nm[:34]:36s} {0 if pd.isna(base) else base:12,.1f}"
                  f" {0 if pd.isna(cur) else cur:12,.1f} {diff:12,.1f}{wipe}")

# 3) 7/14 시점 상하위(엑셀 검증 완료본)와 7/21 하위의 연속성:
#    7/14 엑셀 하위5 대비 7/21 하위5 구성 비교용 참고 출력
g14 = {t: sg.pivot_fund_diff(df, t, D25, pd.Timestamp("2026-07-14")) for t in sg.TEAMS}
g21 = {t: sg.pivot_fund_diff(df, t, D25, D21) for t in sg.TEAMS}
for t in sg.TEAMS:
    a = g14[t]["diff"].sort_values().head(5)
    b = g21[t]["diff"].sort_values().head(5)
    print(f"[3] {t} 하위5 7/14: {[(sg.clean_fund_name(i), round(v,0)) for i,v in a.items()]}")
    print(f"          7/21: {[(sg.clean_fund_name(i), round(v,0)) for i,v in b.items()]}")

# 4) 미매핑 사모 17건 목록 (질문 1)
d21 = df[df["처리일"] == D21]
fi = d21[d21["운용부서"].isin(["FI운용본부", "공동"])]
samo = fi[fi["공모사모구분"] == "사모"]
ben = sg.load_beneficiary()
un = samo[~samo["펀드약칭"].isin(ben.index)]
print("[4] 수익자 매핑 없는 FI 사모펀드:", len(un))
for _, x in un.iterrows():
    hint = "" if pd.isna(x["공백"]) or str(x["공백"]).strip() in ("", "nan") \
        else f"  (원본힌트: {x['공백']})"
    print(f"   {x['펀드약칭']}  {x['펀드명']}  {x['기획실수탁고분류']:.1f}억{hint}")
