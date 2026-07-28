# -*- coding: utf-8 -*-
"""2025-12-31 기준 FI운용본부 수탁고 + 펀드명 'ESG' 포함분 집계.

수탁고 정의는 기존 리포트와 동일:
  값 = 기획실수탁고분류 (ETF는 순자산(후), 그 외 원본액, 억원 단위)
  필터 = 운용부서∈{FI운용본부,공동} · 모자구분∈{일반펀드,자펀드}
         · 종류형구분∈{일반펀드,판매펀드} · 제외조건=""
  본부 계 = 팀분류∈{FI운용1부, FI운용2부, 해외FI운용부} 합
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import sutakgo as sg

TARGET = pd.Timestamp("2025-12-31")

df = sg.load_history()
df["처리일"] = pd.to_datetime(df["처리일"])
df["기획실수탁고분류"] = pd.to_numeric(df["기획실수탁고분류"], errors="coerce")

base = sg._base(df, [TARGET])                       # 리포트와 동일한 필터
base = base[base["팀분류"].isin(sg.TEAMS)]

print("=" * 72)
print(f"[1] FI운용본부 수탁고 — {TARGET.date()} 기준 (억원)")
print("=" * 72)
pv = base.pivot_table(index="팀분류", columns="펀드일임구분",
                      values="기획실수탁고분류", aggfunc="sum")
pv = pv.reindex(index=sg.TEAMS, columns=[c for c in sg.GUBUN_ORDER if c in pv.columns])
pv["계"] = pv.sum(axis=1)
pv.loc["본부 계"] = pv.sum()
print(pv.round(1).to_string(na_rep="-"))
TOTAL = pv.loc["본부 계", "계"]
print(f"\n본부 계 = {TOTAL:,.1f} 억원  ({TOTAL/10000:,.2f} 조원), 펀드 {len(base)}건")

print()
print("=" * 72)
print("[2] 펀드명에 'ESG' 포함 (대소문자 무시)")
print("=" * 72)
esg = base[base["펀드명"].astype(str).str.contains("ESG", case=False, na=False)]
cols = ["펀드약칭", "펀드명", "팀분류", "펀드일임구분", "기획실펀드분류",
        "공모사모구분", "종류형구분", "기획실수탁고분류"]
e = esg[cols].sort_values("기획실수탁고분류", ascending=False)
pd.set_option("display.width", 250, "display.max_colwidth", 60)
print(e.to_string(index=False, float_format=lambda v: f"{v:,.1f}"))
ESG = e["기획실수탁고분류"].sum()
print(f"\nESG 합계 = {ESG:,.1f} 억원   (본부 계 대비 {ESG/TOTAL*100:.2f}%), {len(e)}건")
print("\n[팀별]")
print(esg.groupby("팀분류")["기획실수탁고분류"].agg(["sum", "count"]).round(1).to_string())
print("\n[유형별]")
print(esg.groupby("기획실펀드분류")["기획실수탁고분류"].agg(["sum", "count"]).round(1).to_string())

# --- 참고: 필터에서 빠진 ESG 행(모펀드·운용펀드 등 이중계상 방지분) 확인 ---
print()
print("=" * 72)
print("[참고] 같은 일자에서 필터에 걸러진 'ESG' 행 (이중계상 방지로 제외된 것)")
print("=" * 72)
allday = df[(df["처리일"] == TARGET)
            & df["펀드명"].astype(str).str.contains("ESG", case=False, na=False)]
excl = allday[~allday.index.isin(esg.index)]
if len(excl):
    print(excl[["펀드약칭", "펀드명", "운용부서", "팀분류", "모자구분", "종류형구분",
                "제외조건", "기획실수탁고분류"]]
          .to_string(index=False, float_format=lambda v: f"{v:,.1f}"))
else:
    print("없음")
