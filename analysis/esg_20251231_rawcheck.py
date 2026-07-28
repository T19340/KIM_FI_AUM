# -*- coding: utf-8 -*-
"""2025-12-31: 히스토리(워크북 저장값) vs rawdata CSV 직접 계산 대조."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import sutakgo as sg

RAW = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"
T = pd.Timestamp("2025-12-31")

hist = sg.load_history()
hist["처리일"] = pd.to_datetime(hist["처리일"])
hist["기획실수탁고분류"] = pd.to_numeric(hist["기획실수탁고분류"], errors="coerce")

bos, ben = sg.load_bos(), sg.load_beneficiary()
name_lookup = hist.drop_duplicates("펀드약칭").set_index("펀드약칭")["펀드명"].to_dict()


def from_csv(fname):
    raw = sg.parse_mos_csv(os.path.join(RAW, fname), list(hist.columns))
    print(f"{fname}: {len(raw)}행, 처리일 {sorted(set(raw['처리일'].dt.date))}")
    return sg.compute_derived(raw, bos, ben, name_lookup)


def agg(df, label):
    b = sg._base(df, [T])
    b = b[b["팀분류"].isin(sg.TEAMS)]
    esg = b[b["펀드명"].astype(str).str.contains("ESG", case=False, na=False)]
    print(f"\n[{label}] 대상 {len(b)}건")
    t = b.groupby("팀분류")["기획실수탁고분류"].sum().reindex(sg.TEAMS)
    for k, v in t.items():
        print(f"   {k:10s} {v:12,.1f}")
    print(f"   {'본부 계':10s} {t.sum():12,.1f}")
    print(f"   ESG {len(esg)}건 {esg['기획실수탁고분류'].sum():,.1f}")
    return b, esg


csv = from_csv("MOS2101_20251231.csv")
h = hist[hist["처리일"] == T]
print(f"히스토리 2025-12-31: {len(h)}행")

bh, eh = agg(h, "히스토리(워크북 저장값)")
bc, ec = agg(csv, "rawdata CSV 직접계산")

# --- 행 단위 대조 ---
key = "펀드"
mh = bh.set_index(key)["기획실수탁고분류"]
mc = bc.set_index(key)["기획실수탁고분류"]
print("\n--- 행 단위 차이 ---")
print("히스토리에만:", sorted(set(mh.index) - set(mc.index)))
print("CSV에만    :", sorted(set(mc.index) - set(mh.index)))
common = sorted(set(mh.index) & set(mc.index))
diff = (mc[common] - mh[common]).abs()
print(f"공통 {len(common)}건 중 금액 상이: {(diff > 1e-6).sum()}건, 최대차 {diff.max():.6f}")
if (diff > 1e-6).any():
    print(pd.DataFrame({"hist": mh[common], "csv": mc[common]})[diff > 1e-6])

# 팀분류 배정 차이 (과거 동결 기준 vs 현행 수식)
th = bh.set_index(key)["팀분류"]
tc = bc.set_index(key)["팀분류"]
td = [(k, th[k], tc[k]) for k in common if th[k] != tc[k]]
print(f"\n팀분류 배정 상이: {len(td)}건")
for k, a, b_ in td[:20]:
    print(f"   {k}: 히스토리 {a} → CSV계산 {b_}  ({mh[k]:,.1f}억)")

# ESG 대조
print("\n--- ESG 대조 ---")
sh = eh.set_index(key)["기획실수탁고분류"].sort_index()
sc = ec.set_index(key)["기획실수탁고분류"].sort_index()
print("코드 집합 동일:", set(sh.index) == set(sc.index))
print(f"합계 히스토리 {sh.sum():,.4f} / CSV {sc.sum():,.4f} / 차 {sc.sum()-sh.sum():.6f}")

# 추가 CSV 확인
extra = os.path.join(RAW, "MOS2101_20251231_추가.csv")
if os.path.exists(extra):
    print("\n--- MOS2101_20251231_추가.csv ---")
    ex = sg.parse_mos_csv(extra, list(hist.columns))
    print(f"{len(ex)}행, 처리일 {sorted(set(ex['처리일'].dt.date))}")
    print("펀드코드가 본 CSV에 없는 것:",
          sorted(set(ex['펀드'].astype(str)) - set(csv['펀드'].astype(str)))[:20])
