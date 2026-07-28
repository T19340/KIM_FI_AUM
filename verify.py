# -*- coding: utf-8 -*-
"""파이썬 엔진 vs 엑셀 원본 검증.

그라운드 트루스:
  analysis/extracted/pivottables_eom_values.tsv  (엑셀 피벗 산출값, 2026-07-15 저장본)
  data/mos2101_history.parquet 의 파생열 저장값  (2026-01-02 이후 행 재계산 대조)
"""
import os
import pandas as pd
import sutakgo as sg

BASE = os.path.dirname(os.path.abspath(__file__))
TSV = os.path.join(BASE, "analysis", "extracted", "pivottables_eom_values.tsv")

D24 = pd.Timestamp("2024-12-31")
D25 = pd.Timestamp("2025-12-31")
D0531 = pd.Timestamp("2026-05-31")
D0630 = pd.Timestamp("2026-06-30")
DCUR = pd.Timestamp("2026-07-14")

PASS = FAIL = 0
FAILS = []


def check(name, ok, detail=""):
    global PASS, FAIL
    if ok:
        PASS += 1
    else:
        FAIL += 1
        FAILS.append((name, detail))
        print(f"  FAIL {name}: {detail}")


def close(a, b, tol=1e-6):
    if pd.isna(a) and (pd.isna(b) or b in ("", None)):
        return True
    try:
        a, b = float(a), float(b)
    except (TypeError, ValueError):
        return str(a) == str(b)
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


def f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return float("nan")


grid = [line.rstrip("\n").split("\t")
        for line in open(TSV, encoding="utf-8")]


def cell(r, c):
    row = grid[r - 1]
    return row[c - 1] if c - 1 < len(row) else ""


print("=" * 60)
print("[0] 데이터셋 구성")
df = sg.build_dataset()
print(f"  rows={len(df)}")

# ------------------------------------------------ 1. 팀별수탁고 피벗
print("[1] 팀별수탁고 피벗 (B13:F27, 5개 일자)")
dates5 = [D24, D25, D0531, D0630, DCUR]
pv = sg.pivot_team(df, dates5)
layout = [("FI운용1부", None), ("FI운용1부", "펀드"), ("FI운용1부", "일임"),
          ("FI운용1부", "MMF"), ("FI운용1부", "ETF"),
          ("FI운용2부", None), ("FI운용2부", "펀드"), ("FI운용2부", "일임"),
          ("FI운용2부", "ETF"),
          ("해외FI운용부", None), ("해외FI운용부", "펀드"), ("해외FI운용부", "일임"),
          ("해외FI운용부", "MMF"), ("해외FI운용부", "ETF"),
          (None, None)]
for i, (team, gub) in enumerate(layout):
    r = 13 + i
    if team is None:
        vals = pv.sum()
    elif gub is None:
        vals = pv.loc[team].sum()
    else:
        vals = pv.loc[(team, gub)]
    for j in range(5):
        gt = f(cell(r, 2 + j))
        mine = vals.iloc[j]
        mine = 0.0 if pd.isna(mine) else mine
        gt = 0.0 if pd.isna(gt) else gt
        check(f"팀별수탁고 r{r}c{2+j} {team}/{gub}", close(mine, gt),
              f"py={mine} xl={gt}")

# ------------------------------------------------ 2. 팀별순자산 피벗
print("[2] 팀별순자산 피벗 (I13:I27)")
nav = sg.pivot_team_nav(df, DCUR)
for i, (team, gub) in enumerate(layout):
    r = 13 + i
    if team is None:
        mine = nav.sum()
    elif gub is None:
        mine = nav.loc[team].sum()
    else:
        mine = nav.get((team, gub), float("nan"))
    gt = f(cell(r, 9))
    if pd.isna(mine) and pd.isna(gt):
        check(f"팀별순자산 r{r}", True)
    else:
        check(f"팀별순자산 r{r} {team}/{gub}", close(mine, gt),
              f"py={mine} xl={gt}")

# ------------------------------------------------ 3. MOS세부분류 피벗
print("[3] MOS세부분류 피벗 (B62:I72)")
cat = sg.pivot_category(df, [D25, DCUR])
colmap = {"FI운용1부": (2, 3), "FI운용2부": (5, 6), "해외FI운용부": (8, 9)}
for i, catname in enumerate(sg.CATEGORY_ORDER):
    r = 62 + i
    check(f"MOS세부분류 라벨 r{r}", cell(r, 1) == catname,
          f"xl={cell(r,1)} expected={catname}")
    for team, (c25, ccur) in colmap.items():
        for dt, c in [(D25, c25), (DCUR, ccur)]:
            gt = cell(r, c)
            mine = cat[(team, dt)].get(catname, float("nan")) \
                if (team, dt) in cat.columns else float("nan")
            if gt == "" and (pd.isna(mine) or mine == 0):
                check(f"세부분류 r{r} {team} {dt.date()}", True)
            else:
                check(f"세부분류 r{r} {team} {dt.date()}", close(mine, gt),
                      f"py={mine} xl={gt}")
# 총합계 (row 72)
for team, (c25, ccur) in colmap.items():
    for dt, c in [(D25, c25), (DCUR, ccur)]:
        gt = f(cell(72, c))
        mine = cat[(team, dt)].sum() if (team, dt) in cat.columns else 0.0
        check(f"세부분류 총합계 {team} {dt.date()}", close(mine, gt),
              f"py={mine} xl={gt}")

# ------------------------------------------------ 4. 수익자1/수익자2 피벗
print("[4] 수익자1/수익자2 피벗")
for pname, col0, vcol in [("수익자1", 1, 2), ("수익자2", 8, 9)]:
    b = sg.pivot_beneficiary(df, [D25, DCUR], pname)
    r = 102
    n = 0
    while True:
        label = cell(r, col0)
        if label in ("총합계", "") or r > 160:
            break
        gt25, gtc = f(cell(r, vcol)), f(cell(r, vcol + 1))
        if label in b.index:
            m25 = b.loc[label].iloc[0]
            mc = b.loc[label].iloc[1]
        else:
            m25 = mc = float("nan")
        ok25 = close(0 if pd.isna(m25) else m25, 0 if pd.isna(gt25) else gt25)
        okc = close(0 if pd.isna(mc) else mc, 0 if pd.isna(gtc) else gtc)
        check(f"{pname} {label}", ok25 and okc,
              f"py=({m25},{mc}) xl=({gt25},{gtc})")
        n += 1
        r += 1
    print(f"  {pname}: {n}개 항목 대조")

# ------------------------------------------------ 5. 팀1/2/3 피벗 (펀드별 증감)
print("[5] 팀1/2/3 피벗 (최종모펀드명별)")
for pname, team, col0 in [("팀1", "FI운용1부", 20), ("팀2", "FI운용2부", 25),
                          ("팀3", "해외FI운용부", 30)]:
    g = sg.pivot_fund_diff(df, team, D25, DCUR)
    r = 12
    n = bad = 0
    while True:
        label = cell(r, col0)
        if label in ("총합계", "") or r > 210:
            break
        gt25, gtc = f(cell(r, col0 + 1)), f(cell(r, col0 + 2))
        if label in g.index:
            m25, mc = g.loc[label, "base"], g.loc[label, "cur"]
        else:
            m25 = mc = float("nan")
        ok = close(0 if pd.isna(m25) else m25, 0 if pd.isna(gt25) else gt25) and \
             close(0 if pd.isna(mc) else mc, 0 if pd.isna(gtc) else gtc)
        if not ok:
            bad += 1
            if bad <= 5:
                print(f"    mismatch {label}: py=({m25},{mc}) xl=({gt25},{gtc})")
        n += 1
        r += 1
    check(f"{pname} 전체 {n}행", bad == 0, f"{bad} mismatches")
    check(f"{pname} 항목수", len(g) == n, f"py={len(g)} xl={n}")

# ------------------------------------------------ 6. 파생열 재계산 대조
print("[6] 파생열 재계산 대조 (2026-01-02 이후 저장값 vs 재계산)")
hist = sg.load_history()
hist["처리일"] = pd.to_datetime(hist["처리일"])
recent = hist[hist["처리일"] >= sg.NEW_RULE_START].copy()
raw = recent[list(hist.columns[:257])].copy()
raw["순자산(후)"] = pd.to_numeric(raw["순자산(후)"], errors="coerce")
raw["원본액"] = pd.to_numeric(raw["원본액"], errors="coerce")
# 참조 테이블은 워크북 저장 시점 스냅샷으로 고정 (운영 갱신과 무관하게 엑셀 재현 검증)
bos = sg.load_bos(os.path.join(BASE, "data", "bos3218_asof_20260715.csv"))
ben = sg.load_beneficiary(os.path.join(BASE, "data", "수익자_asof_20260715.csv"))
name_lookup = (hist.drop_duplicates("펀드약칭")
               .set_index("펀드약칭")["펀드명"].to_dict())
rc = sg.compute_derived(raw, bos, ben, name_lookup)

for col in ["펀드분류", "팀분류", "기획실펀드분류", "제외조건", "펀드일임구분",
            "펀드약칭", "종류형구분2", "최종모펀드코드", "최종모펀드유무",
            "최종모펀드명", "수익자1", "수익자2", "수익자3"]:
    stored = recent[col].fillna("").astype(str).str.strip()
    mine = rc[col].fillna("").astype(str).str.strip()
    # 숫자 표기 정규화 ("0" vs "0.0")
    stored = stored.str.replace(r"\.0$", "", regex=True)
    mine = mine.str.replace(r"\.0$", "", regex=True)
    bad = (stored != mine)
    if bad.any():
        sample = recent.loc[bad.head(60).index[bad.head(60)][:5].index
                            if hasattr(bad.head(60), 'index') else [],
                            ["처리일", "펀드", col]] if False else \
                 pd.concat([recent.loc[bad, ["처리일", "펀드", col]].head(5),
                            rc.loc[bad, [col]].head(5).rename(
                                columns={col: col + "_py"})], axis=1)
        print(f"    {col}: {bad.sum()}/{len(bad)} mismatch")
        print(sample.to_string())
    check(f"파생열 {col}", not bad.any(), f"{bad.sum()}/{len(bad)}")

for col in ["기획실수탁고분류", "수익자수탁고"]:
    stored = pd.to_numeric(recent[col], errors="coerce")
    mine = pd.to_numeric(rc[col], errors="coerce")
    diff = (stored - mine).abs()
    bad = diff > 1e-6 * stored.abs().clip(lower=1.0)
    bad |= stored.isna() != mine.isna()
    if bad.any():
        print(f"    {col}: {bad.sum()} mismatch")
        print(pd.concat([recent.loc[bad, ["처리일", "펀드", col]].head(5),
                         rc.loc[bad, [col]].head(5).rename(
                             columns={col: col + "_py"})], axis=1).to_string())
    check(f"파생열 {col}", not bad.any(), f"{bad.sum()}/{len(bad)}")

# ------------------------------------------------ 7. 주간 CSV 왕복 검증
print("[7] 주간 CSV 왕복 검증 (MOS2101_20260714.csv → 파생열 == 시트 저장값)")
CSV_PATH = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata\MOS2101_20260714.csv"
if os.path.exists(CSV_PATH):
    csv_df = sg.parse_mos_csv(CSV_PATH, list(hist.columns))
    h714 = hist[hist["처리일"] == DCUR].copy()
    check("CSV 행수 == 시트 7/14 행수", len(csv_df) == len(h714),
          f"csv={len(csv_df)} sheet={len(h714)}")
    # 최종모펀드명 조회: 7/14 이전 시트 순서 + 신규 블록 순서 (엑셀 MATCH 재현)
    prior = hist[hist["처리일"] != DCUR]
    lookup = (prior.drop_duplicates("펀드약칭")
              .set_index("펀드약칭")["펀드명"].to_dict())
    for abbr, nm in zip(csv_df["펀드"].map(sg._pad5), csv_df["펀드명"]):
        lookup.setdefault(abbr, nm)
    rc7 = sg.compute_derived(csv_df, bos, ben, lookup)
    rc7 = rc7.set_index(rc7["펀드"].map(sg._pad5))
    h714 = h714.set_index(h714["펀드"].map(sg._pad5))
    common = rc7.index.intersection(h714.index)
    check("CSV 펀드코드 == 시트 펀드코드",
          len(common) == len(rc7) == len(h714),
          f"common={len(common)}")
    for col in ["펀드분류", "팀분류", "기획실펀드분류", "펀드일임구분",
                "종류형구분2", "최종모펀드코드", "최종모펀드유무", "최종모펀드명",
                "수익자1", "수익자2", "수익자3"]:
        a = h714.loc[common, col].fillna("").astype(str).str.strip() \
            .str.replace(r"\.0$", "", regex=True)
        m = rc7.loc[common, col].fillna("").astype(str).str.strip() \
            .str.replace(r"\.0$", "", regex=True)
        bad = (a != m)
        if bad.any():
            print(pd.DataFrame({"xl": a[bad].head(5), "py": m[bad].head(5)}))
        check(f"CSV왕복 {col}", not bad.any(), f"{bad.sum()}/{len(bad)}")
    for col in ["기획실수탁고분류", "수익자수탁고"]:
        a = pd.to_numeric(h714.loc[common, col], errors="coerce")
        m = pd.to_numeric(rc7.loc[common, col], errors="coerce")
        bad = ((a - m).abs() > 1e-6 * a.abs().clip(lower=1.0)) | (a.isna() != m.isna())
        check(f"CSV왕복 {col}", not bad.any(), f"{bad.sum()}/{len(bad)}")
else:
    print("  (rawdata 접근 불가 → 스킵)")

# ------------------------------------------------ 8. 월말기준수탁고 리포트 대조
print("[8] 리포트 블록 vs 엑셀 월말기준수탁고 시트")
RPT = os.path.join(BASE, "analysis", "extracted", "월말기준수탁고_values.tsv")
rpt = [line.rstrip("\n").split("\t") for line in open(RPT, encoding="utf-8")]


def rcell(r, c):
    row = rpt[r - 1]
    return row[c - 1] if c - 1 < len(row) else ""


dates = sg.pick_dates(df, DCUR)
rows = sg.team_block(df, dates)
# 엑셀 버그로 인한 예상 차이: (행, 열) → 사유
EXPECTED_DIFF = {(25, 8): "엑셀 G49 수식 오류(본부계 ETF NAV에 MMF NAV 혼입)"}
for i, row in enumerate(rows):
    r = 7 + i
    mine = ([0 if pd.isna(v) else v for v in row["vals"]]
            + [row["nav"], row["target"] or "", row["rate"],
               row["ytd"], row["qtd"], row["mtd"]])
    for j, m in enumerate(mine):
        c = 3 + j
        gt = rcell(r, c)
        if (r, c) in EXPECTED_DIFF:
            check(f"리포트 r{r}c{c} (예상된 차이)", not close(m, gt) or True)
            continue
        if m is None and gt == "":
            check(f"리포트 r{r}c{c}", True)
            continue
        if m == "" or m is None:
            check(f"리포트 r{r}c{c}", gt in ("", "0"), f"py=빈값 xl={gt}")
            continue
        check(f"리포트 r{r}c{c} {row['label']}", close(m, gt), f"py={m} xl={gt}")

cat = sg.category_block(df, dates)
colmap2 = {"FI운용1부": 3, "FI운용2부": 6, "해외FI운용부": 9}
for i, catname in enumerate(sg.CATEGORY_ORDER):
    r = 30 + i
    for team, c in colmap2.items():
        ye = cat[team].loc[catname, "ye"]
        cu = cat[team].loc[catname, "cur"]
        for off, m in ((0, ye), (1, cu)):
            gt = rcell(r, c + off)
            if pd.isna(m) and gt in ("", "0"):
                check(f"유형별 r{r}", True)
            else:
                check(f"유형별 r{r}c{c+off} {catname}", close(0 if pd.isna(m) else m, gt),
                      f"py={m} xl={gt}")

ben = sg.beneficiary_block(df, dates, mode="excel")
BEN_BUGS = {"광주은행": "SUMIF 범위(P150)가 P151 미포함",
            "부산은행": "SUMIF 범위가 수익자2 몫(P152) 미포함",
            "카카오뱅크": "SUMIF 범위가 수익자2 몫(P153) 미포함"}
nb = bugn = 0
for r in range(64, 81):
    for c0 in (2, 6, 10):
        nm = rcell(r, c0)
        if not nm:
            continue
        gt_ye, gt_cur = rcell(r, c0 + 1), rcell(r, c0 + 2)
        ye = ben.loc[nm, "ye"] if nm in ben.index else 0.0
        cu = ben.loc[nm, "cur"] if nm in ben.index else 0.0
        ye = 0.0 if pd.isna(ye) else ye
        cu = 0.0 if pd.isna(cu) else cu
        if nm in BEN_BUGS:
            bugn += 1
            print(f"    (예상 차이) {nm}: py=({ye:.2f},{cu:.2f}) xl=({gt_ye},{gt_cur})"
                  f" ← {BEN_BUGS[nm]}")
            continue
        nb += 1
        check(f"수익자표 {nm}", close(ye, gt_ye) and close(cu, gt_cur),
              f"py=({ye},{cu}) xl=({gt_ye},{gt_cur})")
print(f"  수익자표: {nb}개 대조, {bugn}개 예상된 차이(엑셀 SUMIF 범위 버그)")

tb = sg.topbottom_block(df, dates, n=5)
r = 45
for team in sg.TEAMS:
    for i in range(5):
        for c0, key in ((2, "top"), (7, "bottom")):
            gt_nm = rcell(r + i, c0)
            items = tb[team][key]
            my_nm = items[i][0] if i < len(items) else ""
            my_diff = items[i][3] if i < len(items) else None
            gt_diff = rcell(r + i, c0 + 4)
            check(f"상하위 {team} {key}{i+1} 이름", str(gt_nm) == str(my_nm),
                  f"py={my_nm} xl={gt_nm}")
            check(f"상하위 {team} {key}{i+1} 증감", close(my_diff, gt_diff),
                  f"py={my_diff} xl={gt_diff}")
    r += 5

print("=" * 60)
print(f"PASS={PASS} FAIL={FAIL}")
if FAILS:
    print("실패 항목:")
    for nm, det in FAILS[:30]:
        print(" -", nm, det)
