# -*- coding: utf-8 -*-
"""FI운용본부 수탁고 리포트 엔진.

FI운용본부_수탁고양식_new.xlsb 의 업무(표_MOS2101 파생열 → PivotTables_EoM 피벗
→ 월말기준수탁고 리포트)를 파이썬으로 재현한다.

데이터 흐름
  data/mos2101_history.parquet : 엑셀 표_MOS2101 저장값 (과거 팀분류 기준 동결)
  rawdata MOS2101_YYYYMMDD.csv : 신규 일자 → 파생열 계산 후 append
  data/bos3218.csv             : 모자펀드 관계 (BOS3218 시트)
  data/수익자.csv               : 사모/일임 수익자 매핑 (수익자 시트)
"""
import os
import re
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")

RAW_COL_COUNT = 257          # CSV 261열 중 시트가 보존하는 원본 열 수
NEW_RULE_START = pd.Timestamp("2026-01-02")  # 현행 팀분류 수식 적용 시점

TEAMS = ["FI운용1부", "FI운용2부", "해외FI운용부"]
TEAM_LABEL = {"FI운용1부": "운용1팀", "FI운용2부": "운용2팀", "해외FI운용부": "해외FI운용팀"}
GUBUN_ORDER = ["펀드", "일임", "MMF", "ETF"]
CATEGORY_ORDER = ["채권", "채권파생", "채권혼합", "변액보험", "재간접",
                  "주식혼합", "ETF", "MMF", "재간접파생", "혼합자산"]

# 피벗 공통 필터 (pivot_page_filters.json 추출값)
F_BUSEO_WIDE = {"FI운용본부", "공동"}      # 팀별수탁고·MOS세부분류·수익자1·수익자2
F_BUSEO_FI = {"FI운용본부"}                # 팀별순자산·팀1·팀2·팀3
F_MOJA = {"일반펀드", "자펀드"}            # 모펀드 제외 (이중계상 방지)
F_JONGRYU = {"일반펀드", "판매펀드"}       # 운용펀드 제외 (수익자1·2 피벗은 미적용)

# 2026 목표 (월말기준수탁고 H열 하드코딩값)
TARGETS_2026 = {
    ("FI운용1부", "펀드"): 15000, ("FI운용1부", "ETF"): 10000,
    ("FI운용2부", "펀드"): 15000, ("FI운용2부", "ETF"): 8000,
    ("해외FI운용부", "펀드"): 5000, ("해외FI운용부", "ETF"): 15000,
}


# ---------------------------------------------------------------- 데이터 로드
def load_history():
    return pd.read_parquet(os.path.join(DATA, "mos2101_history.parquet"))


def load_bos(path=None):
    bos = pd.read_csv(path or os.path.join(DATA, "bos3218.csv"), dtype=str)
    for c in ("모펀드", "종류형모펀드", "종류형자펀드", "최종모펀드",
              "종류형모펀드_k", "종류형자펀드_l", "실질최종모펀드"):
        bos[c] = bos[c].map(lambda v: _pad5(v) if pd.notna(v) else v)
    return bos


def load_beneficiary(path=None):
    ben = pd.read_csv(path or os.path.join(DATA, "수익자.csv"), dtype=str)
    ben["열1"] = ben["열1"].map(_pad5)
    ben = ben[~ben["열1"].duplicated()]  # 엑셀 MATCH = 최초 일치
    return ben.set_index("열1")


def parse_mos_csv(path, columns):
    """rawdata의 MOS2101 CSV(261열) → 시트 원본영역(257열) DataFrame."""
    df = pd.read_csv(path, encoding="cp949", skiprows=3, header=None, dtype=str)
    for i in range(df.shape[1], RAW_COL_COUNT):  # 구형 CSV(열 부족) 패딩
        df[i] = None
    df = df.iloc[:, :RAW_COL_COUNT]
    df.columns = columns[:RAW_COL_COUNT]
    for c in ("영업일", "처리일"):
        df[c] = pd.to_datetime(df[c])
    for c in ("순자산(후)", "원본액", "순자산(전)"):
        df[c] = pd.to_numeric(df[c].str.replace(",", ""), errors="coerce")
    return df


# ---------------------------------------------------------------- 파생열 15종
def _pad5(code):
    """엑셀 TEXT([@펀드],"00000") 재현: 숫자형 코드만 5자리 0패딩.

    xlsb/CSV에서 숫자 셀로 읽힌 코드(151.0)도 '00151'로 정규화한다.
    """
    if code is None or (isinstance(code, float) and pd.isna(code)):
        return ""
    s = str(code)
    if re.fullmatch(r"\d+(\.0+)?", s):
        return str(int(float(s))).zfill(5)
    return s


def compute_derived(df, bos, ben, name_lookup):
    """신규 행에 대해 엑셀 표의 파생열 15개를 계산한다.

    name_lookup: 펀드약칭 → 펀드명 (전체 데이터 최초 등장 기준, 최종모펀드명용)
    """
    out = df.copy()
    code = out["펀드"].astype(str)
    fname = out["펀드명"].astype(str)

    is_etf = code.str.startswith("9") & fname.str.contains("ACE", regex=False)
    out["펀드분류"] = out["펀드유형"].where(~is_etf, "ETF")

    out["팀분류"] = out["운용팀"]
    out.loc[out["운용역"] == "김동주", "팀분류"] = "FI운용1부"
    mi = out["운용역"] == "이미연"
    out.loc[mi, "팀분류"] = out.loc[mi, "국내외구분"].map(
        lambda g: "FI운용2부" if g == "국내" else "해외FI운용부")

    out["기획실수탁고분류"] = out["원본액"].where(out["펀드분류"] != "ETF",
                                                 out["순자산(후)"]) / 1e8

    cat = out["펀드유형"].copy()
    cat = cat.where(~cat.eq("단기금융(MMF)"), "MMF")
    cat = cat.where(~is_etf, "ETF")
    cat = cat.where(code != "3JM54", "가상")
    out["기획실펀드분류"] = cat

    out["제외조건"] = (cat == "가상").map({True: "가상", False: ""})

    def gubun(row):
        if row["기획실펀드분류"] == "ETF":
            return "ETF"
        if row["기획실펀드분류"] == "MMF":
            return "MMF"
        if row["펀드구분"] == "투자신탁" and row["공모사모구분"] in ("공모", "사모"):
            return "펀드"
        if row["펀드구분"] == "투자일임":
            return "일임"
        return "FALSE"  # 엑셀 IF 미충족 시 FALSE 반환 재현
    out["펀드일임구분"] = out.apply(gubun, axis=1)

    out["펀드약칭"] = code.map(_pad5)

    # BOS3218 조회 (엑셀 MATCH = 최초 일치 행)
    b = bos.fillna("")
    first_k = ~b["종류형모펀드_k"].duplicated()
    first_l = ~b["종류형자펀드_l"].duplicated()
    k_map = b[first_k].set_index("종류형모펀드_k")  # 종류형모펀드 기준
    l_map = b[first_l].set_index("종류형자펀드_l")  # 종류형자펀드 기준
    k_set, l_set = set(k_map.index) - {""}, set(l_map.index) - {""}

    def bos_calc(abbr):
        """(종류형구분2, 최종모펀드코드)"""
        if abbr in l_set:
            r = l_map.loc[abbr]
            m, n, o = r["최종모부재"], r["종류형모부재"], r["종류형자부재"]
            if m == "O" and n == "O":
                return "종류형자펀드", abbr
            if m == "O":
                return "종류형자펀드", r["종류형모펀드_k"]
            if m == "X" and n == "X" and o == "X":
                return "종류형자펀드", r["최종모펀드"]
            return "종류형자펀드", ""
        if abbr in k_set:
            r = k_map.loc[abbr]
            if r["최종모부재"] == "X":
                return "종류형모펀드", r["최종모펀드"]
            return "종류형모펀드", abbr
        return "최종모펀드", abbr

    res = out["펀드약칭"].map(bos_calc)
    out["종류형구분2"] = res.map(lambda t: t[0])
    out["최종모펀드코드"] = res.map(lambda t: t[1])
    out["최종모펀드유무"] = (out["펀드약칭"] == out["최종모펀드코드"]).map(
        {True: "O", False: "X"})
    out["최종모펀드명"] = out["최종모펀드코드"].map(
        lambda c: name_lookup.get(c, "#N/A") if c else "#N/A")

    # 수익자 매핑 (사모만; 매핑행의 빈 칸은 엑셀 INDEX 결과인 0으로 재현)
    def benef(row, col):
        if row["공모사모구분"] != "사모":
            return ""
        abbr = row["펀드약칭"]
        if abbr not in ben.index:
            return ""
        v = ben.loc[abbr, col]
        if isinstance(v, pd.Series):  # 중복 매핑 시 최초 행 (엑셀 MATCH)
            v = v.iloc[0]
        return 0 if pd.isna(v) or v == "" else v

    for col in ("수익자1", "수익자2", "수익자3"):
        out[col] = out.apply(benef, axis=1, args=(col,))

    def benef_amt(row):
        b1, b2, b3 = row["수익자1"], row["수익자2"], row["수익자3"]
        amt = row["기획실수탁고분류"]
        if (b1 == "" and b2 == "" and b3 == "") or (b2 == 0 and b3 == 0):
            return amt          # 수익자 0~1명 → 전액
        if b3 == 0:
            return amt / 2      # 2명 → 1/2
        return amt / 3          # 3명 → 1/3
    out["수익자수탁고"] = out.apply(benef_amt, axis=1)
    return out


def build_dataset(csv_dir=None, verbose=True):
    """히스토리 + 신규 CSV 병합 데이터셋. 신규 일자만 파생열 계산해 append."""
    hist = load_history()
    frames = [hist]
    if csv_dir:
        have = set(hist["처리일"].dropna().unique())
        bos, ben = load_bos(), load_beneficiary()
        files = sorted(f for f in os.listdir(csv_dir)
                       if re.fullmatch(r"MOS2101_\d{8}\.csv", f))
        # 최종모펀드명 조회용: 전체 시트 최초 등장 순 (엑셀 MATCH 전열 검색 재현)
        name_lookup = (hist.drop_duplicates("펀드약칭")
                       .set_index("펀드약칭")["펀드명"].to_dict())
        for f in files:
            try:
                new = parse_mos_csv(os.path.join(csv_dir, f), list(hist.columns))
            except (pd.errors.EmptyDataError, pd.errors.ParserError) as e:
                if verbose:
                    print(f"  ! {f} 스킵: {e}")
                continue
            new = new[~new["처리일"].isin(have)]
            if new.empty:
                continue
            for abbr, nm in zip(new["펀드"].map(_pad5), new["펀드명"]):
                name_lookup.setdefault(abbr, nm)
            new = compute_derived(new, bos, ben, name_lookup)
            frames.append(new)
            have |= set(new["처리일"].unique())
            if verbose:
                d = sorted(pd.to_datetime(x).date() for x in new["처리일"].unique())
                print(f"  + {f}: {len(new)} rows, dates {d}")
    df = pd.concat(frames, ignore_index=True)
    df["처리일"] = pd.to_datetime(df["처리일"])
    for c in ("기획실수탁고분류", "수익자수탁고", "순자산(후)", "원본액"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


# ---------------------------------------------------------------- 피벗 재현
def _base(df, dates, buseo=F_BUSEO_WIDE, moja=F_MOJA, jongryu=F_JONGRYU):
    m = (df["처리일"].isin(dates)
         & df["운용부서"].isin(buseo)
         & df["제외조건"].fillna("").eq(""))
    if moja is not None:
        m &= df["모자구분"].isin(moja)
    if jongryu is not None:
        m &= df["종류형구분"].isin(jongryu)
    return df[m]


def pivot_team(df, dates):
    """팀별수탁고: 팀분류×펀드일임구분 × 처리일, 합계 기획실수탁고분류 (억원)."""
    d = _base(df, dates)
    g = (d.groupby(["팀분류", "펀드일임구분", "처리일"])["기획실수탁고분류"]
         .sum().unstack("처리일").reindex(columns=dates))
    return g


def pivot_team_nav(df, date):
    """팀별순자산: 순자산(후) 합계, 운용부서=FI운용본부 단독."""
    d = _base(df, [date], buseo=F_BUSEO_FI)
    return d.groupby(["팀분류", "펀드일임구분"])["순자산(후)"].sum()


def pivot_category(df, dates):
    """MOS세부분류: 기획실펀드분류 × (팀분류×처리일)."""
    d = _base(df, dates)
    g = (d.groupby(["기획실펀드분류", "팀분류", "처리일"])["기획실수탁고분류"]
         .sum().unstack(["팀분류", "처리일"]))
    return g


def pivot_fund_diff(df, team, date_base, date_cur, by="최종모펀드명"):
    """팀1/2/3: 최종모펀드별 수탁고와 증감 (최종모펀드유무=O, FI운용본부).

    by: 묶음 키. 엑셀 피벗은 최종모펀드명(코드 최초 등장 이름)으로 묶는다.
    """
    d = _base(df, [date_base, date_cur], buseo=F_BUSEO_FI,
              moja=None, jongryu=None)  # 팀1~3 피벗은 모자·종류형 필터 없음
    d = d[(d["팀분류"] == team) & (d["최종모펀드유무"] == "O")]
    g = (d.groupby([by, "처리일"])["기획실수탁고분류"]
         .sum().unstack("처리일").reindex(columns=[date_base, date_cur]))
    g.columns = ["base", "cur"]
    g["diff"] = g["cur"].fillna(0) - g["base"].fillna(0)
    return g.sort_index()  # 피벗 행 순서(가나다) 재현


def pivot_beneficiary(df, dates, col):
    """수익자1/수익자2: 수익자수탁고 합계 (종류형구분 필터 없음)."""
    d = _base(df, dates, jongryu=None)
    s = d[col].astype(str).str.strip()
    named = (d[col].notna() & ~s.isin(["", "nan", "None"])
             & ~s.str.fullmatch(r"0(\.0+)?"))  # 숫자 0("0","0.0")은 미지정
    d = d[named]
    amt = pd.to_numeric(d["수익자수탁고"], errors="coerce")
    g = (d.assign(_amt=amt).groupby([col, "처리일"])["_amt"]
         .sum().unstack("처리일").reindex(columns=dates))
    return g


# ---------------------------------------------------------------- 리포트 값
def clean_fund_name(name):
    return str(name).replace("한국투자", "").replace("증권투자신탁", "")


def latest_fund_names(df, upto):
    """펀드약칭 → upto(포함) 이전 가장 최근 처리일의 펀드명 (개명 반영)."""
    d = df[df["처리일"] <= upto].sort_values("처리일", kind="stable")
    return d.groupby("펀드약칭")["펀드명"].last().astype(str).str.strip()


def pick_dates(df, target):
    """리포트 기준일 5종: 2024YE, 2025YE, 직전분기말, 직전월말, 최신일."""
    target = pd.Timestamp(target)
    avail = sorted(df["처리일"].dropna().unique())
    avail = [pd.Timestamp(x) for x in avail if pd.Timestamp(x) <= target]

    def last_on_or_before(d):
        c = [x for x in avail if x <= d]
        return c[-1] if c else None

    ye24 = last_on_or_before(pd.Timestamp("2024-12-31"))
    ye25 = last_on_or_before(pd.Timestamp("2025-12-31"))
    qe = target - pd.offsets.QuarterEnd(1)   # 직전 분기말 (분기말 당일이면 이전 분기말)
    me = target - pd.offsets.MonthEnd(1)     # 직전 월말 (월말 당일이면 이전 월말)
    return {"ye24": ye24, "ye25": ye25, "quarter": last_on_or_before(qe),
            "month": last_on_or_before(me), "current": target}


def team_block(df, dates):
    """부서별 수탁고 표 (월말기준수탁고 6~25행) 값 생성."""
    dts = [dates["ye24"], dates["ye25"], dates["quarter"], dates["month"],
           dates["current"]]
    pv = pivot_team(df, dts)
    nav = pivot_team_nav(df, dates["current"]) / 1e8

    rows = []
    total = {c: 0.0 for c in range(5)}
    total_by_gubun = {}
    for team in TEAMS:
        sub = pv.loc[team] if team in pv.index.get_level_values(0) else pd.DataFrame()
        team_vals = sub.sum()
        team_nav = nav.loc[team].sum() if team in nav.index.get_level_values(0) else 0.0
        target = sum(v for (t, g), v in TARGETS_2026.items() if t == team)
        rows.append({"label": TEAM_LABEL[team], "indent": 0, "team": team,
                     "gubun": None, "vals": list(team_vals), "nav": team_nav,
                     "target": target})
        for g in GUBUN_ORDER:
            if g not in sub.index:
                continue
            gv = sub.loc[g]
            rows.append({"label": g, "indent": 1, "team": team, "gubun": g,
                         "vals": list(gv), "nav": nav.get((team, g), 0.0),
                         "target": TARGETS_2026.get((team, g)),
                         })
            total_by_gubun.setdefault(g, [0.0] * 5)
            total_by_gubun[g] = [a + (0 if pd.isna(b) else b)
                                 for a, b in zip(total_by_gubun[g], gv)]
        for i, v in enumerate(team_vals):
            total[i] += 0 if pd.isna(v) else v

    nav_total = sum(r["nav"] for r in rows if r["gubun"] is None)
    rows_total = [{"label": "본부 계", "indent": 0, "team": None, "gubun": None,
                   "vals": [total[i] for i in range(5)], "nav": nav_total,
                   "target": sum(TARGETS_2026.values())}]
    for g in GUBUN_ORDER:
        if g not in total_by_gubun:
            continue
        gnav = sum(r["nav"] for r in rows if r["gubun"] == g)
        gtarget = sum(v for (t, gg), v in TARGETS_2026.items() if gg == g) or None
        rows_total.append({"label": g, "indent": 1, "team": None, "gubun": g,
                           "vals": total_by_gubun[g], "nav": gnav,
                           "target": gtarget})
    rows += rows_total

    # 파생 지표: 달성률·YTD·QTD·MTD
    for r in rows:
        v = [0 if pd.isna(x) else x for x in r["vals"]]
        r["ytd"] = v[4] - v[1]
        r["qtd"] = v[4] - v[2]
        r["mtd"] = v[4] - v[3]
    # 달성률(엑셀 규칙 재현): 팀행 = 팀YTD/팀목표,
    # 펀드행 = (펀드+일임+MMF YTD)/펀드목표, ETF행 = ETF YTD/ETF목표
    by_key = {(r["team"], r["gubun"]): r for r in rows}
    for r in rows:
        r["rate"] = None
        if not r["target"]:
            continue
        if r["gubun"] is None:
            r["rate"] = r["ytd"] / r["target"]
        elif r["gubun"] == "펀드":
            ytd = sum(by_key[(r["team"], g)]["ytd"] for g in ("펀드", "일임", "MMF")
                      if (r["team"], g) in by_key)
            r["rate"] = ytd / r["target"]
        else:
            r["rate"] = r["ytd"] / r["target"]
    return rows


def category_block(df, dates):
    """유형별 수탁고 표 (30~40행): 부서별 (2025YE, 최신) 값."""
    dts = [dates["ye25"], dates["current"]]
    pv = pivot_category(df, dts)
    out = {}
    for team in TEAMS:
        cols = []
        for dt in dts:
            if (team, dt) in pv.columns:
                cols.append(pv[(team, dt)])
            else:
                cols.append(pd.Series(dtype=float))
        tbl = pd.concat(cols, axis=1)
        tbl.columns = ["ye", "cur"]
        out[team] = tbl.reindex(CATEGORY_ORDER)
    return out


def topbottom_block(df, dates, n=5, latest_names=True):
    """YTD 증감 상/하위 표 (45~59행): 팀별 상위 n·하위 n.

    latest_names: 코드로 묶고 기준일 시점 최신 펀드명으로 표시 (개명 반영).
      False면 엑셀 원본처럼 코드 최초 등장 이름 — verify.py 재현용.
    """
    names = latest_fund_names(df, dates["current"]) if latest_names else None
    res = {}
    for team in TEAMS:
        if names is None:
            g = pivot_fund_diff(df, team, dates["ye25"], dates["current"])
        else:
            g = pivot_fund_diff(df, team, dates["ye25"], dates["current"],
                                by="최종모펀드코드")
            g.index = g.index.map(names)
            g = g.sort_index()  # 이름 가나다순 (동률 시 순서 기준)
        g = g[g["diff"].notna()]
        top = g.sort_values("diff", ascending=False, kind="stable").head(n)
        bot = g.sort_values("diff", ascending=True, kind="stable").head(n)
        res[team] = {
            "top": [(clean_fund_name(i), r["base"], r["cur"], r["diff"])
                    for i, r in top.iterrows()],
            "bottom": [(clean_fund_name(i), r["base"], r["cur"], r["diff"])
                       for i, r in bot.iterrows()],
        }
    return res


def beneficiary_alloc(df, dates, ben=None):
    """수익자별 배분액 — 현행 매핑을 전 일자에 재적용 + 가중 배분.

    - 매핑 시트의 수익자N수탁고 열을 가중치로 사용해 기획실수탁고분류를 배분
      (예: 08Q58 삼성카드 300/저축은행 1). 지정 수익자 전원의 가중치가 있을
      때만 가중, 아니면 균등 — 엑셀의 무조건 1/n 균등분할 왜곡을 제거.
    - 시트 저장값이 아닌 현행 매핑 기준이므로 이명 통일·클래스 이관이
      과거 일자 열에도 일관되게 반영된다. 수익자3 몫도 포함(엑셀 표는 누락).
    """
    if ben is None:
        ben = load_beneficiary()
    d = _base(df, dates, jongryu=None)
    d = d[d["공모사모구분"] == "사모"]
    d = d[["펀드약칭", "처리일", "기획실수탁고분류"]]
    m = d.merge(ben.reset_index(), left_on="펀드약칭", right_on="열1", how="inner")
    recs = []
    for _, r in m.iterrows():
        amt = pd.to_numeric(r["기획실수탁고분류"], errors="coerce")
        if pd.isna(amt):
            continue
        slots = []
        for i in (1, 2, 3):
            nm = str(r.get(f"수익자{i}") or "").strip()
            if nm in ("", "0", "nan", "None"):
                continue
            w = pd.to_numeric(r.get(f"수익자{i}수탁고"), errors="coerce")
            slots.append((nm, w))
        if not slots:
            continue
        ws = [w for _, w in slots]
        if any(pd.isna(w) or w <= 0 for w in ws):
            ws = [1.0] * len(slots)          # 가중치 불완전 → 균등
        tot = sum(ws)
        for (nm, _), w in zip(slots, ws):
            recs.append((nm, r["처리일"], amt * w / tot))
    out = pd.DataFrame(recs, columns=["수익자", "처리일", "금액"])
    g = out.groupby(["수익자", "처리일"])["금액"].sum().unstack("처리일")
    return g.reindex(columns=dates)


def beneficiary_block(df, dates, ben=None, mode="live"):
    """수익자 분류 표 (61~80행).

    mode="live"  : 현행 매핑 재적용 + 가중 배분 (운영 기본)
    mode="excel" : 시트 저장값 + 균등분할, 수익자1+2만 (엑셀 재현·검증용)
    """
    dts = [dates["ye25"], dates["current"]]
    if mode == "excel":
        b1 = pivot_beneficiary(df, dts, "수익자1")
        b2 = pivot_beneficiary(df, dts, "수익자2")
        allb = pd.concat([b1, b2]).groupby(level=0).sum()
    else:
        allb = beneficiary_alloc(df, dts, ben)
    allb.columns = ["ye", "cur"]
    return allb
