# -*- coding: utf-8 -*-
"""임의 기준일의 FI운용본부 수탁고 조회 (콘솔 출력).

리포트 xlsx 생성(make_report.py) 없이 숫자만 빠르게 확인할 때 쓴다.
집계 정의는 리포트와 동일하다:
  값   = 기획실수탁고분류 (ETF는 순자산(후), 그 외 원본액, 억원)
  필터 = 운용부서∈{FI운용본부,공동} · 모자구분∈{일반펀드,자펀드}
         · 종류형구분∈{일반펀드,판매펀드} · 제외조건=""
  본부 계 = 팀분류∈{FI운용1부, FI운용2부, 해외FI운용부} 합

사용:
  python aum.py                                  # 가장 최근 처리일
  python aum.py --date 2025-12-31                # 기준일 지정
  python aum.py --date 2025-12-31 --name ESG     # 펀드명 부분일치 집계 추가
  python aum.py --date 2025-12-31 --name ESG --list   # 해당 펀드 전건 나열
  python aum.py --dates                          # 사용 가능한 처리일 목록
"""
import argparse
import os
import sys

import pandas as pd

import sutakgo as sg

RAWDATA_DIR = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"


def load(csv_dir):
    df = (sg.build_dataset(csv_dir) if csv_dir else sg.load_history())
    df["처리일"] = pd.to_datetime(df["처리일"])
    for c in ("기획실수탁고분류", "순자산(후)", "원본액"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def resolve_date(df, want):
    """지정일이 처리일에 없으면 그 이전의 가장 가까운 처리일로 물러난다."""
    avail = sorted(pd.Timestamp(x) for x in df["처리일"].dropna().unique())
    if want is None:
        return avail[-1]
    want = pd.Timestamp(want)
    if want in avail:
        return want
    prev = [d for d in avail if d <= want]
    if not prev:
        sys.exit(f"오류: {want.date()} 이전 처리일이 없습니다.")
    print(f"! {want.date()} 는 처리일에 없어 직전 처리일 {prev[-1].date()} 로 대체합니다.")
    return prev[-1]


def base_frame(df, date):
    b = sg._base(df, [date])
    return b[b["팀분류"].isin(sg.TEAMS)]


def team_table(b):
    pv = b.pivot_table(index="팀분류", columns="펀드일임구분",
                       values="기획실수탁고분류", aggfunc="sum")
    pv = pv.reindex(index=sg.TEAMS,
                    columns=[c for c in sg.GUBUN_ORDER if c in pv.columns])
    pv["계"] = pv.sum(axis=1)
    pv.loc["본부 계"] = pv.sum()
    return pv


def main():
    ap = argparse.ArgumentParser(description="FI운용본부 수탁고 조회")
    ap.add_argument("--date", default=None, help="기준일 YYYY-MM-DD (기본: 최신 처리일)")
    ap.add_argument("--name", default=None,
                    help="펀드명 부분일치 필터 (대소문자 무시). 예: ESG")
    ap.add_argument("--list", action="store_true", help="필터 대상 펀드 전건 나열")
    ap.add_argument("--dates", action="store_true", help="사용 가능한 처리일 목록만 출력")
    ap.add_argument("--csv-dir", default=RAWDATA_DIR,
                    help="rawdata 폴더 (기본 K: 부서공유). 접근 불가하면 히스토리만 사용")
    ap.add_argument("--no-csv", action="store_true", help="rawdata를 보지 않고 히스토리만 사용")
    args = ap.parse_args()

    csv_dir = None if args.no_csv else (
        args.csv_dir if os.path.isdir(args.csv_dir) else None)
    if csv_dir is None and not args.no_csv:
        print(f"! rawdata 폴더 접근 불가 → 히스토리만 사용 ({args.csv_dir})")
    df = load(csv_dir)

    if args.dates:
        ds = sorted(pd.Timestamp(x).date() for x in df["처리일"].dropna().unique())
        print(f"처리일 {len(ds)}개: {ds[0]} ~ {ds[-1]}")
        for d in ds:
            print("  ", d)
        return

    date = resolve_date(df, args.date)
    b = base_frame(df, date)
    pd.set_option("display.width", 250, "display.max_colwidth", 55)

    print("=" * 74)
    print(f"FI운용본부 수탁고 — {date.date()} 기준 (억원)")
    print("=" * 74)
    print(team_table(b).round(1).to_string(na_rep="-"))
    total = b["기획실수탁고분류"].sum()
    print(f"\n본부 계 {total:,.1f} 억원 ({total/10000:,.2f} 조원), {len(b)}건")

    if not args.name:
        return

    key = args.name
    sub = b[b["펀드명"].astype(str).str.contains(key, case=False, na=False)]
    amt = sub["기획실수탁고분류"].sum()
    print()
    print("=" * 74)
    print(f"펀드명에 '{key}' 포함 (대소문자 무시)")
    print("=" * 74)
    print(f"합계 {amt:,.1f} 억원 — 본부 계의 {amt/total*100:.2f}%, {len(sub)}건\n")
    if sub.empty:
        return
    print("[팀별]")
    print(sub.groupby("팀분류")["기획실수탁고분류"]
          .agg(["sum", "count"]).round(1).to_string())
    print("\n[유형별]")
    print(sub.groupby("기획실펀드분류")["기획실수탁고분류"]
          .agg(["sum", "count"]).round(1).to_string())

    if args.list:
        print("\n[펀드 명세]")
        cols = ["펀드약칭", "펀드명", "팀분류", "펀드일임구분", "기획실펀드분류",
                "공모사모구분", "종류형구분", "기획실수탁고분류"]
        print(sub[cols].sort_values("기획실수탁고분류", ascending=False)
              .to_string(index=False, float_format=lambda v: f"{v:,.1f}"))

    # 이중계상 방지로 빠진 모펀드/운용펀드 행 — 필터가 제대로 걸렸는지 눈으로 확인용
    day = df[(df["처리일"] == date)
             & df["펀드명"].astype(str).str.contains(key, case=False, na=False)]
    excl = day[~day.index.isin(sub.index)]
    if len(excl):
        print("\n[참고] 필터에서 제외된 행 (모펀드·운용펀드 등 이중계상 방지)")
        print(excl[["펀드약칭", "펀드명", "운용부서", "팀분류", "모자구분",
                    "종류형구분", "제외조건", "기획실수탁고분류"]]
              .to_string(index=False, float_format=lambda v: f"{v:,.1f}"))


if __name__ == "__main__":
    main()
