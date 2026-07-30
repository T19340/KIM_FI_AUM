# -*- coding: utf-8 -*-
"""월말기준수탁고 리포트(xlsx) 생성 — 양식 파일에 Excel COM으로 값 채움.

template/FI운용본부수탁고_양식.xlsx 사본에 계산값만 써넣는다. COM을 쓰는 이유:
openpyxl 저장은 데이터 막대의 x14 확장(0축·음수 붉은 역방향 막대)을 벗겨낸다.
양식·산출물에는 절대 openpyxl로 저장하지 않는다 (읽기는 무방).

수익자 명단·서식 관리는 양식 파일에서 직접 한다 (수익자 표 B/F/J열).

사용:
  python make_report.py                        # rawdata 최신 처리일 기준
  python make_report.py --date 2026-07-14     # 기준일 지정
"""
import argparse
import os
import shutil
import pandas as pd
import win32com.client

import sutakgo as sg

RAWDATA_DIR = r"K:\부서 공유\FI운용1부\10. 개인별폴더\김현수\rawdata"
TEMPLATE = os.path.join(sg.BASE, "template", "FI운용본부수탁고_양식.xlsx")
EXPECTED_TEAM_ROWS = 19   # 부서별 표(7~25행)의 양식 행수


def qlabel(d):
    return f"{(d.month - 1) // 3 + 1}Q{d.year}"


def mlabel(d):
    return f"{d.month}M{d.year}"


def _num(v):
    return None if v is None or (isinstance(v, float) and pd.isna(v)) else float(v)


def _xlserial(ts):
    """날짜 → 엑셀 일련번호(정수). pywin32의 datetime 변환은 로컬→UTC 보정이
    끼어들어 9시간이 밀리므로 datetime 객체를 COM에 직접 넘기지 않는다."""
    ts = pd.Timestamp(ts)
    if ts.tzinfo is not None:
        ts = ts.tz_localize(None)
    return int((ts.normalize() - pd.Timestamp("1899-12-30")).days)


def _mapped_codes(ben):
    """수익자1~3 중 하나라도 유효값이 있는 펀드코드 — 실제로 배분되는 매핑."""
    out = set()
    for code, r in ben.iterrows():
        for i in (1, 2, 3):
            if str(r.get(f"수익자{i}") or "").strip() not in ("", "0", "nan", "None"):
                out.add(code)
                break
    return out


def check_beneficiary_coverage(df, dates, ben=None):
    """수익자 표의 침묵 실패 2종을 리포트 두 열(2025YE·현재)에서 점검한다.

    ① 매핑 없는 사모펀드 — 부서별 표에는 들어가는데 수익자 표에서 조용히 빠진다.
      (2026-07-28 리포트에서 신규 5JM61이 이렇게 1,468억 누락됐다.)
    ② 모펀드(운용펀드)에 붙은 매핑 — 클래스도 매핑돼 있으면 수익자 표만 이중계상된다.

    끝에 대조 한 줄을 남긴다: 수익자 배분 합 == 사모 계상 총액이면 둘 다 0이다.
    """
    if ben is None:
        ben = sg.load_beneficiary()
    mapped = _mapped_codes(ben)
    for key, label in (("ye25", "2025YE"), ("current", "현재")):
        d = dates.get(key)
        if d is None:
            continue
        base = sg._base(df, [d])
        base = base[base["팀분류"].isin(sg.TEAMS)
                    & (base["공모사모구분"] == "사모")].copy()
        base["_amt"] = pd.to_numeric(base["기획실수탁고분류"], errors="coerce")
        total = base["_amt"].sum()

        un = base[~base["펀드약칭"].isin(mapped)]
        if len(un):
            print(f"경고: [{label}] 수익자 매핑 없는 사모펀드 {len(un)}건 "
                  f"{un['_amt'].sum():,.1f}억 — 수익자 표에서 빠짐 "
                  "(data/수익자.csv에 클래스 단위로 추가)")
            for _, r in un.sort_values("_amt", ascending=False).iterrows():
                print(f"   {r['펀드약칭']} {str(r['펀드명'])[:40]} "
                      f"{r['팀분류']} {r['_amt']:,.1f}억")

        mo = sg._base(df, [d], jongryu=None)
        mo = mo[(mo["공모사모구분"] == "사모") & (mo["종류형구분"] == "운용펀드")
                & mo["펀드약칭"].isin(mapped)].copy()
        if len(mo):
            mo["_amt"] = pd.to_numeric(mo["기획실수탁고분류"], errors="coerce")
            print(f"경고: [{label}] 모펀드(운용펀드)에 붙은 매핑 {len(mo)}건 "
                  f"{mo['_amt'].sum():,.1f}억 — 클래스도 매핑돼 있으면 이중계상 "
                  "(매핑은 클래스 단위, 모는 비움)")
            for _, r in mo.iterrows():
                print(f"   {r['펀드약칭']} {str(r['펀드명'])[:40]} {r['_amt']:,.1f}억")

        alloc = sg.beneficiary_alloc(df, [d], ben).iloc[:, 0].sum()
        diff = alloc - total
        verdict = "일치" if abs(diff) < 0.05 else f"차이 {diff:+,.1f}억 (위 경고 확인)"
        print(f"  대조 [{label}] 수익자 배분 {alloc:,.1f}억 / "
              f"사모 계상 {total:,.1f}억 → {verdict}")


def fill_template(df, dates, out_path):
    cur = _xlserial(dates["current"])
    ye25 = _xlserial(dates["ye25"])
    rows = sg.team_block(df, dates)
    cat = sg.category_block(df, dates)
    tb = sg.topbottom_block(df, dates, n=5)
    ben = sg.beneficiary_block(df, dates)

    shutil.copy2(TEMPLATE, out_path)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(os.path.abspath(out_path), UpdateLinks=0)
    try:
        ws = wb.Worksheets("월말기준수탁고")
        ws.Range("L3").Value = cur

        # ---- 1. 부서별 수탁고 (6~25행) ----
        ws.Range("E6").Value = qlabel(dates["quarter"])
        ws.Range("F6").Value = mlabel(dates["month"])
        ws.Range("G6").Value = cur
        if len(rows) != EXPECTED_TEAM_ROWS:
            print(f"경고: 부서별 표 행수 {len(rows)} != 양식 {EXPECTED_TEAM_ROWS} — "
                  "구분 구성이 바뀜. 양식 행 구성을 확인할 것.")
        tpl_labels = [str(v[0] or "").strip() for v in ws.Range("B7:B25").Value]
        arr = []
        for i in range(EXPECTED_TEAM_ROWS):
            if i >= len(rows):
                arr.append([None] * 11)
                continue
            row = rows[i]
            if tpl_labels[i] != row["label"]:
                print(f"경고: {7+i}행 라벨 불일치 "
                      f"(양식 '{tpl_labels[i]}' vs 계산 '{row['label']}')")
            arr.append([_num(v) for v in row["vals"]]
                       + [_num(row["nav"]), row["target"] or None, _num(row["rate"]),
                          _num(row["ytd"]), _num(row["qtd"]), _num(row["mtd"])])
        ws.Range("C7:M25").Value = arr

        # ---- 2. 유형별 수탁고 (29~40행) ----
        for a in ("C29", "F29", "I29"):
            ws.Range(a).Value = ye25
        for a in ("D29", "G29", "J29"):
            ws.Range(a).Value = cur
        block = []
        for catname in sg.CATEGORY_ORDER:
            r = []
            tot_ytd = tot_cur = 0.0
            any_val = False
            for team in sg.TEAMS:
                ye = cat[team].loc[catname, "ye"]
                cu = cat[team].loc[catname, "cur"]
                ye_v = 0.0 if pd.isna(ye) else ye
                cu_v = 0.0 if pd.isna(cu) else cu
                r += [_num(ye), _num(cu),
                      None if (pd.isna(ye) and pd.isna(cu)) else cu_v - ye_v]
                tot_ytd += cu_v - ye_v
                tot_cur += cu_v
                any_val |= not (pd.isna(ye) and pd.isna(cu))
            r += [tot_ytd if any_val else None, tot_cur if any_val else None]
            block.append(r)
        total = []
        for team in sg.TEAMS:
            ye_s, cu_s = cat[team]["ye"].sum(), cat[team]["cur"].sum()
            total += [ye_s, cu_s, cu_s - ye_s]
        total += [sum(cat[t]["cur"].sum() - cat[t]["ye"].sum() for t in sg.TEAMS),
                  sum(cat[t]["cur"].sum() for t in sg.TEAMS)]
        block.append(total)
        ws.Range("C30:M40").Value = block

        # ---- 3. YTD 증감 상/하위 5 (44~59행) ----
        for a in ("D44", "I44"):
            ws.Range(a).Value = ye25
        for a in ("E44", "J44"):
            ws.Range(a).Value = cur
        b3 = []
        for team in sg.TEAMS:
            for i in range(5):
                r = []
                for key in ("top", "bottom"):
                    items = tb[team][key]
                    if i < len(items):
                        name, base, curv, diff = items[i]
                        r += [name, None, _num(base), _num(curv), _num(diff)]
                    else:
                        r += [None] * 5
                b3.append(r)
        ws.Range("B45:K59").Value = b3

        # ---- 4. 사모/일임 수익자분류 (63행~) — 명단은 양식이 기준 ----
        for a in ("C63", "G63", "K63"):
            ws.Range(a).Value = ye25
        for a in ("D63", "H63", "L63"):
            ws.Range(a).Value = cur
        used = set()
        for name_c, v1 in (("B", "C"), ("F", "G"), ("J", "K")):
            raw = ws.Range(f"{name_c}64:{name_c}120").Value
            names = []
            for (v,) in raw:
                if v is None or str(v).strip() == "":
                    break
                names.append(str(v).strip())
            if not names:
                continue
            vals = []
            for nm in names:
                ye = ben.loc[nm, "ye"] if nm in ben.index else 0.0
                cu = ben.loc[nm, "cur"] if nm in ben.index else 0.0
                ye = 0.0 if pd.isna(ye) else ye
                cu = 0.0 if pd.isna(cu) else cu
                vals.append([ye, cu, cu - ye])
                used.add(nm)
            v3 = chr(ord(v1) + 2)
            ws.Range(f"{v1}64:{v3}{63 + len(names)}").Value = vals
        missing = [n for n in ben.index
                   if n not in used and (abs(ben.loc[n].fillna(0)) > 0.5).any()]
        if missing:
            print("경고: 수익자 표 명단에 없는 수익자 (양식에 행 추가 필요):")
            for n in missing:
                v = ben.loc[n].fillna(0)
                print(f"   {n}: 25YE {v.iloc[0]:,.1f} / 현재 {v.iloc[1]:,.1f}")

        wb.Save()
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None, help="기준일 YYYY-MM-DD (기본: 최신 처리일)")
    ap.add_argument("--csv-dir", default=RAWDATA_DIR)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    csv_dir = args.csv_dir if os.path.isdir(args.csv_dir) else None
    if csv_dir is None:
        print(f"경고: rawdata 폴더 접근 불가 → 히스토리만 사용 ({args.csv_dir})")
    df = sg.build_dataset(csv_dir)
    target = pd.Timestamp(args.date) if args.date else df["처리일"].max()
    dates = sg.pick_dates(df, target)
    print("기준일:", {k: (v.date() if v is not None else None)
                     for k, v in dates.items()})
    check_beneficiary_coverage(df, dates)

    out = args.out or os.path.join(
        sg.BASE, "output", f"FI운용본부수탁고_{target.strftime('%Y%m%d')}.xlsx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    fill_template(df, dates, out)
    print("저장:", out)


if __name__ == "__main__":
    main()
