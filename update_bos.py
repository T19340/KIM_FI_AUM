# -*- coding: utf-8 -*-
"""BOS3218 원본 출력(xlsx, 9열)으로 data/bos3218.csv 갱신.

원본 BOS 3218 화면 출력은 9열(순번~종류형자펀드명)뿐이고, 워크북의 J~Q열은
시트 내 파생이다. 현행 데이터 전수에서 부재 플래그가 전부 "X"이므로 파생 규칙은:
  B(모펀드)·F(종류형모펀드) fill-down 후  J=모펀드, K=종류형모펀드, L=종류형자펀드,
  M=N=O="X", P(실질최종모펀드)=모펀드, Q(최종모펀드명)=모펀드명(fill-down)

사용: python update_bos.py <BOS3218_YYYYMMDD.xlsx>
"""
import os
import shutil
import sys
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sutakgo import _pad5

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "data", "bos3218.csv")
COLS = ["순번", "모펀드", "모펀드명", "설정해지구분", "투자비율",
        "종류형모펀드", "종류형모펀드명", "종류형자펀드", "종류형자펀드명",
        "최종모펀드", "종류형모펀드_k", "종류형자펀드_l",
        "최종모부재", "종류형모부재", "종류형자부재", "실질최종모펀드", "최종모펀드명"]


def derive(raw, prefilled=False):
    """원본 9열 → 17열 파생. raw 열순서는 COLS[:9]와 동일해야 한다.

    BOS 출력의 모펀드(B)열 규약: None=위 블록 계속, ''(빈 문자열)=모펀드 부재
    구역 시작 마커. fill-down 시 ''가 그대로 전파되어 구역 전체를 표시한다.
    prefilled=True 면 이미 fill-down된 데이터(기존 CSV)로 재파생만 한다.
    """
    b = raw.copy()
    b.columns = COLS[:len(b.columns)]
    notblank = b[COLS[1:9]].notna() & b[COLS[1:9]].astype(str).ne("")
    b = b[notblank.any(axis=1) | (b["모펀드"] == "")]  # 완전 빈 행 제거(마커 행 보존)
    if not prefilled:
        b["모펀드"] = b["모펀드"].ffill()        # ''(부재 마커)도 값으로 전파됨
        b["모펀드명"] = b["모펀드명"].ffill()
        b["종류형모펀드"] = b["종류형모펀드"].ffill()
    for c in ("모펀드", "종류형모펀드", "종류형자펀드"):
        b[c] = b[c].map(lambda v: _pad5(v) if pd.notna(v) and v != "" else v)
    no_mo = b["모펀드"].isna() | b["모펀드"].astype(str).str.strip().eq("")
    b["최종모펀드"] = b["모펀드"].where(~no_mo, b["종류형모펀드"])
    b["종류형모펀드_k"] = b["종류형모펀드"]
    b["종류형자펀드_l"] = b["종류형자펀드"]
    b["최종모부재"] = "X"
    b["종류형모부재"] = "X"
    b["종류형자부재"] = "X"
    b["실질최종모펀드"] = b["최종모펀드"]
    b["최종모펀드명"] = b["모펀드명"].where(~no_mo, b["종류형모펀드명"])
    return b[COLS]


def selftest():
    """기존 data/bos3218.csv(워크북 파생값)로 파생 규칙 검증.

    기존 CSV는 이미 fill-down 완료본이며 모펀드 부재 구역은 빈 값으로 저장되어
    있다 (읽으면 NaN) → prefilled=True 로 재파생만 대조.
    """
    old = pd.read_csv(OUT, dtype=str)
    rd = derive(old[COLS[:9]], prefilled=True)
    ok = True
    for c in ["최종모펀드", "종류형모펀드_k", "종류형자펀드_l", "실질최종모펀드"]:
        a = old[c].map(lambda v: _pad5(v) if pd.notna(v) else "")
        m = rd[c].map(lambda v: _pad5(v) if pd.notna(v) else "")
        bad = (a != m).sum()
        print(f"  selftest {c}: 불일치 {bad}/{len(old)}")
        ok &= bad == 0
    return ok


def _read_bos_drm(src):
    """DRM 래핑 출력본은 zip 구조가 아니라 openpyxl이 못 연다(BadZipFile).

    Excel COM으로 값만 통으로 읽는다 — COM도 ''(부재 마커)와 None(빈 셀)을
    구분해 돌려주는 것을 실측 확인(2026-08-06). Excel의 SaveAs는 DRM이 다시
    래핑하므로 변환 저장 경로는 쓸 수 없다.
    """
    from win32com.client import DispatchEx
    app = DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(os.path.abspath(src), ReadOnly=True)
        ws = wb.Worksheets(1)
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        vals = ws.Range(ws.Cells(1, 1), ws.Cells(last, 9)).Value
        wb.Close(False)
    finally:
        app.Quit()

    def norm(v):
        if v is None:
            return None
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    rows = [tuple(norm(v) for v in r) for r in vals]
    return pd.DataFrame(rows[1:], columns=[str(c).replace("\n", "")
                                           for c in rows[0]], dtype=object)


def read_bos_xlsx(src):
    """BOS3218 xlsx를 openpyxl로 직접 읽는다.

    pd.read_excel 은 '모펀드 부재 시작' 마커인 빈 문자열('') 셀을 NaN으로
    바꿔버려 부재 구역 전체가 앞 블록 모펀드를 상속하는 치명적 오류를 낳는다.
    openpyxl은 ''(마커)와 None(위 블록 계속)을 구분해 보존한다.
    DRM 래핑 출력본이면 Excel COM 경로로 자동 전환한다.
    """
    import zipfile

    import openpyxl
    try:
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    except zipfile.BadZipFile:
        print("  (DRM 래핑 감지 — Excel COM으로 읽기)")
        return _read_bos_drm(src)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))

    def norm(v):
        # None(빈 셀)은 None 유지, ''(마커) 보존, 숫자는 문자열화
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return str(v)
    raw = pd.DataFrame(rows[1:], columns=[str(c).replace("\n", "")
                                          for c in rows[0]], dtype=object)
    return raw.map(norm)


def fund_root_map(bos):
    """펀드코드 → 최종모펀드 (K·L 최초 등장 기준)."""
    m = {}
    for col in ("종류형자펀드_l", "종류형모펀드_k"):
        for code, j in zip(bos[col], bos["최종모펀드"]):
            if pd.notna(code) and code != "":
                m.setdefault(code, j)
    return m


def main():
    if len(sys.argv) < 2 or sys.argv[1] == "--selftest":
        print("파생 규칙 자가검증 (기존 데이터 대비):")
        sys.exit(0 if selftest() else 1)
    src = sys.argv[1]
    print("파생 규칙 자가검증:")
    if not selftest():
        print("자가검증 실패 — 갱신 중단"); sys.exit(1)
    raw = read_bos_xlsx(src)
    new = derive(raw)
    n_marker = (new["모펀드"].fillna("").astype(str).str.strip() == "").sum()
    print(f"모펀드 부재 구역: {n_marker}행")
    if n_marker == 0:
        print("경고: 부재 구역 0행 — 마커 소실 의심. 갱신 중단"); sys.exit(1)

    # 이전본 대비 최종모펀드 변경 diff (회귀 감지)
    old = pd.read_csv(OUT, dtype=str)
    om, nm = fund_root_map(old.fillna("")), fund_root_map(new.fillna(""))
    common = set(om) & set(nm)
    changed = sorted(c for c in common if str(om[c]) != str(nm[c]))
    print(f"최종모펀드 변경: {len(changed)}건 / 공통 {len(common)}펀드 "
          f"(신규 {len(set(nm)-set(om))}, 소멸 {len(set(om)-set(nm))})")
    for c in changed[:10]:
        print(f"   {c}: {om[c]} → {nm[c]}")
    if len(changed) > 50:
        print("경고: 변경 50건 초과 — 구조 이상 의심. 갱신 중단"); sys.exit(1)

    shutil.copy2(OUT, OUT.replace(".csv", "_prev.csv"))
    new.to_csv(OUT, index=False, encoding="utf-8-sig")
    print(f"갱신 완료: {len(new)}행 (이전본은 bos3218_prev.csv 백업)")


if __name__ == "__main__":
    main()
