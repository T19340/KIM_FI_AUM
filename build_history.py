# -*- coding: utf-8 -*-
"""xlsb에서 히스토리 데이터 1회 추출 → data/ 스토어 구축.

- MOS2101  → data/mos2101_history.parquet (표_MOS2101 전체, 저장값 기준)
- BOS3218  → data/bos3218.csv (B·F열 fill-down 적용)
- 수익자    → data/수익자.csv
"""
import os, re, sys
import pandas as pd
from pyxlsb import open_workbook, convert_date

SRC = r"D:\복호화\FI운용본부_수탁고양식_new.xlsb"
BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
os.makedirs(DATA, exist_ok=True)

DATE_COLS_MOS = {"영업일", "처리일", "설정일", "최초설정일", "해지일", "직전결산일", "다음결산일", "다음보수일"}


def read_sheet(wb, name, max_col=None):
    """실제 행 번호(cell.r) 기준 dense grid 반환 (빈 행 보존)."""
    cells = {}
    max_r = max_c = -1
    with wb.get_sheet(name) as sh:
        for row in sh.rows():
            for c in row:
                if c.v is None:
                    continue
                if max_col and c.c >= max_col:
                    continue
                cells[(c.r, c.c)] = c.v
                if c.r > max_r:
                    max_r = c.r
                if c.c > max_c:
                    max_c = c.c
    return [[cells.get((r, c)) for c in range(max_c + 1)]
            for r in range(max_r + 1)]


def uniquify(names):
    seen, out = {}, []
    for i, n in enumerate(names):
        n = str(n).strip() if n is not None else ""
        if not n:
            n = f"col{i+1}"
        if n in seen:
            seen[n] += 1
            n = f"{n}__{seen[n]}"
        else:
            seen[n] = 0
        out.append(n)
    return out


def main():
    wb = open_workbook(SRC)
    print("sheets:", wb.sheets, flush=True)

    # ---- MOS2101 (272열 = 표_MOS2101 범위) ----
    print("reading MOS2101 ...", flush=True)
    t = read_sheet(wb, "MOS2101", max_col=272)
    header = uniquify(t[0])
    df = pd.DataFrame(t[1:], columns=header)
    for c in df.columns:
        base = re.sub(r"__\d+$", "", c)
        if base in DATE_COLS_MOS:
            df[c] = df[c].map(lambda v: convert_date(v)
                              if isinstance(v, (int, float)) and pd.notna(v) else None)
    print("MOS2101 shape:", df.shape, flush=True)
    print("처리일 range:", df["처리일"].min(), "~", df["처리일"].max(), flush=True)
    # 혼합 타입 열 정리: 전부 숫자로 읽히면 숫자, 아니면 문자열
    for c in df.columns:
        if df[c].dtype == object:
            s = df[c]
            num = pd.to_numeric(s, errors="coerce")
            if num.notna().sum() == s.notna().sum():
                df[c] = num
            else:
                df[c] = s.map(lambda v: v if v is None else str(v))
    df.to_parquet(os.path.join(DATA, "mos2101_history.parquet"))

    # ---- BOS3218 ----
    print("reading BOS3218 ...", flush=True)
    t = read_sheet(wb, "BOS3218")
    bos_cols = ["순번", "모펀드", "모펀드명", "설정해지구분", "투자비율",
                "종류형모펀드", "종류형모펀드명", "종류형자펀드", "종류형자펀드명",
                "최종모펀드", "종류형모펀드_k", "종류형자펀드_l",
                "최종모부재", "종류형모부재", "종류형자부재", "실질최종모펀드", "최종모펀드명"]
    data_rows = t[1:]  # 헤더 1행 스킵 (헤더 셀 내 줄바꿈이 3행처럼 보이지만 실제 1행)
    width = max(len(r) for r in data_rows)
    data_rows = [r + [None] * (width - len(r)) for r in data_rows]
    bos = pd.DataFrame(data_rows, columns=bos_cols[:width] if width <= 17 else uniquify([None]*width))
    # 매크로 BOS3218Update 와 동일: B(모펀드)·F(종류형모펀드) fill-down
    bos["모펀드"] = bos["모펀드"].ffill()
    bos["종류형모펀드"] = bos["종류형모펀드"].ffill()
    bos = bos[bos.drop(columns=["순번"]).notna().any(axis=1)]
    print("BOS3218 shape:", bos.shape, flush=True)
    bos.to_csv(os.path.join(DATA, "bos3218.csv"), index=False, encoding="utf-8-sig")

    # ---- 수익자 ----
    print("reading 수익자 ...", flush=True)
    t = read_sheet(wb, "수익자")
    ben = pd.DataFrame(t[1:], columns=uniquify(t[0]))
    ben = ben[ben["열1"].notna()]
    print("수익자 shape:", ben.shape, flush=True)
    ben.to_csv(os.path.join(DATA, "수익자.csv"), index=False, encoding="utf-8-sig")

    print("DONE", flush=True)


if __name__ == "__main__":
    main()
