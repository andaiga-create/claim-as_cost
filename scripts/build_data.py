"""data/ 의 xls(누적+월별)를 읽어 대시보드용 집계 JSON(docs/data.json)을 생성한다.

엑셀 보고서(raw 시트)의 계산식을 그대로 재현:
  고객/호선/장비 = 프로젝트코드 분해, 상선/특수선 = 코드 끝 2자리,
  불량코드/불량명칭/유무상 = A/S번호로 list 조회, claim/as = 불량코드 첫 글자(a/c)
"""
import glob, json, os, re, sys
from datetime import datetime
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DIR = os.path.join(ROOT, "data", "raw_data")
LIST_DIR = os.path.join(ROOT, "data", "as_list")
OUT = os.path.join(ROOT, "docs", "data.json")


def read_xls(path):
    return pd.read_excel(path, dtype=str)


def month_of(path):
    m = re.search(r"_(\d{6})\.xls$", path)
    return m.group(1) if m else None


def load_raw():
    files = sorted(glob.glob(os.path.join(RAW_DIR, "*.xls")), key=lambda p: (month_of(p) is not None, month_of(p) or ""))
    df = pd.DataFrame()
    for f in files:
        d = read_xls(f)
        d = d[d["A/S번호"].notna() & d["입고일자"].notna()].copy()
        d["_dt"] = pd.to_datetime(d["입고일자"], errors="coerce")
        d = d[d["_dt"].notna()]
        ym = month_of(f)
        if ym and len(df):  # 월별 파일은 같은 월의 기존 데이터를 대체 (재업로드 안전)
            df = df[df["_dt"].dt.strftime("%Y%m") != ym]
        df = pd.concat([df, d], ignore_index=True)
    df["공급가"] = pd.to_numeric(df["공급가"].str.replace(",", ""), errors="coerce").fillna(0)
    return df


def load_list():
    files = sorted(glob.glob(os.path.join(LIST_DIR, "*.xls")))
    df = pd.concat([read_xls(f) for f in files], ignore_index=True)  # 엑셀 list 시트와 같은 순서(파일명순)
    df["_dt"] = pd.to_datetime(df["접수일자"], errors="coerce")
    return df


def split_code(code):
    parts = str(code).split("-")
    if len(parts) < 4:
        return (parts[0] if parts else ""), "", ""
    return parts[0], parts[1], "-".join(parts[2:-1])


def main():
    raw, lst = load_raw(), load_list()

    # list 조회: 엑셀 VLOOKUP(범위 조회, 3번째 인수 생략)과 동일하게 근사 일치(이진 탐색)로 재현
    keys = lst["접수번호"].str.lower().tolist()

    def approx_idx(x):
        lo, hi = 0, len(keys) - 1
        x = str(x).lower()
        while lo <= hi:
            mid = (lo + hi) // 2
            if keys[mid] == x:  # 엑셀은 탐색 중 정확 일치를 만나면 그 행을 반환 (엑셀 결과와 100% 일치 확인)
                return mid
            if keys[mid] < x:
                lo = mid + 1
            else:
                hi = mid - 1
        return hi  # -1 이면 #N/A

    idx = raw["A/S번호"].map(approx_idx)
    for col, src in [("불량코드", "불량코드"), ("불량명칭", "불량명칭"), ("유무상", "유상구분")]:
        vals = lst[src].tolist()
        raw[col] = idx.map(lambda i: vals[i] if i >= 0 else None)
    lst = lst.drop_duplicates("접수번호", keep="last")  # 이후 접수 건수 집계용
    raw["유무상"] = raw["유무상"].str.replace(r"^\d\s*", "", regex=True)

    sp = raw["프로젝트코드"].fillna("").map(split_code)
    raw["고객"] = sp.map(lambda x: x[0])
    raw["호선"] = sp.map(lambda x: x[1])
    raw["장비"] = sp.map(lambda x: x[2])
    tail = raw["프로젝트코드"].fillna("").str[-2:]
    raw["선종"] = tail.map(lambda t: "상선" if t in ("C1", "C2") else "특수선" if t == "S1" else "기타")
    first = raw["불량코드"].fillna("").str[:1].str.lower()
    raw["구분"] = first.map(lambda c: "A/S" if c == "a" else "CLAIM" if c == "c" else "미분류")
    raw["불량명칭"] = raw["불량명칭"].fillna("미분류")
    raw["유무상"] = raw["유무상"].fillna("미분류")
    raw["연"] = raw["_dt"].dt.year
    raw["월"] = raw["_dt"].dt.month

    dims = ["연", "월", "구분", "유무상", "고객", "선종", "호선", "장비", "거래처명", "불량명칭"]
    g = raw.groupby(dims, dropna=False).agg(공급가=("공급가", "sum"), 건수=("공급가", "size")).reset_index()
    g["거래처명"] = g["거래처명"].fillna("")

    # 접수 건수 (list)
    lst["연"] = lst["_dt"].dt.year
    lst["월"] = lst["_dt"].dt.month
    lst["AS구분"] = lst["AS구분"].fillna("").str.replace(r"^\d\s*", "", regex=True)
    lst["유무상"] = lst["유상구분"].fillna("").str.replace(r"^\d\s*", "", regex=True)
    lst = lst[lst["_dt"].notna() & (lst["연"] >= 2020)]
    lg = lst.groupby(["연", "월", "AS구분", "유무상"]).size().reset_index(name="건수")

    out = {
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "latest": raw["_dt"].max().strftime("%Y-%m-%d"),
        "dims": dims,
        "rows": g[dims + ["공급가", "건수"]].values.tolist(),
        "receipts": lg.values.tolist(),
        "unmatched": {"rows": int((raw["구분"] == "미분류").sum()), "amount": float(raw.loc[raw["구분"] == "미분류", "공급가"].sum())},
        "receipt_cols": ["연", "월", "AS구분", "유무상", "건수"],
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"), default=lambda o: o.item() if hasattr(o, "item") else str(o))
    print(f"raw {len(raw)}행, list {len(lst)}행 -> {len(g)}개 집계행, {OUT}")
    print(raw.pivot_table(index="연", columns="구분", values="공급가", aggfunc="sum", margins=True).round(0).to_string())
    print(raw[raw["유무상"] == "무상"].pivot_table(index="연", columns="구분", values="공급가", aggfunc="sum").round(0).to_string())


if __name__ == "__main__":
    sys.exit(main())
