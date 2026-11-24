# -*- coding: utf-8 -*-
"""엑셀 보고서 생성 (update.bat 실행 시 build_data.py 가 호출)

기존 보고서의 raw / list 구조를 그대로 두되, 접수 목록 조회(VLOOKUP)를 '정확 일치'로 고치고
웹 대시보드와 같은 기본 조건(무상 · A/S+CLAIM · 상선+특수선)으로 집계한다.

시트
  dashboard : KPI + 차트 (집계 시트 참조)
  집계      : 연도별 / 월별 / 순위 표. 모든 금액은 raw 시트를 SUMIFS 로 다시 계산하는 수식
  raw       : 매입 내역 (불량코드·불량명칭·유/무상은 list 조회 수식, 나머지 파생 열도 수식)
  list      : 접수 목록 (조회 대상)
  매출액    : docs/settings.json 의 연도별 매출
  미분류    : list 에서 못 찾았거나 불량코드가 없는 건 (ERP 확인용)
  설명      : 집계 규칙
수식에는 계산된 값을 함께 저장하므로 보호된 보기/미리보기에서도 숫자가 보인다.
엑셀 파일은 GitHub 에 올라가지 않는다 (.gitignore: 엑셀보고서/).
"""
import os
from datetime import datetime

import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as CN

SHEET = "'집계'!"
WON = "#,##0"
FONT = "맑은 고딕"
INDIGO, TEAL, GRAY = "#4F46E5", "#0D9488", "#7A8599"


def crit(v):
    """SUMIFS 조건 문자열 (와일드카드·연산자 문자 이스케이프). 빈 값은 '' 로 빈 셀과 일치"""
    s = str(v)
    if s == "":
        return '""'
    s = s.replace("~", "~~").replace("*", "~*").replace("?", "~?")
    if s[0] in "=<>":
        s = "=" + s
    return '"' + s.replace('"', '""') + '"'


def write(root, raw, lk, lst_all, mode, settings, latest):
    """raw : 파생 열이 계산된 매입 DataFrame (_lidx = lk 의 행 위치, -1 이면 못 찾음)
    lk  : 조회 대상 목록 DataFrame (exact: 중복 제거, excel: 원본 순서 그대로)
    lst_all : 접수 건수 집계용 목록 (중복 제거)"""
    ly, lm = int(latest[:4]), int(latest[5:7])
    rev = {int(k): v for k, v in (settings.get("revenue") or {}).items()}
    ratio_basis = settings.get("ratio_basis", "avg3")
    tgt = (settings.get("kpi") or {}).get(str(ly)) or {}

    out_dir = os.path.join(root, "엑셀보고서")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"CLAIM_AS_보고서_{ly}{lm:02d}.xlsx")
    wb = xlsxwriter.Workbook(path, {"default_format_properties": {"font_name": FONT, "font_size": 10}})
    wb.set_calc_mode("auto")
    f = lambda **k: wb.add_format({"font_name": FONT, "font_size": 10, **k})
    F_HEAD = f(bold=True, font_color="#FFFFFF", bg_color=INDIGO, align="center", valign="vcenter", border=0)
    F_TITLE = f(bold=True, font_size=12)
    F_NOTE = f(font_size=9, font_color=GRAY)
    F_WON, F_PCT, F_INT = f(num_format=WON), f(num_format="0.00%"), f(num_format="0")
    F_BWON = f(num_format=WON, bold=True)
    F_DATE = f(num_format="yyyy-mm-dd")
    F_TXT = f()
    F_BOLD = f(bold=True)

    ws_d = wb.add_worksheet("dashboard")
    ws_a = wb.add_worksheet("집계")
    ws_r = wb.add_worksheet("raw")
    ws_l = wb.add_worksheet("list")
    ws_v = wb.add_worksheet("매출액")
    ws_u = wb.add_worksheet("미분류")
    ws_n = wb.add_worksheet("설명")

    # ------------------------------------------------------------ list
    lcols = ["접수번호", "접수일자", "프로젝트코드", "프로젝트명", "고객명", "불량코드", "불량명칭", "AS구분", "유상구분", "연", "월", "집계포함"]
    lk = lk.reset_index(drop=True)
    last_keep = ~lk["접수번호"].duplicated(keep="last")  # 중복 접수번호는 마지막 행만 건수에 포함 (웹과 동일)
    m = len(lk)
    ws_l.add_table(0, 0, m, len(lcols) - 1, {"name": "as_list", "style": "Table Style Light 9",
                                             "columns": [{"header": h} for h in lcols]})
    ldt = pd.to_datetime(lk["접수일자"], errors="coerce")
    for i in range(m):
        row = lk.iloc[i]
        for c, k in enumerate(["접수번호", "접수일자", "프로젝트코드", "프로젝트명", "고객명", "불량코드", "불량명칭", "AS구분", "유상구분"]):
            v = row[k]
            if pd.notna(v) and v != "":
                ws_l.write_string(i + 1, c, str(v))
        if pd.notna(ldt.iloc[i]):
            ws_l.write_number(i + 1, 9, int(ldt.iloc[i].year))
            ws_l.write_number(i + 1, 10, int(ldt.iloc[i].month))
        ws_l.write_number(i + 1, 11, 1 if last_keep.iloc[i] else 0)
    for c, w in enumerate([14, 11, 20, 30, 18, 9, 18, 14, 9, 6, 5, 8]):
        ws_l.set_column(c, c, w)
    ws_l.freeze_panes(1, 0)
    LR = m + 1  # list 마지막 행(1부터)

    # ------------------------------------------------------------ raw
    rcols = ["입고일자", "거래처명", "품목번호", "품명", "규격", "재질", "단위", "수량", "단가", "공급가", "프로젝트코드",
             "A/S번호", "비고", "불량코드", "불량명칭", "고객", "호선", "장비", "상선/특수선", "claim/as", "유/무상",
             "연", "월", "list행", "집계대상"]
    raw = raw.sort_values(["_dt", "A/S번호"], kind="stable").reset_index(drop=True)
    n = len(raw)
    E = n + 1
    ws_r.add_table(0, 0, n, len(rcols) - 1, {"name": "raw_data", "style": "Table Style Medium 2",
                                             "columns": [{"header": h} for h in rcols]})
    LC = lambda col: f"list!${col}$2:${col}${LR}"
    mt = 1 if mode == "excel" else 0
    mask_all = ((raw["유무상"] == "무상") & raw["구분"].isin(["A/S", "CLAIM"]) & raw["선종"].isin(["상선", "특수선"]))
    for i in range(n):
        r, x = i + 2, raw.iloc[i]  # r = 엑셀 행 번호
        ws_r.write_datetime(i + 1, 0, x["_dt"].to_pydatetime(), F_DATE)
        for c, k in [(1, "거래처명"), (2, "품목번호"), (3, "품명"), (4, "규격"), (5, "재질"), (6, "단위"), (7, "수량"),
                     (8, "단가"), (10, "프로젝트코드"), (11, "A/S번호"), (12, "비고")]:
            v = x[k]
            if pd.notna(v) and v != "":
                ws_r.write_string(i + 1, c, str(v))
        ws_r.write_number(i + 1, 9, float(x["공급가"]), F_WON)
        li = int(x["_lidx"]) + 1  # 0 = 못 찾음
        K = f"K{r}"
        dash = f'(LEN({K})-LEN(SUBSTITUTE({K},"-","")))'
        W = ws_r.write_formula
        W(i + 1, 23, f'=IFERROR(MATCH(L{r},{LC("A")},{mt}),0)', F_INT, li)

        def look(col):  # list 에서 값을 가져오되 못 찾았거나 비어 있으면 "미분류"
            return f'=IF(X{r}=0,"미분류",IF(INDEX({LC(col)},X{r})="","미분류",INDEX({LC(col)},X{r})))'
        W(i + 1, 13, look("F"), F_TXT, x["불량코드"] if pd.notna(x["불량코드"]) and x["불량코드"] != "" else "미분류")
        W(i + 1, 14, look("G"), F_TXT, x["불량명칭"])
        v = f"INDEX({LC('I')},X{r})"
        W(i + 1, 20, f'=IF(X{r}=0,"미분류",IF({v}="","미분류",IFERROR(TRIM(MID({v},FIND(" ",{v})+1,50)),{v})))', F_TXT, x["유무상"])
        W(i + 1, 15, f'=IFERROR(LEFT({K},FIND("-",{K})-1),{K})', F_TXT, x["고객"])
        W(i + 1, 16, f'=IF({dash}<3,"",MID({K},FIND("-",{K})+1,FIND("-",{K},FIND("-",{K})+1)-FIND("-",{K})-1))', F_TXT, x["호선"])
        W(i + 1, 17, f'=IF({dash}<3,"",MID({K},LEN(P{r})+LEN(Q{r})+3,FIND("|",SUBSTITUTE({K},"-","|",{dash}))-(LEN(P{r})+LEN(Q{r})+3)))', F_TXT, x["장비"])
        W(i + 1, 18, f'=IF(OR(EXACT(RIGHT({K},2),"C1"),EXACT(RIGHT({K},2),"C2")),"상선",IF(EXACT(RIGHT({K},2),"S1"),"특수선","기타"))', F_TXT, x["선종"])
        W(i + 1, 19, f'=IF(LEFT(N{r},1)="a","A/S",IF(LEFT(N{r},1)="c","CLAIM","미분류"))', F_TXT, x["구분"])
        W(i + 1, 21, f"=YEAR(A{r})", F_INT, int(x["연"]))
        W(i + 1, 22, f"=MONTH(A{r})", F_INT, int(x["월"]))
        W(i + 1, 24, f'=IF(AND(U{r}="무상",OR(T{r}="A/S",T{r}="CLAIM"),OR(S{r}="상선",S{r}="특수선")),1,0)', F_INT, int(mask_all.iloc[i]))
    for c, w in enumerate([11, 22, 12, 22, 12, 8, 6, 7, 11, 12, 22, 14, 24, 9, 18, 8, 10, 12, 10, 9, 9, 6, 5, 7, 8]):
        ws_r.set_column(c, c, w)
    ws_r.freeze_panes(1, 0)

    # ------------------------------------------------------------ 매출액
    ys = sorted(raw["연"].unique())
    y0, y1 = int(min(min(ys), min(rev) if rev else 9999)), int(max(max(ys), max(rev) if rev else 0))
    ws_v.write_row(0, 0, ["연도", "매출액(원)"], F_HEAD)
    vrow = {}
    for y in range(y0, y1 + 1):
        vrow[y] = y - y0 + 2
        ws_v.write_number(y - y0 + 1, 0, y)
        if rev.get(y) is not None:
            ws_v.write_number(y - y0 + 1, 1, rev[y], F_WON)
        else:
            ws_v.write_blank(y - y0 + 1, 1, None, F_WON)
    ws_v.write(0, 3, "※ 값은 docs/settings.json 의 revenue 에서 가져옵니다. update.bat 실행 시 갱신됩니다.", F_NOTE)
    ws_v.set_column(0, 0, 8)
    ws_v.set_column(1, 1, 18)

    # ------------------------------------------------------------ 집계
    A = ws_a
    A.set_column(0, 0, 22)
    A.set_column(1, 8, 15)
    A.write(0, 0, "A/S·CLAIM 비용 집계 — 모든 금액은 raw 시트를 SUMIFS 로 다시 계산한 수식입니다", f(bold=True, font_size=13))
    A.write(1, 0, "조건: 유/무상 = 무상 · 구분 = A/S + CLAIM · 선종 = 상선 + 특수선 (웹 대시보드 기본값, raw 시트의 '집계대상'=1)", F_NOTE)
    rg = lambda col: f"raw!${col}$2:${col}${E}"
    AMT, YR, MO, TG, GB = rg("J"), rg("V"), rg("W"), rg("Y"), rg("T")
    DIM = {"고객": rg("P"), "호선": rg("Q"), "장비": rg("R"), "거래처명": rg("B"), "불량명칭": rg("O"), "선종": rg("S")}
    sel = raw[mask_all]
    fl = lambda g, y, extra="": f"=SUMIFS({AMT},{TG},1,{YR},{y},{GB},\"{g}\"{extra})"
    pos = {}

    # 1) 연도별
    r0 = 4
    A.write(r0 - 1, 0, "연도별", F_TITLE)
    A.write_row(r0, 0, ["연도", "A/S", "CLAIM", "합계", "기준 매출", "매출 대비"], F_HEAD)
    for i, y in enumerate(ys):
        y, rr = int(y), r0 + 2 + i  # rr: 엑셀 행 번호
        sy = sel[sel["연"] == y]
        a, c = sy[sy["구분"] == "A/S"]["공급가"].sum(), sy[sy["구분"] == "CLAIM"]["공급가"].sum()
        A.write_number(rr - 1, 0, y)
        A.write_formula(rr - 1, 1, fl("A/S", f"$A{rr}"), F_WON, a)
        A.write_formula(rr - 1, 2, fl("CLAIM", f"$A{rr}"), F_WON, c)
        A.write_formula(rr - 1, 3, f"=B{rr}+C{rr}", F_BWON, a + c)
        if ratio_basis == "year":
            vs = [rev.get(y)]
            fb = f'=IF(매출액!$B${vrow[y]}="","",매출액!$B${vrow[y]})'
        else:
            vs = [rev.get(k) for k in (y - 2, y - 1, y)]
            lo = vrow[max(y - 2, y0)]
            fb = f'=IFERROR(AVERAGE(매출액!$B${lo}:$B${vrow[y]}),"")'
        vs = [v for v in vs if v is not None]
        base = sum(vs) / len(vs) if vs else ""
        A.write_formula(rr - 1, 4, fb, F_WON, base)
        A.write_formula(rr - 1, 5, f'=IF(E{rr}="","",D{rr}/E{rr})', F_PCT, (a + c) / base if base != "" else "")
        pos[y] = rr
    ya, yb = r0 + 2, r0 + 1 + len(ys)
    pos["year"] = (ya, yb)

    # 2) 최신 연도 월별
    r0 = yb + 3
    A.write(r0 - 1, 0, f"{ly}년 월별", F_TITLE)
    A.write_row(r0, 0, ["월", "A/S", "CLAIM", "합계", "누계", f"{ly - 1}년 합계", f"{ly - 1}년 누계"], F_HEAD)
    cum = pcum = 0
    for i in range(12):
        mm, rr = i + 1, r0 + 2 + i

        def part(y, g):
            s = sel[(sel["연"] == y) & (sel["월"] == mm) & (sel["구분"] == g)]
            return s["공급가"].sum()
        a, c, pa, pc = part(ly, "A/S"), part(ly, "CLAIM"), part(ly - 1, "A/S"), part(ly - 1, "CLAIM")
        A.write_number(rr - 1, 0, mm, f(num_format='0"월"', align="left"))
        A.write_formula(rr - 1, 1, fl("A/S", ly, f",{MO},$A{rr}"), F_WON, a)
        A.write_formula(rr - 1, 2, fl("CLAIM", ly, f",{MO},$A{rr}"), F_WON, c)
        A.write_formula(rr - 1, 3, f"=B{rr}+C{rr}", F_BWON, a + c)
        cum += a + c
        pcum += pa + pc
        if mm <= lm:
            A.write_formula(rr - 1, 4, f"=SUM($D${r0 + 2}:D{rr})", F_WON, cum)
        pt = f"=SUMIFS({AMT},{TG},1,{YR},{ly - 1},{MO},$A{rr})"
        A.write_formula(rr - 1, 5, pt, F_WON, pa + pc)
        A.write_formula(rr - 1, 6, f"=SUM($F${r0 + 2}:F{rr})", F_WON, pcum)
    pos["month"] = (r0 + 2, r0 + 13)

    # 3) KPI 계산
    r0 = r0 + 16
    A.write(r0 - 1, 0, f"KPI ({ly}년 1~{lm}월)", F_TITLE)
    cur = sel[sel["연"] == ly]
    prv = sel[(sel["연"] == ly - 1) & (sel["월"] <= lm)]
    cur_v, prv_v = cur["공급가"].sum(), prv["공급가"].sum()
    claim_v = cur[cur["구분"] == "CLAIM"]["공급가"].sum()
    la = lst_all[(lst_all["연"] == ly) & (lst_all["월"] <= lm) & (lst_all["유무상"] == "무상")]
    rc_all, rc_as, rc_cl = len(la), int(la["AS구분"].str.contains("A/S", regex=False).sum()), int(la["AS구분"].str.contains("CLAIM", regex=False).sum())
    rows_kpi = [
        ("총 비용", f"=D{pos[ly]}", cur_v, F_WON),
        ("전년 동기", f"=SUMIFS({AMT},{TG},1,{YR},{ly - 1},{MO},\"<=\"&{lm})", prv_v, F_WON),
        ("전년 동기 대비", f'=IF(B{r0 + 2}=0,"-",B{r0 + 1}/B{r0 + 2}-1)', (cur_v / prv_v - 1) if prv_v else "-", f(num_format="+0.0%;-0.0%")),
        ("CLAIM 금액", f"=C{pos[ly]}", claim_v, F_WON),
        ("CLAIM 비중", f'=IF(B{r0 + 1}=0,"-",B{r0 + 4}/B{r0 + 1})', claim_v / cur_v if cur_v else "-", f(num_format="0.0%")),
        ("접수 건수(무상)", f'=COUNTIFS(list!$J$2:$J${LR},{ly},list!$K$2:$K${LR},"<="&{lm},list!$I$2:$I${LR},"*무상",list!$L$2:$L${LR},1)', rc_all, F_INT),
        ("  A/S 접수", f'=COUNTIFS(list!$J$2:$J${LR},{ly},list!$K$2:$K${LR},"<="&{lm},list!$I$2:$I${LR},"*무상",list!$H$2:$H${LR},"*A/S*",list!$L$2:$L${LR},1)', rc_as, F_INT),
        ("  CLAIM 접수", f'=COUNTIFS(list!$J$2:$J${LR},{ly},list!$K$2:$K${LR},"<="&{lm},list!$I$2:$I${LR},"*무상",list!$H$2:$H${LR},"*CLAIM*",list!$L$2:$L${LR},1)', rc_cl, F_INT),
        ("미분류 건수(전체 조건)", f'=COUNTIF({GB},"미분류")', int((raw["구분"] == "미분류").sum()), F_INT),
        ("미분류 금액(전체 조건)", f'=SUMIFS({AMT},{GB},"미분류")', float(raw.loc[raw["구분"] == "미분류", "공급가"].sum()), F_WON),
    ]
    for i, (lab, fm, val, fmt) in enumerate(rows_kpi):
        A.write(r0 + i, 0, lab)
        A.write_formula(r0 + i, 1, fm, fmt, val)
    pos["kpi"] = r0 + 1
    kcell = {k: f"'집계'!$B${r0 + 1 + i}" for i, k in enumerate(["cur", "prv", "delta", "claim", "claim_r", "rc", "rc_as", "rc_cl", "un_n", "un_v"])}

    # 4) 순위 표 (최신 연도, 합계 기준 내림차순)
    r0 = r0 + len(rows_kpi) + 3

    def top(r0, label, key, dimcol, limit):
        s = sel[sel["연"] == ly].copy()
        s["_k"] = s[key].fillna("").astype(str)
        tot = s.groupby("_k")["공급가"].sum().sort_values(ascending=False)
        tot = tot[tot != 0].head(limit)
        A.write(r0 - 1, 0, f"{ly}년 {label}", F_TITLE)
        A.write_row(r0, 0, [label.split()[0], "A/S", "CLAIM", "합계"], F_HEAD)
        for i, k in enumerate(tot.index):
            rr = r0 + 2 + i
            A.write_string(rr - 1, 0, k if k else "(없음)")
            for c, g in ((1, "A/S"), (2, "CLAIM")):
                val = s[(s["_k"] == k) & (s["구분"] == g)]["공급가"].sum()
                A.write_formula(rr - 1, c, fl(g, ly, f",{dimcol},{crit(k)}"), F_WON, val)
            A.write_formula(rr - 1, 3, f"=B{rr}+C{rr}", F_BWON, tot[k])
        return r0 + 2, r0 + 1 + len(tot)

    pos["ship"] = top(r0, "호선 TOP 10", "호선", DIM["호선"], 10)
    pos["equip"] = top(pos["ship"][1] + 3, "장비 TOP 10", "장비", DIM["장비"], 10)
    pos["vendor"] = top(pos["equip"][1] + 3, "거래처 TOP 10", "거래처명", DIM["거래처명"], 10)
    pos["defect"] = top(pos["vendor"][1] + 3, "불량명칭 TOP 9", "불량명칭", DIM["불량명칭"], 9)
    pos["cust"] = top(pos["defect"][1] + 3, "고객 전체", "고객", DIM["고객"], 30)
    pos["seg"] = top(pos["cust"][1] + 3, "선종 전체", "선종", DIM["선종"], 5)
    A.freeze_panes(3, 0)

    # ------------------------------------------------------------ dashboard
    D = ws_d
    D.hide_gridlines(2)
    D.set_column(0, 17, 11)
    D.write(0, 0, f"A/S · CLAIM 비용 현황 ({ly}년 1~{lm}월 기준)", f(bold=True, font_size=18))
    D.write(1, 0, "무상 · A/S+CLAIM · 상선+특수선 · 금액은 집계 시트(수식) 참조 · 웹 대시보드와 같은 규칙", F_NOTE)
    card = lambda **k: f(bg_color="#EEF1F6", **k)
    delta_s = f"+{(cur_v / prv_v - 1) * 100:.1f}%" if prv_v and cur_v >= prv_v else (f"{(cur_v / prv_v - 1) * 100:.1f}%" if prv_v else "-")
    cards = [
        ("총 비용", f"={kcell['cur']}", cur_v, card(bold=True, font_size=16, num_format=WON, align="left"),
         '="무상 · A/S+CLAIM · 상선+특수선"', "무상 · A/S+CLAIM · 상선+특수선"),
        (f"전년({ly - 1}) 동기 대비", f"={kcell['delta']}", (cur_v / prv_v - 1) if prv_v else "-",
         card(bold=True, font_size=16, num_format="+0.0%;-0.0%", align="left"),
         f'="전년 "&TEXT({kcell["prv"]},"#,##0")&"원"', f"전년 {prv_v:,.0f}원"),
        ("매출 대비 비용", f"='집계'!$F${pos[ly]}", None, card(bold=True, font_size=16, num_format="0.00%", align="left"),
         None, None),
        ("CLAIM 비중", f"={kcell['claim_r']}", claim_v / cur_v if cur_v else "-", card(bold=True, font_size=16, num_format="0.0%", align="left"),
         f'="CLAIM "&TEXT({kcell["claim"]},"#,##0")&"원"', f"CLAIM {claim_v:,.0f}원"),
        ("접수 건수(무상)", f"={kcell['rc']}", rc_all, card(bold=True, font_size=16, num_format='#,##0"건"', align="left"),
         f'="A/S "&{kcell["rc_as"]}&" · CLAIM "&{kcell["rc_cl"]}', f"A/S {rc_as} · CLAIM {rc_cl}"),
    ]
    vs_ly = [rev.get(k) for k in ((ly,) if ratio_basis == "year" else (ly - 2, ly - 1, ly))]
    vs_ly = [v for v in vs_ly if v is not None]
    base_ly = sum(vs_ly) / len(vs_ly) if vs_ly else ""
    ratio_val = cur_v / base_ly if base_ly != "" else "-"
    for i, (lab, fm, val, fmt, sf, sv) in enumerate(cards):
        c = i * 3
        D.merge_range(3, c, 3, c + 2, lab, card(font_size=10, font_color=GRAY))
        if val is None:
            val = ratio_val
        D.merge_range(4, c, 4, c + 2, "", fmt)
        D.write_formula(4, c, fm, fmt, val)
        if sf is None:  # 매출 대비: 목표 표시
            tr = tgt.get("ratio_target")
            txt = f"목표 {tr * 100:.2f}%" if tr else (f"기준 매출 {base_ly:,.0f}원" if base_ly != "" else "매출 미입력")
            sf, sv = (f'="목표 {tr * 100:.2f}%"' if tr else f'=IF(\'집계\'!$E${pos[ly]}="","매출 미입력","기준 매출 "&TEXT(\'집계\'!$E${pos[ly]},"#,##0")&"원")'), txt
        D.merge_range(5, c, 5, c + 2, "", card(font_size=9, font_color=GRAY, align="left"))
        D.write_formula(5, c, sf, card(font_size=9, font_color=GRAY, align="left"), sv)
    D.set_row(4, 26)

    def rng(sheet, c, a, b):
        return [sheet, a - 1, c, b - 1, c]
    ya, yb = pos["year"]
    ma, mb = pos["month"]
    PALETTE = {"A/S": INDIGO, "CLAIM": "#F59E0B"}

    def col_chart(title, cats, series, ctype="column", sub="stacked", w=560, h=300):
        ch = wb.add_chart({"type": ctype, "subtype": sub} if ctype != "pie" else {"type": "pie"})
        for name, col, a, b in series:
            ch.add_series({"name": name, "categories": ["집계", cats[0] - 1, cats[1], cats[2] - 1, cats[1]],
                           "values": ["집계", a - 1, col, b - 1, col],
                           "fill": {"color": PALETTE.get(name, INDIGO)}, "gap": 60})
        ch.set_title({"name": title, "name_font": {"size": 11, "bold": True}})
        ch.set_size({"width": w, "height": h})
        ch.set_legend({"position": "bottom"})
        ch.set_y_axis({"num_format": '#,##0,,"백만"', "major_gridlines": {"visible": True, "line": {"color": "#E5E8EF"}}})
        return ch

    ch = col_chart("연도별 비용", (ya, 0, yb), [("A/S", 1, ya, yb), ("CLAIM", 2, ya, yb)])
    D.insert_chart("A8", ch)
    ch = col_chart(f"{ly}년 월별 비용", (ma, 0, mb), [("A/S", 1, ma, mb), ("CLAIM", 2, ma, mb)])
    D.insert_chart("G8", ch)
    ln = wb.add_chart({"type": "line"})
    ln.add_series({"name": f"{ly}년 누계", "categories": ["집계", ma - 1, 0, mb - 1, 0], "values": ["집계", ma - 1, 4, mb - 1, 4],
                   "line": {"color": INDIGO, "width": 2.25}})
    ln.add_series({"name": f"{ly - 1}년 누계", "categories": ["집계", ma - 1, 0, mb - 1, 0], "values": ["집계", ma - 1, 6, mb - 1, 6],
                   "line": {"color": "#9AA3B5", "width": 1.75, "dash_type": "dash"}})
    ln.set_title({"name": "누적 추이", "name_font": {"size": 11, "bold": True}})
    ln.set_size({"width": 560, "height": 300})
    ln.set_legend({"position": "bottom"})
    ln.set_y_axis({"num_format": '#,##0,,"백만"', "major_gridlines": {"visible": True, "line": {"color": "#E5E8EF"}}})
    D.insert_chart("M8", ln)
    for key, title, anchor in [("ship", "호선 TOP 10", "A24"), ("equip", "장비 TOP 10", "G24"), ("vendor", "거래처 TOP 10", "M24"),
                               ("defect", "불량명칭 TOP 9", "A41"), ("cust", "고객별", "G41")]:
        a, b = pos[key]
        if b >= a:
            b = min(b, a + 14)
            ch = col_chart(f"{ly}년 {title}", (a, 0, b), [("A/S", 1, a, b), ("CLAIM", 2, a, b)], ctype="bar", h=380)
            ch.set_x_axis({"reverse": True, "num_font": {"size": 9}})
            ch.set_y_axis({"num_format": '#,##0,,"백만"', "major_gridlines": {"visible": True, "line": {"color": "#E5E8EF"}}})
            D.insert_chart(anchor, ch)
    a, b = pos["seg"]
    pie = wb.add_chart({"type": "doughnut"})
    pie.add_series({"name": "선종", "categories": ["집계", a - 1, 0, b - 1, 0], "values": ["집계", a - 1, 3, b - 1, 3],
                    "data_labels": {"percentage": True}})
    pie.set_title({"name": f"{ly}년 선종별", "name_font": {"size": 11, "bold": True}})
    pie.set_size({"width": 560, "height": 380})
    pie.set_legend({"position": "bottom"})
    D.insert_chart("M41", pie)

    # ------------------------------------------------------------ 미분류
    un = raw[raw["구분"] == "미분류"].copy()
    ws_u.write_row(0, 0, ["입고일자", "A/S번호", "거래처명", "공급가", "프로젝트코드", "사유"], F_HEAD)
    for i, (_, x) in enumerate(un.iterrows()):
        ws_u.write_datetime(i + 1, 0, x["_dt"].to_pydatetime(), F_DATE)
        ws_u.write_string(i + 1, 1, str(x["A/S번호"]))
        ws_u.write_string(i + 1, 2, str(x["거래처명"]) if pd.notna(x["거래처명"]) else "")
        ws_u.write_number(i + 1, 3, float(x["공급가"]), F_WON)
        ws_u.write_string(i + 1, 4, str(x["프로젝트코드"]) if pd.notna(x["프로젝트코드"]) else "")
        ws_u.write_string(i + 1, 5, "접수 목록에 없음" if x["_lidx"] < 0 else "불량코드 비어 있음")
    ws_u.write(0, 7, "※ ERP 에서 접수번호를 확인해 as_list 에 추가하면 다음 갱신부터 분류됩니다.", F_NOTE)
    for c, w in enumerate([11, 15, 24, 12, 24, 20]):
        ws_u.set_column(c, c, w)
    ws_u.freeze_panes(1, 0)

    # ------------------------------------------------------------ 설명
    notes = [
        "CLAIM·AS 비용 보고서 (자동 생성)",
        f"생성: {datetime.now():%Y-%m-%d %H:%M}   데이터 기준일: {latest}   조회 방식: {'정확 일치' if mode == 'exact' else '엑셀 근사 일치(재현)'}",
        "",
        "■ 집계 규칙 (웹 대시보드와 동일)",
        "  · 고객/호선/장비 = 프로젝트코드를 '-' 로 분해 (예: HHI-P167-HOC-S1 → HHI / P167 / HOC)",
        "  · 선종 = 프로젝트코드 끝 2자리 (C1·C2 = 상선, S1 = 특수선, 그 외 = 기타)",
        "  · 불량코드 / 불량명칭 / 유·무상 = A/S번호로 list 시트 조회 (raw 의 'list행' 열이 일치하는 행 번호)",
        "  · 구분 = 불량코드 첫 글자 (a = A/S, c = CLAIM, 그 외·못 찾음 = 미분류)",
        "  · 집계 대상 = 무상 · A/S+CLAIM · 상선+특수선 (raw 시트 '집계대상' 열 = 1)",
        "",
        "■ VLOOKUP 오류 수정",
        "  · 기존 보고서의 =VLOOKUP(A/S번호, as_list, n) 은 4번째 인수가 없어 근사 일치로 동작했습니다.",
        "    list 가 정렬돼 있지 않아 엉뚱한 접수 건이 연결되거나, 월 자료를 추가할 때 과거 분류가 바뀌었습니다.",
        "  · 이 보고서는 MATCH(…,0) 정확 일치로 조회하고, 못 찾으면 '미분류' 로 표시합니다.",
        "  · 기존 보고서를 직접 고치려면: =IFERROR(VLOOKUP(raw_data[[#This Row],[A/S번호]],as_list[],5,FALSE),\"미분류\")",
        "    (불량코드 5, 불량명칭 6, 유/무상 8 — 3곳 모두 4번째 인수 FALSE)",
        "",
        "■ 갱신: 프로젝트 폴더의 update.bat 을 실행하면 웹 데이터와 이 보고서가 함께 다시 만들어집니다.",
        "  · 목표/매출 수정: docs/settings.json  ·  파일은 엑셀보고서 폴더에 저장되며 GitHub 에는 올라가지 않습니다.",
    ]
    for i, t in enumerate(notes):
        ws_n.write(i, 0, t, f(bold=True, font_size=13) if i == 0 else F_TXT)
    ws_n.set_column(0, 0, 120)

    ws_d.activate()
    wb.close()
    return path
