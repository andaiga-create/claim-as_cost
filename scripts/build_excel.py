# -*- coding: utf-8 -*-
"""기존 엑셀 보고서를 이용해 새 보고서 파일을 만든다 (update.bat 실행 시 build_data.py 가 호출).

하는 일
 1. 엑셀보고서/raw_data/as_claim_raw_data.xlsx, 엑셀보고서/as_list/as_list.xlsx 생성
    - 보고서의 파워쿼리(raw_data, as_list)가 읽는 형식(Sheet1, 원본 열 그대로)
    - 월별 파일 병합·재업로드 교체·A/S번호 하이픈 보정이 반영된 최종 데이터 (웹과 동일)
 2. 가장 최근 보고서(엑셀보고서/ 또는 프로젝트 폴더의 CLAIM_AS_보고서_*.xlsx)를 복사해 새 파일 생성
    - 피벗테이블·슬라이서·차트·데이터 모델·대시보드 등은 손대지 않는다
    - 파워쿼리 원본 폴더(Z:\\...)만 위 두 폴더로 바꾼다
    - raw 시트의 VLOOKUP(근사 일치)을 정확 일치(FALSE) + 못 찾으면 "미분류" 로 고친다
새 보고서를 열고 [데이터] > [모두 새로 고침] 을 누르면 데이터와 피벗이 갱신된다.
엑셀 파일은 GitHub 에 올라가지 않는다 (.gitignore: 엑셀보고서/).
"""
import base64, glob, io, os, re, shutil, struct, tempfile, zipfile
from datetime import datetime

import pandas as pd
import xlsxwriter

RAW_NUM = {"거래처", "수량", "단가", "공급가", "부가세", "합계금액", "전표번호", "발주구분", "입출고근거", "발주번호", "순번", "수불구분", "정산여부"}
LIST_NUM = {"순번", "고객코드", "불량수량"}


def _num(v):
    try:
        f = float(str(v).replace(",", "").strip())
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def write_source(path, df, cols, numeric, extra_col):
    """파워쿼리가 읽는 원본 형식으로 저장. 쿼리가 Column33/Column24 (이름 없는 마지막 열)를 기대하므로 빈 열을 하나 둔다."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet("Sheet1")
    for c, h in enumerate(cols):
        ws.write_string(0, c, h)
    ws.write_string(1, len(cols), " ")  # 헤더 없는 마지막 열 → 파워쿼리에서 Column33 / Column24 로 읽힘
    vals = df[cols].values.tolist()
    for i, row in enumerate(vals, start=1):
        for c, v in enumerate(row):
            if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                continue
            if cols[c] in numeric:
                n = _num(v)
                if n is not None:
                    ws.write_number(i, c, n)
                    continue
            ws.write_string(i, c, str(v))
    wb.close()
    return len(df)


def _patch_mashup(xml_bytes, raw_dir, list_dir):
    text = xml_bytes.decode("utf-16")
    m = re.search(r"(<DataMashup[^>]*>)([^<]+)(</DataMashup>)", text)
    raw = base64.b64decode(m.group(2))
    o = 0
    ver = raw[o:o + 4]; o += 4
    pl = struct.unpack("<I", raw[o:o + 4])[0]; o += 4
    pkg = raw[o:o + pl]; o += pl
    pm = struct.unpack("<I", raw[o:o + 4])[0]; o += 4
    perm = raw[o:o + pm]; o += pm
    ml = struct.unpack("<I", raw[o:o + 4])[0]; o += 4
    meta = raw[o:o + ml]
    zin = zipfile.ZipFile(io.BytesIO(pkg))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "Formulas/Section1.m":
                s = data.decode("utf-8")
                s = re.sub(r'Folder\.Files\("[^"]*?[\\/]raw_data"\)', lambda _: 'Folder.Files("%s")' % raw_dir, s)
                s = re.sub(r'Folder\.Files\("[^"]*?[\\/]as_list"\)', lambda _: 'Folder.Files("%s")' % list_dir, s)
                data = s.encode("utf-8")
            zout.writestr(info.filename, data)
    pkg2 = buf.getvalue()
    # 권한 바인딩(패키지 해시)은 비워 둔다 → 엑셀이 새로 계산 (개인 정보 수준 확인 창이 한 번 뜰 수 있음)
    new = ver + struct.pack("<I", len(pkg2)) + pkg2 + struct.pack("<I", len(perm)) + perm + struct.pack("<I", len(meta)) + meta + struct.pack("<I", 0)
    text = text[:m.start(2)] + base64.b64encode(new).decode("ascii") + text[m.end(2):]
    return b"\xff\xfe" + text.encode("utf-16-le")


def _fix_vlookup(s):
    pat = re.compile(r"VLOOKUP\((raw_data\[\[#This Row\],\[A/S번호\]\]),as_list\[\],(\d)\)")
    return pat.subn(lambda m: 'IFERROR(VLOOKUP(%s,as_list[],%s,FALSE),"미분류")' % (m.group(1), m.group(2)), s)


def _pick_base(root, out):
    c = []
    for d in (root, os.path.join(root, "엑셀보고서")):
        c += [p for p in glob.glob(os.path.join(d, "CLAIM_AS_보고서_*.xlsx")) if not os.path.basename(p).startswith("~$")]
    c = [p for p in c if os.path.abspath(p) != os.path.abspath(out)]
    return max(c, key=os.path.getmtime) if c else None


def run(root, raw, raw_cols, lk, list_cols, mode, latest):
    ly, lm = int(latest[:4]), int(latest[5:7])
    d = os.path.join(root, "엑셀보고서")
    raw_dir, list_dir = os.path.join(d, "raw_data"), os.path.join(d, "as_list")
    for folder in (raw_dir, list_dir):  # 쿼리가 폴더의 모든 파일을 읽으므로 이전 파일은 지운다
        os.makedirs(folder, exist_ok=True)
        for p in glob.glob(os.path.join(folder, "*")):
            try:
                os.remove(p)
            except OSError as e:
                print(f"[경고] {p} 를 지우지 못했습니다 ({e}). 엑셀에서 열려 있으면 닫아 주세요.")
    r = raw.sort_values("_dt", kind="stable")
    n1 = write_source(os.path.join(raw_dir, "as_claim_raw_data.xlsx"), r, raw_cols, RAW_NUM, "Column33")
    n2 = write_source(os.path.join(list_dir, "as_list.xlsx"), lk, list_cols, LIST_NUM, "Column24")
    print(f"원본 데이터 파일: raw {n1}행, list {n2}행 -> {d}")

    out = os.path.join(d, f"CLAIM_AS_보고서_{ly}{lm:02d}.xlsx")
    base = _pick_base(root, out)
    if not base:
        print("[경고] 기준이 될 기존 엑셀 보고서(CLAIM_AS_보고서_*.xlsx)가 없어 새 보고서는 만들지 않았습니다.")
        return None
    if os.path.exists(out):  # 같은 달에 다시 실행해도 직접 수정한 파일을 잃지 않도록 백업
        bk = os.path.join(d, "백업")
        os.makedirs(bk, exist_ok=True)
        try:
            shutil.move(out, os.path.join(bk, f"{datetime.now():%Y%m%d_%H%M%S}_{os.path.basename(out)}"))
        except OSError as e:
            print(f"[경고] 기존 {os.path.basename(out)} 가 열려 있어 새 보고서를 저장하지 못했습니다 ({e}).")
            return None
    tmp = os.path.join(tempfile.gettempdir(), "claim_as_base.xlsx")
    shutil.copyfile(base, tmp)  # OneDrive 동기화 중 잠금을 피하기 위해 복사본을 읽는다
    fixed = 0
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename
            if name.startswith("customXml/item") and name.endswith(".xml") and b"D\x00a\x00t\x00a\x00M\x00a\x00s\x00h\x00u\x00p" in data:
                data = _patch_mashup(data, raw_dir, list_dir)
            elif mode != "excel" and name in ("xl/worksheets/sheet3.xml", "xl/tables/table2.xml"):
                s, k = _fix_vlookup(data.decode("utf-8"))
                data, fixed = s.encode("utf-8"), fixed + k
            elif name == "xl/workbook.xml":
                s = data.decode("utf-8")
                s = re.sub(r"<calcPr([^>]*?)/>", lambda m: m.group(0) if "fullCalcOnLoad" in m.group(0) else "<calcPr%s fullCalcOnLoad=\"1\"/>" % m.group(1), s)
                data = s.encode("utf-8")
            zout.writestr(info, data)
    os.remove(tmp)
    print(f"엑셀 보고서: {out}  (기준: {os.path.basename(base)}, VLOOKUP {fixed}곳 수정)")
    return out
