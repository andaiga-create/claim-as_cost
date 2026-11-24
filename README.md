# CLAIM · AS 비용 대시보드

ERP에서 내려받은 A/S·CLAIM 매입 내역(xls)을 집계해 GitHub Pages로 보여주는 정적 대시보드입니다.

- 사이트: https://andaiga-create.github.io/claim-as_cost/
- 저장소: https://github.com/andaiga-create/claim-as_cost

---

## 1. 매월 데이터 업데이트 절차

### ① 파일 준비
ERP에서 해당 월 자료를 내려받아 **파일명을 아래 형식으로 맞춰** 넣습니다. (`YYYYMM` = 대상 월)

| 폴더 | 파일명 | 내용 |
|---|---|---|
| `data/raw_data/` | `as_claim_raw_data_YYYYMM.xls` | 매입(입고) 내역 |
| `data/as_list/` | `as_list_YYYYMM.xls` | A/S·클레임 접수 목록 |

예) 11월 자료 → `as_claim_raw_data_202611.xls`, `as_list_202611.xls`

> 파일명 형식이 다르면 월별 교체로 인식되지 않아 데이터가 중복될 수 있습니다.
> 같은 월 파일을 다시 넣어도 해당 월은 **교체**되므로 중복되지 않습니다.

### ② (필요할 때만) 목표·매출 수정
`docs/settings.json` 을 메모장 등으로 열어 수정합니다. (자세한 항목은 아래 3번)

### ③ `update.bat` 실행
프로젝트 폴더의 **`update.bat`을 더블클릭**합니다. 순서대로 자동 수행됩니다.

1. xls 변환 → `docs/data.json` 생성 + **엑셀 보고서 생성** (`엑셀보고서/CLAIM_AS_보고서_YYYYMM.xlsx`, 아래 2-1)
2. `git commit`
3. `git push` (GitHub 업로드)

"Done" 메시지가 나오면 완료입니다. **1~2분 뒤** 사이트에서 `Ctrl+F5`로 새로고침하세요.

> 변환 결과만 점검하고 싶으면 명령 프롬프트에서 `update.bat build` 로 실행합니다. (git 업로드 생략, 엑셀 보고서는 생성됨)

---

## 2. 폴더 구조

```
README.md
update.bat              # 변환 + 업로드 (더블클릭)
scripts/build_data.py   # xls → docs/data.json 변환
scripts/build_excel.py  # 엑셀 보고서(원본 데이터 + 새 파일) 생성
data/                   # 원본 xls (PC에만 보관, GitHub 에는 올라가지 않음)
  raw_data/             #   as_claim_raw_data.xls(누적) + _YYYYMM.xls(월별)
  as_list/              #   as_list_2020-20260228.xls(누적) + _YYYYMM.xls(월별)
엑셀보고서/              # update.bat 이 만드는 엑셀 보고서 (GitHub 에는 올라가지 않음)
docs/                   # 웹사이트 (GitHub Pages 가 이 폴더를 게시)
  index.html, style.css, app.js
  data.json             #   변환 결과(집계값)
  settings.json         #   매출·목표 설정
```

- `data/` 폴더는 `.gitignore` 로 제외되어 **원본 xls는 업로드되지 않습니다.** 사이트에는 집계값(`data.json`)만 공개됩니다.
- GitHub Pages 설정: Settings → Pages → Branch `main` / 폴더 `/docs`.

---

## 2-1. 엑셀 보고서 (`엑셀보고서/CLAIM_AS_보고서_YYYYMM.xlsx`)

`update.bat` 을 실행하면 **기존 엑셀 보고서를 복사한 새 파일**이 만들어집니다. 피벗테이블·슬라이서·차트·데이터 모델·dashboard 는 그대로이고, 아래 3가지만 바뀝니다.

1. **파워쿼리 원본 데이터 생성** — 웹과 같은 최종 데이터(월별 병합, 같은 월 재업로드 교체, A/S번호 하이픈 보정 반영)를 쿼리가 읽는 형식으로 저장합니다.
   - `엑셀보고서/raw_data/as_claim_raw_data.xlsx`, `엑셀보고서/as_list/as_list.xlsx`
   - 새 보고서의 쿼리 원본 폴더(기존 `Z:\...aw_data`, `as_list`)가 위 두 폴더로 바뀌어 있습니다.
2. **VLOOKUP 오류 수정** — raw 시트의 불량코드·불량명칭·유/무상 조회를 `IFERROR(VLOOKUP(…,FALSE),"미분류")`(정확 일치)로 바꿉니다. 웹과 같은 금액이 나옵니다.
3. 파일 열 때 수식을 다시 계산하도록 설정

**새 보고서를 연 뒤 [데이터] → [모두 새로 고침]** 을 한 번 눌러야 raw/list 와 피벗이 새 데이터로 바뀝니다. (파일을 만들 때는 이전 데이터가 들어 있고, 개인 정보 수준 확인 창이 뜨면 허용)

- 기준 파일: `엑셀보고서/` 와 프로젝트 폴더의 `CLAIM_AS_보고서_*.xlsx` 중 **가장 최근에 수정한 파일**. dashboard 를 직접 고쳐 저장해 두면 다음 달에는 그 파일이 기준이 됩니다.
- 같은 달에 `update.bat` 을 다시 실행하면 기존 파일은 `엑셀보고서/백업/` 으로 옮겨집니다. 새 보고서가 엑셀에서 열려 있으면 닫고 실행하세요.
- 기준 보고서가 없으면 원본 데이터 파일만 만들고 새 보고서는 만들지 않습니다.
- `lookup_mode` 가 `excel` 이면 VLOOKUP 은 수정하지 않습니다.
- 필요한 패키지: `pip install pandas xlrd xlsxwriter` (update.bat 이 자동 설치). 엑셀 단계만 건너뛰려면 `python scriptsuild_data.py --no-excel`.

---

## 3. 목표·매출 설정 (`docs/settings.json`)

```json
{
  "lookup_mode": "exact",
  "ratio_basis": "avg3",
  "revenue": { "2025": 17919000000, "2026": 24497611317 },
  "kpi": { "2026": { "cost_target": null, "ratio_target": 0.007 } }
}
```

| 항목 | 설명 |
|---|---|
| `lookup_mode` | 접수 목록 조회 방식. `"exact"` = A/S번호 정확 일치(기본, 안정적), `"excel"` = 엑셀 VLOOKUP 근사 일치 재현. 아래 4번 참고 |
| `revenue` | 연도별 매출액(원). 새 연도가 되면 추가 |
| `ratio_basis` | `"avg3"` = 최근 3개년 평균 매출 기준(엑셀과 동일), `"year"` = 해당 연도 매출 기준 |
| `kpi.연도.ratio_target` | 매출 대비 비용 목표. 0.7% → `0.007`. 쓰지 않으면 `null` |
| `kpi.연도.cost_target` | 연간 비용 목표 금액(원). 입력하면 누적 차트에 목표선 표시. 쓰지 않으면 `null` |

수정 후 `update.bat` 을 실행하면 반영됩니다. (xls 가 그대로여도 상관없습니다.)

---

## 4. 집계 규칙 (엑셀 보고서와 동일)

- **고객 / 호선 / 장비**: 프로젝트코드(`HHI-P167-HOC-S1`)를 `-` 로 분해 → `HHI` / `P167` / `HOC`
- **선종**: 프로젝트코드 끝 2자리 — `C1`·`C2` = 상선, `S1` = 특수선, 그 외 = 기타
- **불량코드 / 불량명칭 / 유·무상**: A/S번호로 접수 목록(`as_list`)을 조회
- **A/S · CLAIM 구분**: 불량코드가 `A`로 시작하면 A/S, `C`로 시작하면 CLAIM, 그 외 미분류
- 대시보드 기본 필터: 최신 연도 · 무상 · A/S+CLAIM · 상선+특수선 (엑셀 피벗과 같은 조건)
- 좌측 월(1~12월) 버튼: 여러 달 선택 가능, 선택한 연도에 데이터가 없는 달은 비활성. 월 추이 차트는 항상 12개월을 보여주고 선택하지 않은 달만 흐리게 표시

### 접수 목록 조회 방식 (`lookup_mode`)

| 값 | 동작 | 장단점 |
|---|---|---|
| `"excel"` | 엑셀 `VLOOKUP`(4번째 인수 생략 = 근사 일치)을 그대로 재현 | 엑셀 보고서와 숫자가 일치. 단, 접수 목록이 완전히 정렬돼 있지 않아 **월 파일을 추가하면 과거 일부 건의 분류가 달라질 수 있음** |
| `"exact"` (현재 기본) | A/S번호가 정확히 같은 접수 건만 연결 | 매달 과거 값이 변하지 않음. 접수 목록에 없는 A/S번호는 `미분류`가 되며, 기본 필터(A/S+CLAIM, 무상)에서는 제외되어 합계가 엑셀보다 적게 보임 |

변경은 `docs/settings.json` 의 `lookup_mode` 값을 바꾸고 `update.bat` 을 실행하면 됩니다. 현재 방식은 대시보드 좌측 하단에 표시됩니다.

---

## 5. 문제 해결

| 증상 | 확인 |
|---|---|
| 더블클릭해도 창이 바로 꺼짐 | 명령 프롬프트에서 `update.bat` 을 실행해 메시지 확인 |
| `No usable Python` | Python 3 설치 후 `pip install pandas xlrd` |
| `Conversion failed` | `data/` 안 xls 파일이 열려 있는지(엑셀에서 열어둔 경우), 파일명 형식 확인 |
| `Push failed` | 인터넷 연결, GitHub 로그인(자격 증명) 확인 |
| 사이트에 최신 데이터가 안 보임 | 1~2분 대기 후 `Ctrl+F5` |
| 숫자가 엑셀과 다름 | 엑셀 `VLOOKUP` 이 아직 근사 일치(4번째 인수 없음)이면 값이 다름 → 엑셀 수식에 `FALSE` 추가, 또는 `lookup_mode` 를 `excel` 로. 엑셀 보고서에 해당 월 자료를 모두 붙여넣었는지도 확인 |

---

## 6. 미분류가 생기는 이유와 해결

`exact` 모드에서 `미분류`가 되는 건은 매입 내역의 A/S번호가 접수 목록(`as_list`)에 없는 경우입니다.
현재 130건 중 약 100건이 2017~2019년 접수번호입니다. 접수 목록이 2020년부터라서 생깁니다.
**2019년 이전 접수 목록(`as_list_YYYY.xls`)을 `data/as_list/` 에 추가**하면 연결되어 미분류가 줄어듭니다.

`update.bat` 이 만드는 엑셀 보고서는 이미 정확 일치로 수정돼 있습니다. 기존 보고서를 직접 고치려면 `raw` 시트의 조회 수식 3곳(불량코드·불량명칭·유/무상)에 `FALSE` 를 붙입니다.

```
=IFERROR(VLOOKUP(raw_data[[#This Row],[A/S번호]],as_list[],5,FALSE),"미분류")   ' 불량코드
=IFERROR(VLOOKUP(raw_data[[#This Row],[A/S번호]],as_list[],6,FALSE),"미분류")   ' 불량명칭
=IFERROR(VLOOKUP(raw_data[[#This Row],[A/S번호]],as_list[],8,FALSE),"미분류")   ' 유/무상
```
