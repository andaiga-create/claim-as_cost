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

1. xls 변환 → `docs/data.json` 생성
2. `git commit`
3. `git push` (GitHub 업로드)

"Done" 메시지가 나오면 완료입니다. **1~2분 뒤** 사이트에서 `Ctrl+F5`로 새로고침하세요.

> 변환 결과만 점검하고 싶으면 명령 프롬프트에서 `update.bat build` 로 실행합니다. (git 업로드 생략)

---

## 2. 폴더 구조

```
README.md
update.bat              # 변환 + 업로드 (더블클릭)
scripts/build_data.py   # xls → docs/data.json 변환
data/                   # 원본 xls (PC에만 보관, GitHub 에는 올라가지 않음)
  raw_data/             #   as_claim_raw_data.xls(누적) + _YYYYMM.xls(월별)
  as_list/              #   as_list_2020-20260228.xls(누적) + _YYYYMM.xls(월별)
docs/                   # 웹사이트 (GitHub Pages 가 이 폴더를 게시)
  index.html, style.css, app.js
  data.json             #   변환 결과(집계값)
  settings.json         #   매출·목표 설정
```

- `data/` 폴더는 `.gitignore` 로 제외되어 **원본 xls는 업로드되지 않습니다.** 사이트에는 집계값(`data.json`)만 공개됩니다.
- GitHub Pages 설정: Settings → Pages → Branch `main` / 폴더 `/docs`.

---

## 3. 목표·매출 설정 (`docs/settings.json`)

```json
{
  "ratio_basis": "avg3",
  "revenue": { "2025": 17919000000, "2026": 24497611317 },
  "kpi": { "2026": { "cost_target": null, "ratio_target": 0.007 } }
}
```

| 항목 | 설명 |
|---|---|
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

> 접수 목록 조회는 엑셀의 `VLOOKUP`(근사 일치) 동작을 그대로 재현합니다. 엑셀과 숫자를 맞추기 위한 것이며, 그 특성상 **월 파일을 추가하면 과거 일부 건의 분류가 달라질 수 있습니다.**

---

## 5. 문제 해결

| 증상 | 확인 |
|---|---|
| 더블클릭해도 창이 바로 꺼짐 | 명령 프롬프트에서 `update.bat` 을 실행해 메시지 확인 |
| `No usable Python` | Python 3 설치 후 `pip install pandas xlrd` |
| `Conversion failed` | `data/` 안 xls 파일이 열려 있는지(엑셀에서 열어둔 경우), 파일명 형식 확인 |
| `Push failed` | 인터넷 연결, GitHub 로그인(자격 증명) 확인 |
| 사이트에 최신 데이터가 안 보임 | 1~2분 대기 후 `Ctrl+F5` |
| 숫자가 엑셀과 다름 | 엑셀 보고서에 해당 월 자료를 모두 붙여넣었는지 확인 |
