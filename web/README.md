# FEED M/H Calculator — HTML Web App

`FEED_MH_Calculator_V0_4_OPExcelGrid.py` (Tkinter 데스크톱 프로그램)와 **같은 계산 결과**를 내는
브라우저용 버전입니다. 설치·파이썬·서버 없이 동작합니다.

## 실행 방법

`web/index.html` 을 브라우저에서 열면 됩니다. (더블클릭 또는 브라우저로 끌어놓기)

사내 서버 등에 올려서 쓰려면 `web` 폴더를 그대로 정적 호스팅하면 됩니다.

```
python -m http.server 8000 --directory web    # 예시
```

## 구성

| 파일 | 내용 |
| --- | --- |
| `index.html` | 화면 뼈대 (Master Control + 10개 탭) |
| `css/styles.css` | 데스크톱 프로그램과 동일한 색상 팔레트 / Excel 형태 표 서식 |
| `js/data.js` | 파이썬 소스에서 자동 추출한 기준 데이터 (Input 96행, Output 54행, 산출기준 CI 32 / TEL 17행) + 로고 |
| `js/model.js` | `Model` 클래스 이식 (난이도 산정, 수량 자동계산, 내부/외주 배분, Case 로직) |
| `js/app.js` | 탭 렌더링, 입력 처리, 상태 저장/복원 |
| `js/report.js` | Word Report 생성 (`report_utils.py` 구성과 동일) |

## 탭 구성 (데스크톱과 동일)

`Input 수정` · `Guide / Help` · `Summary` · `Output_CI` · `Output_TEL` ·
`OP1` · `OP2-단종` · `OP2-종합` · `산출기준_CI` · `산출기준_TEL`

## 계산 결과 동일성

파이썬 `Model` 과 자바스크립트 `Model` 을 다음 조건으로 교차 검증했습니다.

- Part(전체/CI/TEL) × Case(단종/종합) × 비율(원본/외주최소화/사용자) × FEED 난이도 3종 = **57개 시나리오**
- 각 시나리오의 **모든 Activity 행**(수량·난이도·내부/외주 Unit M/H·내부/외주/Total M/H)과 합계·비율 → 전부 일치
- 숫자 표기도 파이썬 `format(v, ',.0f')` / `',.1f'` / `',.2f'` 과 동일 (정확히 .5 인 값의 짝수 반올림 포함)

## 데스크톱 프로그램과 달라진 점

| 항목 | 데스크톱 | 웹 |
| --- | --- | --- |
| 입력값 저장 | 프로그램 폴더의 `FEED_MH_Calculator_Last_Input.json` | 브라우저 localStorage 자동 저장 + JSON 파일 내보내기/가져오기 |
| Word Report | `python-docx` 로 `.docx` 생성 | Word 가 그대로 여는 `.doc`(HTML) 파일 다운로드. 차트는 표 기반 막대그래프로 대체 |
| 산출기준 값 수정 | 셀 더블클릭 후 입력 | 셀을 바로 클릭해 입력 |
| 산출기준 탭의 Version 표기 | `일반 Ver.` 고정 | 외주최소화 값에 따라 `일반 Ver.` / `외주최소화 Ver.` 표시 (다른 탭과 동일하게 맞춤) |
| OP / 산출기준 그리드 | Tkinter Canvas 직접 그리기 | HTML 표 (같은 열 구성·색상·머리글 구조, 화면 폭에 맞춰 가로 스크롤) |

내보낸 JSON 은 데스크톱 프로그램의 `FEED_MH_Calculator_Last_Input.json` 과 형식이 같아
양쪽에서 서로 불러올 수 있습니다.

## 데이터 갱신

`FEED_MH_Calculator_V0_4_OPExcelGrid.py` 의 기준 데이터가 바뀌면 `js/data.js` 를 다시 만들어야 합니다.

```bash
python3 tools/gen_web_data.py
```
