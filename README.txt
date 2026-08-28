FEED M/H Calculator V0.4 OPExcelGrid / 2026-08-11

- OP1 / OP2-단종 / OP2-종합 탭을 Excel-like Grid UI로 변경
- 엑셀 원본처럼 제목, Version, Base Data Summary, Notes, Project, 다단 헤더 구성
- Activity를 GENERAL / SPECIFICATION / CALCULATION / 구매관련 / DRAWING / MTO / OTHERS 섹션 Tree 구조로 표시
- 내부/외주/Total M/H 색상 구분 및 하단 Summary 추가
- 산출기준_CI/TEL Excel Grid와 Word Report 기능 유지


[Web (HTML) version]
- MH_Calculator.html is a single self-contained file: open it in a browser and it runs.
  No installation, no Python, no server, and no external files - styles, calculation model,
  report generator, base data and logo are all inside that one file.
- Works from a local disk, a USB stick, an email attachment or a static web server.
- The tab structure (Edit Inputs / Phase Split / Guide / Summary / Output_CI / Output_TEL /
  OP1 / OP2-Single / OP2-Comprehensive / Standards_CI / Standards_TEL) and the results are
  identical to the desktop program. The interface is in English.
- Titles and labels name no project stage, since the same calculation is used for FEED work,
  detail engineering and bid estimates.
- Phase Split is web-only: it assigns each activity a share of phase 1 (0-100%) so a job that
  runs in two stages can be split, with a duration and a headcount per phase. Name the two
  phases yourself in Master Control. Leaving every activity at 100 reproduces the desktop
  program's figures exactly.
- Inputs are saved in the browser automatically, and JSON export/import lets the web app
  and the desktop program exchange input files.
- See WEB_APP.md for details.

[Detail Engineering / Proposal web apps]
- DE_MH_Calculator.html and Proposal_MH_Calculator.html are built from the 실행사업 Rev.3
  Excel workbooks (MH_Calculator_Detail Engineering.xlsx / MH_Calculator_Propoal.xlsx).
- Each is a single self-contained file, like MH_Calculator.html.
- They do not re-implement the workbook: they carry its formula graph and evaluate it in the
  browser, so every figure matches Excel cell for cell - for any quantities entered, not only
  the sample project. Both were checked against every cached value in their workbook: 0
  mismatches.
- They are separate apps because the two workbooks are genuinely different: the proposal book
  carries 49 hand-edited formulas the execution book does not. No setting turns one into the
  other.
- Tabs: Summary / OP1_CI / OP1_TEL / OP2-1 / OP2-2 / OP2-3 / Input_CI / Input_TEL /
  Standards_CI / Standards_TEL / Guide.
- The Input tabs are drawn as the workbook's three side-by-side blocks, with its own headers,
  its own dropdown lists on the Project Condition selectors, and the Excel row number on every
  row. Calculated cells are shown but not editable.
- See DE_APPS.md for details.
