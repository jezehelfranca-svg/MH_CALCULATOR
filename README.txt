FEED M/H Calculator V0.4 OPExcelGrid / 2026-08-11

- OP1 / OP2-단종 / OP2-종합 탭을 Excel-like Grid UI로 변경
- 엑셀 원본처럼 제목, Version, Base Data Summary, Notes, Project, 다단 헤더 구성
- Activity를 GENERAL / SPECIFICATION / CALCULATION / 구매관련 / DRAWING / MTO / OTHERS 섹션 Tree 구조로 표시
- 내부/외주/Total M/H 색상 구분 및 하단 Summary 추가
- 산출기준_CI/TEL Excel Grid와 Word Report 기능 유지

[웹(HTML) 버전]
- web/index.html 을 브라우저에서 열면 설치 없이 같은 계산을 사용할 수 있습니다.
- 탭 구성(Input 수정 / Guide / Summary / Output_CI / Output_TEL / OP1 / OP2-단종 / OP2-종합 /
  산출기준_CI / 산출기준_TEL)과 계산 결과는 데스크톱 프로그램과 동일합니다.
- 입력값은 브라우저에 자동 저장되며, JSON 내보내기/가져오기로 데스크톱 프로그램과 주고받을 수 있습니다.
- 자세한 내용은 web/README.md 참고
