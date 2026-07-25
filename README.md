# Mibang Poodle Campaign Dashboard

Mibang 전용 푸들 캠페인 운영 사이트입니다.

## 사이트

GitHub Pages 활성화 후:

- https://parkwbjp.github.io/mibang/

## Google Sheets

- https://docs.google.com/spreadsheets/d/1kuzNyMSObduViqgl0uVYcKg5cJs7in1RR1budvcXj84/edit

## 데이터 연동

1. 위 Google Sheet를 엽니다.
2. `확장 프로그램` → `Apps Script`
3. 이 저장소의 `apps-script.gs` 내용을 붙여넣습니다.
4. `배포` → `새 배포` → `웹 앱`
5. 실행 사용자: `나`, 액세스 사용자: `모든 사용자`
6. 생성된 `/exec` URL을 사이트의 `연동 설정`에 저장합니다.
7. 사이트에서 `업무 종료·Google Sheets 저장` 버튼을 누릅니다.

업로드 대상:

- `DM_LIST`: DM 대상과 상태
- `TASK_STATUS`: 날짜별 완료 업무
- `DAILY_LOG`: 일일 운영 실적
