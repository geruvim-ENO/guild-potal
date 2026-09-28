# 작업 규칙

## 수정·커밋
- 수정 요청은 제안만 하지 말고 파일을 직접 수정한다.
- 수정 후 작업 단위마다 로컬에서 직접 커밋한다.
- 이 PC의 PATH에는 git이 없으므로 GitHub Desktop 내장 git을 사용한다:
  `C:\Users\geruv\AppData\Local\GitHubDesktop\app-*\resources\app\git\cmd\git.exe`

## push 전 절차
1. 직접 테스트한다.
   - `index.html`: 앱 내 브라우저로 열어 화면·동작 확인
   - `apps-script.gs`: 문법 및 로직 점검 (실제 Google Sheets 연동은 로컬에서 실행 불가 — 확인 범위를 명시)
2. 테스트 항목별 결과(통과/실패)를 사용자에게 보고한다.
3. 실패 항목은 수정 → 재테스트를 모두 통과할 때까지 반복한다.
4. 테스트 내용과 결과를 노션에 기록한다.
   - Notion 커넥터 사용. 기록 페이지: 「길드 포털」 https://app.notion.com/p/3e9c9259b7cd81108f27fa00cbece668
   - 「테스트 기록」 섹션에 회차별(날짜·커밋·항목별 결과)로 최신이 위에 오도록 추가
5. 사용자에게 push 허락을 받은 뒤 직접 `git push` 한다.
