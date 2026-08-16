# 방탄한어른등 길드 포털

RF 온라인 넥스트 길드 관리 포털. 화면은 GitHub Pages, 데이터는 Google Sheets에 저장됩니다.

## 구조

```
브라우저 ──> GitHub Pages (index.html)  ← 화면·디자인
                    │ fetch
                    ▼
          Apps Script 웹앱 (/exec)      ← 인증·읽기·쓰기
                    │
                    ▼
             Google Sheets              ← 데이터 원본
```

## 설치

### 1. Apps Script 설정
1. 구글 시트 열기 → 확장 프로그램 → Apps Script
2. `guild_sync_v20.gs` 내용을 Code.gs에 붙여넣기
3. 상단 값 확인·수정
   - `SPREADSHEET_ID` : 시트 주소의 `/d/` 와 `/edit` 사이 문자열
   - `SECRET` : 아무 문자열로 변경 (토큰 서명용, 이후 변경 금지)
   - `PW` / `PW_MEMBER` : 운영진 / 길드원 비밀번호
4. 배포 → 새 배포 → 유형 **웹 앱** → 실행 계정 **나**, 액세스 **모든 사용자** → 배포
5. 웹 앱 URL(`.../exec`) 복사

> 이후 코드를 고칠 때는 **배포 → 배포 관리 → 연필(수정) → 새 버전 → 배포**
> ("새 배포"를 누르면 주소가 바뀝니다)

### 2. GitHub Pages 설정
1. 새 저장소 생성 (Public)
2. `index.html` 업로드
3. `index.html` 상단의 `SYNC_URL` 을 위에서 복사한 `/exec` 주소로 수정
4. Settings → Pages → Source: `Deploy from a branch` → Branch: `main` / `(root)` → Save
5. 1~2분 후 `https://<사용자명>.github.io/<저장소명>/` 접속

## 사용

| 구분 | 비밀번호 | 권한 |
|---|---|---|
| 운영진 | (Code.gs의 `PW`) | 모든 탭 수정, 설정 변경 |
| 길드원 | (Code.gs의 `PW_MEMBER`) | 허용된 탭만 수정 (기본: 스킬·마유) |

자세한 사용법은 포털 내 **설명서** 탭 참조.

## 업데이트

- 화면 수정 → GitHub에서 `index.html` 수정·커밋 → 1분 내 자동 반영
- 서버 로직 수정 → Apps Script 수정 → **새 버전으로 재배포**

## 백업

- 구글 시트 → 파일 → 사본 만들기 (권장: 큰 변경 전마다)
- 포털 하단 **변경사항 내보내기(JSON)** 버튼
