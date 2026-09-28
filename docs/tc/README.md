# 테스트케이스(TC)

| 파일 | 내용 |
|---|---|
| `포트폴리오_길드포탈_TC_260928_v2.xlsx` | TC 문서 (기본기능·스모크·Sanity·BAT·BVT·회귀 6개 탭) |
| `src/tc_basic.json` | 기본기능 TC 원본 데이터 |
| `src/tc_others.json` | 나머지 5개 탭 원본 데이터 |
| `src/cfg.json` | 공통 문구(사전조건 약어 등)·Read Me 내용·입출력 경로 |
| `src/build.ps1` | JSON → xlsx 생성 스크립트 (Windows + Excel 필요) |

## 수정 방법
1. `src/*.json`에서 TC를 추가·수정한다 (변경 이력은 JSON diff로 확인).
2. `cfg.json`의 `src`(양식 원본)·`dst`(출력 경로)를 확인한다.
3. PowerShell에서 실행한다: `powershell -ExecutionPolicy Bypass -File docs\tc\src\build.ps1`
4. 생성된 xlsx를 이 폴더에 덮어쓰고 함께 커밋한다.

## JSON 항목
`id` ID · `c/d/e` 대/중/소분류 · `p` 우선순위 · `pre` 사전조건 · `st` 테스트 단계(배열) · `ex` 기대결과(배열, 1항목=1행) · `ref` 참고사항 · `tq` 비고(QA 기법)
`pre`/`st`의 `@ADM`, `@SAVE` 같은 약어는 `cfg.json`의 `tokens`에서 문장으로 바뀐다.
