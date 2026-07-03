# EASTEND 오프라인 멤버십 (매장 QR 가입)

매장에 비치한 QR을 고객이 스캔 → 모바일 가입 폼(개인정보 동의 포함) → Google Sheet에 저장 → 관리자 페이지에서 조회/CSV 다운로드.

## 구성

| 구분 | 위치 | 설명 |
|---|---|---|
| 가입 폼 | `docs/index.html` | `?store=C001` 파라미터로 매장 자동 인식, 브랜드별 테마 |
| QR 포스터 | `docs/qr.html` | 매장별 QR 생성 + A4 인쇄/PDF 저장 |
| 관리자 | `docs/admin.html` | 토큰 로그인, 통계/검색/필터, CSV 다운로드 |
| 매장 마스터 | `docs/stores.js` | 매장 추가/삭제는 이 파일만 수정 |
| 서버 설정 | `docs/config.js` | Apps Script 웹앱 URL |
| 백엔드 | `gas/Code.gs` | Google Apps Script (시트 저장 + 조회 API) |

**개인정보는 GitHub에 절대 저장되지 않음.** GitHub Pages는 폼/화면(코드)만 제공하고, 고객 데이터는 비공개 Google Sheet에만 저장된다.

## 배포 순서

### 1. Google Sheet + Apps Script (데이터 저장소)
1. 새 Google 스프레드시트 생성 (예: `오프라인 멤버십 회원DB`) — **비공개 유지**
2. 확장 프로그램 → Apps Script → `gas/Code.gs` 내용 전체 붙여넣기
3. `setup` 함수 1회 실행 → 로그에 출력되는 **관리자 토큰** 보관
4. 배포 → 새 배포 → 웹 앱
   - 실행 계정: **나**
   - 액세스 권한: **모든 사용자(익명 포함)** ← 고객이 로그인 없이 제출해야 하므로 필수
5. 웹앱 URL(`https://script.google.com/macros/s/.../exec`) 복사

### 2. GitHub Pages (프론트)
1. `docs/config.js`의 `API_URL`에 웹앱 URL 붙여넣기
2. GitHub에 push → Settings → Pages → Branch: `main`, Folder: `/docs`
3. 접속 URL: `https://wonsosw-cmd.github.io/eastend-members/`

### 3. QR 배포
- `https://wonsosw-cmd.github.io/eastend-members/qr.html` 접속 → 매장 선택 → 인쇄/PDF 저장
- 또는 `py -3.14 tools/make_qr.py` 로 PNG 일괄 생성

## 운영

- **매장 추가/삭제**: `docs/stores.js` 수정 후 push (QR·폼·관리자 모두 자동 반영)
- **관리자 접속**: `admin.html` → setup에서 발급된 토큰 입력
- **중복 가입**: 같은 전화번호 재제출 시 신규 등록 대신 최근방문/방문횟수 업데이트
- **Code.gs 수정 시**: Apps Script에서 배포 → 배포 관리 → 기존 배포 **수정**(새 버전) 으로 올려야 URL이 유지됨

## 개인정보 유의사항

- 시트/스크립트는 회사 계정 소유로 유지하고 링크 공유 금지
- 수집 항목: 이름, 휴대전화(필수) / 생년월일 6자리, 성별(선택) + 동의 여부·일시 증적
- 고객이 동의 철회 요청 시 시트에서 해당 행 삭제(파기)로 처리
