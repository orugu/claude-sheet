# claude-sheets-cli

Claude Code(터미널)가 Google Sheets를 직접 읽고 쓰기 위한 로컬 CLI 도구.
브라우저 자동화 없이 Google Sheets API v4를 직접 호출한다.

## 1. Google Cloud 설정 (최초 1회, 사용자가 직접)

1. https://console.cloud.google.com 접속 → 새 프로젝트 생성
2. **API 및 서비스 → 라이브러리** → "Google Sheets API" 검색 → 사용 설정
3. **API 및 서비스 → OAuth 동의 화면**
   - User Type: 외부(테스트 단계로 충분)
   - 앱 이름/이메일 등 최소 정보만 입력
   - 범위(Scopes)는 기본값으로 두고 넘어가도 됨
   - **테스트 사용자**에 본인 Google 계정 이메일 추가 (안 하면 로그인 시 차단됨)
4. **API 및 서비스 → 사용자 인증 정보 → 사용자 인증 정보 만들기 → OAuth 클라이언트 ID**
   - 애플리케이션 유형: **데스크톱 앱**
   - 생성 후 나오는 **클라이언트 ID**, **클라이언트 보안 비밀번호** 복사

## 2. 로컬 설정

```bash
cp config.example.json config.json
# config.json 열어서 client_id / client_secret 붙여넣기

node auth.js
# 터미널에 뜨는 URL을 브라우저에서 열고 Google 계정 로그인/동의
# 완료되면 token.json이 생성됨 (최초 1회만 하면 됨, 이후 자동 갱신)
```

## 3. 사용법

```bash
# 값 읽기
node sheets.js get <spreadsheetId> "Sheet1!A1:D10"

# 여러 범위 한번에 읽기
node sheets.js batchGet <spreadsheetId> '["Sheet1!A1:B2","Sheet1!D1:D5"]'

# 값/수식 쓰기 (USER_ENTERED라 "=SUM(A1:A5)" 같은 수식도 그대로 인식됨)
node sheets.js update <spreadsheetId> "Sheet1!A1:B2" '[["이름","점수"],["철수",90]]'

# 행 추가
node sheets.js append <spreadsheetId> "Sheet1!A1" '[["새 행","값"]]'

# 범위 지우기
node sheets.js clear <spreadsheetId> "Sheet1!A1:D10"

# 시트 구조 확인 (시트 목록, ID 등)
node sheets.js metadata <spreadsheetId>

# 서식/차트/병합 등 고급 조작 — Sheets API batchUpdate 요청을 그대로 전달
node sheets.js batchUpdate <spreadsheetId> '[{"repeatCell":{...}}]'
```

`spreadsheetId`는 시트 URL 중간의 긴 문자열이다:
`https://docs.google.com/spreadsheets/d/`**`이_부분`**`/edit`

## 참고

- `config.json`, `token.json`에는 민감정보가 들어있음 — git에 커밋 금지(.gitignore 처리됨)
- 스코프는 스프레드시트 전체 쓰기 + 드라이브 읽기전용. 필요시 `auth.js`의 SCOPES 수정 후 재인증
