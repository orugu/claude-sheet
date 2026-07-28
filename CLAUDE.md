# claude-sheets-cli

Claude Code(터미널)에서 브라우저 없이 Google Sheets API v4를 직접 호출하는 로컬 CLI 도구.
`sheets.js` 하나로 값 읽기/쓰기부터 서식·병합·열너비까지 다룬다. 명령/플래그 전체 목록은
`sheets.js` 맨 위 주석과 `README.md` 참고. 이 파일은 **실전에서 겪은 함정과 정해진 작업 순서**를 정리한 것.

## 왜 이 도구를 쓰는가

Google Drive MCP 커넥터가 자주 끊기거나("Server not found") 세션이 만료돼서, OAuth로 직접
Sheets API를 호출하는 이 CLI를 만들었다. `~/.claude-sheets-cli` 아래에서 `node sheets.js <command> ...`
형태로 쓴다. 인증은 `node auth.js`로 최초 1회만 하면 `token.json`이 자동 갱신된다.

## 반드시 지킬 순서/규칙

### 1. 쓰기 전에 항상 `tabs`로 시트ID/이름 확인, 쓴 다음엔 항상 `get`으로 검증
도구가 "성공"이라고 응답해도 실제 셀 내용·서식을 `get`으로 다시 읽어서 확인하기 전엔 사용자에게
완료라고 보고하지 않는다. (다른 hwp 작업에서도 같은 원칙 — 도구 자체의 성공 응답을 못 믿는다.)

### 2. `append`는 쓰지 말고 `appendRow`를 쓴다
Sheets API의 `append`는 "표 인식" 휴리스틱이라, 시작 셀이 빈 값이면 엉뚱한 행(예: 1행)에 꽂힐 수 있다.
`appendRow <spreadsheetId> "SheetName" <keyColLetter> <valuesJSON>`는 keyColLetter 열을 실제로 스캔해서
마지막 데이터 행 다음에 정확히 쓴다. `--startCol=C`로 시작 열을 지정하면 A열부터가 아니라 특정 열부터
쓸 수 있다(수식열 건너뛸 때 필수).

### 3. ARRAYFORMULA 열(자동 순번 등)에 직접 값을 쓰면 깨진다
`update`/`append`/`appendRow`는 쓰기 전에 대상 범위가 걸친 열에 `ARRAYFORMULA`가 있는지 자동 검사한다.
걸리면 기본적으로 거부하고 한국어로 이유를 설명한다. 정말 그 열에 쓰고 싶을 때만 `--force`.
(실제 사고: 자동 순번 열에 직접 값을 써서 `#REF!` 에러 발생 — 그 열 아래 셀만 지워서 스필 복구했었음.)

### 4. 서식은 자동으로 안 맞는다 — `copyFormat`으로 명시적으로 맞춰야 함
`appendRow`는 기본적으로 바로 위 행(or `--formatFrom=N`)의 서식을 자동 복사하지만, `update`로 직접
여러 행을 쓸 때는 서식이 전혀 안 따라온다. 새로 쓴 행이 기존 행과 정렬/폰트/배경이 다르면
`copyFormat <spreadsheetId> "Sheet1!A4:E4" "Sheet1!A5:E29"` 처럼 서식만 복사한다 — 소스가 1행이고
대상이 여러 행이어도 자동으로 타일링되어 전체에 적용된다(실측 확인, 25행 이상에서도 정상 동작).

### 5. 배경색이 위/아래 행과 다르게 보여도 버그가 아닐 수 있다
시트에 `bandedRanges`(줄무늬 서식)가 걸려있으면 행 위치에 따라 배경색이 자동으로 바뀐다.
색이 이상해 보이면 먼저 `metadata --grid` 또는 `spreadsheets.get`으로 `bandedRanges`부터 확인할 것 —
색을 "고치려고" 서식을 덮어쓰면 오히려 밴딩을 깨뜨린다.

### 6. 라벨+값 페어 서식(표 명세서 템플릿 등)은 실제로 **병합 셀**이다
"데이터베이스명 | 값 | | | 테이블명 | 값" 같은 템플릿 행은 값 칸(B:D, F:I)이 병합되어 있다.
겉보기엔 여러 칸에 같은 파란 배경이 칠해진 것처럼 보이지만(병합 안 해도 서식만 복사하면 비슷해 보임),
실제로는 `merge` 명령으로 병합해야 한다. 새 템플릿 블록을 여러 개(예: 테이블 28개) 만들 때는
전체 병합 범위를 배열로 모아서 `merge <spreadsheetId> <rangesJSON>` 한 번에 처리 — 블록마다 개별
호출하면 API 호출이 블록 수만큼 늘어난다.

### 7. `autoFit`(COLUMNS)은 잘 되지만, `autoFit`(ROWS)은 WRAP된 셀의 줄바꿈 높이를 못 잡는다
**실측 확인된 API 한계**: 열 폭을 줄이고 `wrap`으로 WRAP을 걸어도, `autoFit --dimension=ROWS`는
한 줄 기준으로만 높이를 계산해서 260px 폭에 141자짜리 텍스트가 들어있어도 21px(한 줄) 그대로 둔다.
UI에서 행 경계를 더블클릭하는 것과 API의 `autoResizeDimensions(ROWS)`는 동작이 다르다.

**해결 순서 (긴 텍스트 열을 옆으로 안 넓히고 세로로만 늘리고 싶을 때):**
```
node sheets.js setWidth <id> "Sheet1!G:G" 260        # 1. 폭을 원하는 값으로 고정
node sheets.js wrap <id> "Sheet1!G1:G388" WRAP        # 2. 그 범위에 줄바꿈 켜기
node sheets.js fitWrap <id> "Sheet1!G1:G388" --colWidth=260   # 3. 글자수 기반으로 행 높이 직접 계산/적용
```
`fitWrap`은 `autoResizeDimensions`를 쓰지 않고, 셀 텍스트 길이 ÷ (열폭 기준 줄당 글자수)로 필요한 줄 수를
추정해서 그 행만 `updateDimensionProperties`로 높이를 지정한다. 기본 가정: `charWidth=6px`(Roboto 10pt
기준 어림값), `lineHeight=21px`, `padding=8px`. 폰트/크기가 다르면 `--charWidth=`로 보정.
한 줄로 충분한 행은 건드리지 않는다(불필요한 API 호출 방지).

## 인증/보안

- `config.json`, `token.json`은 git에 커밋 금지(.gitignore 처리됨), 민감정보 포함.
- SSH 등 다른 시스템의 비밀번호를 이 저장소나 스크립트에 하드코딩하지 않는다 — Sheets 인증과는
  무관한 별개 보안 원칙이지만 같은 세션에서 섞어 쓰지 않도록 주의.

## 이 파일을 갱신해야 할 때

새로운 Sheets API 함정을 발견하면(예: 특정 request 타입이 예상과 다르게 동작, 새로운 안전장치가
필요한 사고 등) 여기에 실측 근거와 함께 추가한다. 원인 불명확한 추측은 적지 말 것 — 실제로
`get`/`spreadsheets.get`으로 재현·확인한 것만 적는다.
