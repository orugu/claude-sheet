# claude-sheets-cli

Claude Code(터미널)에서 브라우저 없이 Google Sheets API v4를 직접 호출하는 로컬 CLI 도구.
`sheets.js` 하나로 값 읽기/쓰기, 서식·병합·열너비부터 시트 추가/삭제, 행/열 삽입, 정렬, 데이터
검증(드롭다운), 찾기/바꾸기, 필터, 피벗 테이블, 차트, 보호된 범위까지 다룬다. 명령/플래그 전체
목록은 `sheets.js` 맨 위 주석과 `README.md` 참고. 이 파일은 **실전에서 겪은 함정과 정해진 작업
순서**를 정리한 것.

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

### 8. 시트 추가/삭제/이름변경 뒤에는 `getSheetId` 캐시가 무효화되어야 함
`sheets.js` 내부에서 시트 이름→sheetId 조회 결과를 프로세스 내 캐시에 저장해두는데(`sheetIdCache`),
`addSheet`/`deleteSheet`/`renameSheet`/`duplicateSheet` 같은 구조 변경 명령 뒤에는 캐시를
`invalidateSheetIdCache()`로 지워야 한다. 안 지우면 같은 CLI 실행 안에서 방금 바뀐 시트 목록을
못 보고 이전 상태로 계산한다(지금은 각 함수에서 자동으로 처리하도록 반영해둠).

### 9. `sort`의 `dimensionIndex`는 시트 절대 열 인덱스, `pivot`의 `sourceColumnOffset`은 소스 범위 상대 오프셋 — 서로 다름
`sortRange`(Sheets API)의 `sortSpecs[].dimensionIndex`는 시트 전체 기준 0-based 열 인덱스를 그대로 쓴다
(범위 시작 기준 상대값 아님). 반면 피벗 테이블의 `rows`/`columns`/`values`의 `sourceColumnOffset`은
**소스 범위 시작 열 기준 상대 오프셋**이다(예: 소스가 `C1:F6`이면 D열은 offset 1). 헷갈리기 쉬워서
`sort` 명령은 열문자(`{"col":"D"}`)를 절대 인덱스로, `addPivot`은 같은 열문자를 소스 범위 기준
상대 오프셋으로 각각 자동 변환해준다 — 호출할 때는 그냥 열문자만 넘기면 됨.

### 10. `findReplace`는 `sheetId` 또는 `allSheets:true` 둘 중 하나만 있어야 함
시트 하나만 대상으로 하려면 `sheetId`를, 전체 시트 대상이면 `allSheets:true`를 준다 — 실제로
`findReplace` 명령의 두 번째 인자를 `"SheetName"` 또는 `"ALL"`로 받아서 분기 처리하도록 구현해둠.

### 11. `addNamedRange`/`protect`/`addBanding`은 생성된 ID를 응답에서 반드시 저장해둘 것
나중에 지우거나 참조하려면 각각 `namedRangeId`/`protectedRangeId`/`bandedRangeId`가 필요한데,
이건 생성 시점 응답에만 있고 이름만으로는 나중에 다시 못 찾는다(예: `metadata`로 전체를 뒤져야 함).
`deleteNamedRange`는 이름이 아니라 이 id를 받는다.

### 12. `addChart` 응답에서 anchor의 `rowIndex`가 안 보여도 정상(0이라 생략된 것)
Sheets API 응답은 protobuf 스타일이라 값이 0(기본값)이면 필드 자체를 JSON에서 생략한다.
`addChart`로 1행에 차트를 앵커했는데 응답에 `anchorCell.rowIndex`가 안 보이면 버그가 아니라
그냥 0행이라는 뜻 — `columnIndex`만 보이고 `rowIndex` 없으면 "0행"으로 해석할 것.

### 13. `addConditionalFormat`의 `formula` 모드는 상대참조 — 범위의 "첫 셀" 기준으로 자동 확장됨
`--formula='=$D2="완료"'`처럼 쓸 때, `$D2`는 컬럼만 고정(`$D`)하고 행은 상대참조(`2`)로 둬야
범위 안의 각 행마다 그 행의 D열을 보고 판단한다(엑셀/시트의 조건부 서식 수식 규칙과 동일 — 범위의
좌상단 셀 기준으로 수식이 각 셀에 상대적으로 적용됨). `$D$2`처럼 행까지 고정하면 전체 범위가
2행의 값 하나만 보고 통째로 켜지거나 꺼진다.

### 14. `parseRange`는 원래 "3:3" 같은 순수 행 범위(열 없음)를 못 읽었음 — 고침
`setRowHeight` 만들다가 실제로 걸린 버그. `A1:C5` 형태만 파싱하던 정규식이 "3:3"(전체 열에 걸친
행 범위)을 거부했다. `parseRange`에 순수 숫자(`^\d+(:\d+)?$`) 케이스를 먼저 검사하도록 추가하고,
`startCol`/`endCol`이 `null`이면 `rangeToGridRange`가 `startColumnIndex`/`endColumnIndex` 자체를
생략하도록(시트 전체 열을 의미) 고쳤다. `setWidth`가 쓰는 "G:G"(열만, 행 없음) 형태는 원래도 됐었음 —
이번에 고친 건 반대 방향(행만, 열 없음)이었다.

### 15. `updateSheetProperties.gridProperties`는 `showGrid`가 아니라 `hideGridlines`(반전된 값)
그리드라인 껐다 켜는 기능 만들다가 실측: `showGrid` 필드로 보내면 400 에러(`Cannot find field`)가
난다. 실제 필드명은 `hideGridlines`이고 의미가 반대다(끄고 싶으면 `true`). `setGridlines` 명령은
`on`/`off`를 받아서 내부적으로 `!show`로 뒤집어 보내도록 처리해둠 — 호출하는 쪽은 반전을 신경 안 써도 됨.

### 16. `autoFill`은 시드 셀이 1개면 그냥 복사, 2개 이상이어야 패턴(등차수열 등)을 인식함
실측 확인: `range`에 `N1`만 `1`로 채워두고 `N1:N6`에 autoFill을 걸면 전부 `1`로 복사만 된다.
`N1=1, N2=2`처럼 시드를 2개 이상 채워야 `3,4,5,6`으로 등차 확장된다 — 이건 버그가 아니라 실제
Sheets UI에서 채우기 핸들을 드래그할 때와 동일한 동작(셀 1개만 드래그하면 복사, 2개 이상 선택하고
드래그해야 패턴 인식). `autoFillRange`를 쓸 때는 항상 패턴을 판단할 수 있는 시드 셀을 최소 2개
이상 채워두고 나머지 빈 칸까지 포함한 전체 범위를 넘길 것.

## 검증된 명령 전체 목록 (2026-07-28 기준, 실제 스프레드시트에 대해 전부 실행/확인함)

값: `get`, `batchGet`, `update`, `append`, `appendRow`, `clear`
서식: `copyFormat`, `merge`, `autoFit`, `setWidth`, `wrap`, `fitWrap`, `numberFormat`, `note`, `addBanding`, `condFormat`
시트 관리: `addSheet`, `deleteSheet`, `renameSheet`, `duplicateSheet`, `freeze`, `copySheetTo`
행/열: `insertRows`, `insertCols`, `deleteRows`, `deleteCols`, `sort`
데이터 품질: `validate`, `clearValidation`, `findReplace`, `splitText`
구조/보호/분석: `addNamedRange`, `deleteNamedRange`, `protect`, `setFilter`, `clearFilter`, `addChart`, `addPivot`
삭제/이동: `deleteProtect`, `deleteBanding`, `deleteCondFormat`, `deleteChart`, `moveChart`
시각 속성: `setTabColor`, `hideSheet`, `showSheet`, `setGridlines`, `hideRows`, `hideCols`, `showRows`, `showCols`
그룹핑: `groupRows`, `groupCols`, `ungroupRows`, `ungroupCols`
필터 뷰: `addFilterView`, `deleteFilterView`, `duplicateFilterView`
개발자 메타데이터: `addMetadata`, `findMetadata`, `deleteMetadata`
데이터 정리: `trimWhitespace`, `deleteDuplicates`, `autoFillRange`
기타: `tabs`, `metadata`, `batchUpdate`

총 73개 명령. 전부 `1b5T1LgILmBZLl4QPuLb66ASqj6EG03fQVy39WMCBVlA`(실험실 스프레드시트)에서 실행 →
`get`/`metadata`로 반영 확인 → 테스트 흔적 정리(clear/delete) 순서로 검증했다.

모두 실험용 스프레드시트(`1b5T1LgILmBZLl4QPuLb66ASqj6EG03fQVy39WMCBVlA`)에서 실제로 실행하고
`get`/`metadata`로 결과까지 확인한 것들이다 — 문서만 보고 추측한 게 아님.

## 인증/보안

- `config.json`, `token.json`은 git에 커밋 금지(.gitignore 처리됨), 민감정보 포함.
- SSH 등 다른 시스템의 비밀번호를 이 저장소나 스크립트에 하드코딩하지 않는다 — Sheets 인증과는
  무관한 별개 보안 원칙이지만 같은 세션에서 섞어 쓰지 않도록 주의.

## 이 파일을 갱신해야 할 때

새로운 Sheets API 함정을 발견하면(예: 특정 request 타입이 예상과 다르게 동작, 새로운 안전장치가
필요한 사고 등) 여기에 실측 근거와 함께 추가한다. 원인 불명확한 추측은 적지 말 것 — 실제로
`get`/`spreadsheets.get`으로 재현·확인한 것만 적는다.
