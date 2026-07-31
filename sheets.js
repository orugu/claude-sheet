#!/usr/bin/env node
// Google Sheets 조작용 CLI. 값 읽기/쓰기부터 batchUpdate(서식/수식/차트/병합 등 고급 조작)까지.
//
// 사용법:
//   node sheets.js tabs <spreadsheetId>                                시트(탭) 이름/ID 목록
//   node sheets.js get <spreadsheetId> <range> [--formulas]            --formulas: 값 대신 수식 원문
//   node sheets.js batchGet <spreadsheetId> <rangesJSON>                예: '["Sheet1!A1:B2","Sheet1!D1:D5"]'
//   node sheets.js update <spreadsheetId> <range> <valuesJSON> [--force]
//   node sheets.js append <spreadsheetId> <range> <valuesJSON> [--force]   (표 인식 방식, 위치가 예상과 다를 수 있음 — appendRow 권장)
//   node sheets.js appendRow <spreadsheetId> "SheetName" <keyColLetter> <valuesJSON> [--formatFrom=행번호] [--startCol=열문자]
//                                                                        keyColLetter 기준으로 실제 마지막 데이터 행 다음에 정확히 씀.
//                                                                        --formatFrom=4 처럼 주면 그 행의 서식(정렬/폰트/배경 등)을 새 행에 복사.
//                                                                        기본값: 바로 위 행(직전 데이터 행)에서 자동 복사.
//                                                                        --startCol=C 처럼 주면 values를 A열이 아니라 C열부터 씀 (수식열 등을 건너뛸 때).
//   node sheets.js copyFormat <spreadsheetId> <sourceRange> <destRange>   서식만 복사(값 안 건드림). 같은/다른 시트 다 가능
//   node sheets.js merge <spreadsheetId> <rangesJSON> [--mergeType=MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS]
//                                                                        여러 범위를 한 번의 batchUpdate로 병합. 예: '["Sheet1!B1:D1","Sheet1!F1:I1"]'
//   node sheets.js autoFit <spreadsheetId> <range> [--dimension=BOTH|COLUMNS|ROWS]
//                                                                        열/행 경계 더블클릭한 것과 동일 — 셀 내용 길이에 맞춰 폭/높이 자동조정
//   node sheets.js setWidth <spreadsheetId> <range> <pixels>            열 폭을 고정값으로 지정 (autoFit이 너무 넓게 잡았을 때)
//   node sheets.js wrap <spreadsheetId> <range> [WRAP|CLIP|OVERFLOW_CELL]
//                                                                        긴 텍스트를 옆으로 안 넓히고 셀 안에서 줄바꿈(세로로 길어짐). 기본 WRAP.
//   node sheets.js fitWrap <spreadsheetId> <range> [--colWidth=260] [--charWidth=6] [--lineHeight=21]
//                                                                        ⚠️ autoFit(ROWS)는 WRAP된 셀의 줄바꿈 높이를 못 잡음(실측 확인).
//                                                                        글자수 기반으로 행 높이 직접 계산해서 적용. 순서: setWidth → wrap → fitWrap
//   node sheets.js clear <spreadsheetId> <range>
//   node sheets.js metadata <spreadsheetId> [--grid]
//   node sheets.js batchUpdate <spreadsheetId> <requestsJSON>          Sheets API batchUpdate requests 배열 그대로
//
// --- 시트(탭) 관리 ---
//   node sheets.js addSheet <spreadsheetId> "새 시트 이름" [--index=N]
//   node sheets.js deleteSheet <spreadsheetId> "시트이름"              ⚠️ 되돌릴 수 없음
//   node sheets.js renameSheet <spreadsheetId> "기존이름" "새이름"
//   node sheets.js duplicateSheet <spreadsheetId> "원본시트" [새이름] [--index=N]
//   node sheets.js freeze <spreadsheetId> "SheetName" [고정행수] [고정열수]   예: freeze id Sheet1 1 0
//
// --- 행/열 편집 ---
//   node sheets.js insertRows|insertCols <spreadsheetId> "SheetName" <기준번호> [개수=1] [--inheritBefore]
//   node sheets.js deleteRows|deleteCols <spreadsheetId> "SheetName" <시작번호> [끝번호]   ⚠️ 되돌릴 수 없음
//   node sheets.js sort <spreadsheetId> <range> <sortSpecJSON>          헤더 제외 범위. 예: '[{"col":"D","order":"DESC"}]'
//
// --- 데이터 품질/입력 보조 ---
//   node sheets.js validate <spreadsheetId> <range> <valuesJSON> [--noStrict] [--noDropdown]  드롭다운 목록(데이터 확인)
//   node sheets.js clearValidation <spreadsheetId> <range>
//   node sheets.js findReplace <spreadsheetId> <"SheetName"|ALL> <찾을값> <바꿀값> [--matchCase] [--entireCell] [--regex]
//   node sheets.js numberFormat <spreadsheetId> <range> <NUMBER|INTEGER|CURRENCY|PERCENT|DATE|TIME|DATE_TIME|SCIENTIFIC|커스텀패턴>
//   node sheets.js note <spreadsheetId> <range> <메모내용>                셀 메모(note). ""로 주면 삭제
//
// --- 구조/보호/시각화 ---
//   node sheets.js addNamedRange <spreadsheetId> <이름> <range>
//   node sheets.js deleteNamedRange <spreadsheetId> <namedRangeId>
//   node sheets.js protect <spreadsheetId> <range> [--description=] [--warningOnly] [--editors=a@x.com,b@x.com]
//   node sheets.js addBanding <spreadsheetId> <range> [--firstColor=#hex] [--secondColor=#hex] [--headerColor=#hex]
//   node sheets.js addChart <spreadsheetId> "SheetName" <COLUMN|LINE|BAR|AREA|SCATTER> <dataRange> [--title=] [--anchorCell=] [--width=] [--height=]
//   node sheets.js setFilter <spreadsheetId> <range>                    기본 필터(드롭다운 화살표) 설정, 헤더 행 포함 범위
//   node sheets.js clearFilter <spreadsheetId> "SheetName"
//   node sheets.js splitText <spreadsheetId> <range> <구분자>            텍스트 나누기 (쉼표/탭/세미콜론/공백/커스텀)
//   node sheets.js addPivot <spreadsheetId> <sourceRange> <대상시트> <앵커셀> <rowsJSON> <valuesJSON> [colsJSON]
//                                                                        예: addPivot id "Sheet1!A1:F6" Sheet1 H1 '[{"col":"C"}]' '[{"col":"D","fn":"COUNTA"}]'
//   node sheets.js condFormat <spreadsheetId> <range> <colorScale|formula> [--minColor=] [--midColor=] [--maxColor=] [--formula="=..."] [--backgroundColor=]
//   node sheets.js copySheetTo <spreadsheetId> "SheetName" <대상spreadsheetId>   다른(또는 같은) 스프레드시트로 시트 복사
//   node sheets.js renameSpreadsheet <spreadsheetId> "새 제목"           스프레드시트 파일 자체의 제목 변경
//   node sheets.js setBorder <spreadsheetId> <range> <all|outer|inner|top,bottom,...> [--color=#000000] [--style=SOLID|DASHED|DOTTED|DOUBLE|SOLID_MEDIUM|SOLID_THICK]
//   node sheets.js unmerge <spreadsheetId> <range>
//   node sheets.js setRowHeight <spreadsheetId> "Sheet1!3:3" <pixels>
//   node sheets.js moveRows|moveCols <spreadsheetId> "SheetName" <시작번호> <끝번호> <이동목적지(그 앞으로)>
//
// --- 삭제 계열 (id는 각 add* 명령의 응답 또는 metadata에서 확인) ---
//   node sheets.js deleteProtect <spreadsheetId> <protectedRangeId>
//   node sheets.js deleteBanding <spreadsheetId> <bandedRangeId>
//   node sheets.js deleteCondFormat <spreadsheetId> "SheetName" <index>       0-based
//   node sheets.js deleteChart <spreadsheetId> <chartId>
//   node sheets.js moveChart <spreadsheetId> <chartId> "SheetName" <앵커셀> [--width=] [--height=]
//
// --- 시트 시각 속성 ---
//   node sheets.js setTabColor <spreadsheetId> "SheetName" <#RRGGBB>
//   node sheets.js hideSheet|showSheet <spreadsheetId> "SheetName"
//   node sheets.js setGridlines <spreadsheetId> "SheetName" <on|off>
//   node sheets.js hideRows|hideCols|showRows|showCols <spreadsheetId> "SheetName" <시작번호> [끝번호]
//   node sheets.js groupRows|groupCols <spreadsheetId> "SheetName" <시작번호> <끝번호>       접기/펼치기 그룹
//   node sheets.js ungroupRows|ungroupCols <spreadsheetId> "SheetName" <시작번호> <끝번호>
//
// --- 필터 뷰 (기본 필터 setFilter와 별개로 여러 개 저장 가능) ---
//   node sheets.js addFilterView <spreadsheetId> <range(헤더포함)> <제목>
//   node sheets.js deleteFilterView <spreadsheetId> <filterId>
//   node sheets.js duplicateFilterView <spreadsheetId> <filterId>
//
// --- 개발자 메타데이터 (키-값을 시트/범위/문서 전체에 붙여서 나중에 검색 가능) ---
//   node sheets.js addMetadata <spreadsheetId> <key> <value> [--sheetTarget="SheetName"] [--range="3:5"] [--visibility=DOCUMENT|PROJECT]
//   node sheets.js findMetadata <spreadsheetId> <key>
//   node sheets.js deleteMetadata <spreadsheetId> <key>
//
// --- 데이터 정리 ---
//   node sheets.js trimWhitespace <spreadsheetId> <range>
//   node sheets.js deleteDuplicates <spreadsheetId> <range> [비교열JSON 예:'["B","C"]']
//   node sheets.js autoFillRange <spreadsheetId> <range>       패턴 있는 셀+빈 셀 전체 범위 (드래그 채우기와 동일)
//
// --- 기타 셀/시트 조작 ---
//   node sheets.js cutPaste <spreadsheetId> <소스range> <목적지시작셀>   ⚠️ 원본은 지워짐(잘라내기)
//   node sheets.js insertCells|deleteCells <spreadsheetId> <range> <ROWS|COLUMNS>   부분 범위만 밀어냄/당겨짐(행·열 통째 아님)
//   node sheets.js collapseGroup|expandGroup <spreadsheetId> "SheetName" <ROWS|COLUMNS> <시작번호> <끝번호>
//   node sheets.js randomize <spreadsheetId> <range>            행 순서 무작위로 섞기
//   node sheets.js setLocale <spreadsheetId> <locale> [timeZone]   예: setLocale id ko_KR Asia/Seoul
//
// 참고: 표 안 서식은 셀 자체 서식(userEnteredFormat)과 줄무늬(밴딩) 서식이 섞여 있을 수 있다.
// 밴딩이 걸린 열은 행 위치에 따라 배경색이 자동으로 바뀌니, 배경색이 위/아래 행과 달라 보여도
// 버그가 아닐 수 있다 — metadata로 bandedRanges 먼저 확인해볼 것.
//
// ⚠️ 안전장치: update/append/appendRow는 쓰기 전에 대상 범위가 걸쳐있는 "열"에
// ARRAYFORMULA(자동 계산 열, 예: 순번)가 있는지 검사한다. 걸리면 기본적으로 거부하고,
// 정말 그 열에 쓰고 싶으면 --force를 붙여야 한다. (겪은 사고: 자동 순번 열에 직접 값을
// 써서 배열수식이 깨지고 #REF! 에러가 난 적 있음 — 그거 방지용)

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

const CONFIG_PATH = path.join(__dirname, "config.json");
const TOKEN_PATH = path.join(__dirname, "token.json");

function fail(msg) {
  console.error("오류:", msg);
  process.exit(1);
}

// JSON.parse를 try/catch 없이 그대로 쓰던 명령들(batchGet/update/append/appendRow/sort/
// validate/addPivot/batchUpdate)이 깨진 JSON을 받으면 원본 JS 에러("Unexpected token...")를
// 그대로 노출했다(실측 확인) - merge 명령이 이미 쓰던 "파싱 실패 — 예: ..." 스타일로 통일.
function parseJsonArg(json, exampleHint) {
  try {
    return JSON.parse(json);
  } catch {
    fail(`JSON 파싱 실패${exampleHint ? ` — 예: ${exampleHint}` : ""} (받은 값: ${json})`);
  }
}

function getClient() {
  if (!fs.existsSync(CONFIG_PATH)) fail("config.json 없음 — config.example.json 참고해서 만들어주세요.");
  if (!fs.existsSync(TOKEN_PATH)) fail("token.json 없음 — 먼저 `node auth.js` 실행해서 인증하세요.");

  const { client_id, client_secret } = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
  const tokens = JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"));

  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret);
  oAuth2Client.setCredentials(tokens);

  oAuth2Client.on("tokens", (newTokens) => {
    const merged = { ...tokens, ...newTokens };
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(merged, null, 2), { mode: 0o600 });
  });

  return google.sheets({ version: "v4", auth: oAuth2Client });
}

function printResult(dataOrObj) {
  console.log(JSON.stringify(dataOrObj, null, 2));
}

// ---------- range / 열 문자 유틸 ----------

function colLetterToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1; // 0-based
}

function indexToColLetter(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// "SheetName!A5:C5" 또는 "'Sheet Name'!A1" 파싱
function parseRange(range) {
  const idx = range.lastIndexOf("!");
  if (idx === -1) fail(`범위에 시트 이름이 없음: ${range} (예: "Sheet1!A1:C5")`);
  let sheetName = range.slice(0, idx);
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1);
  const cellPart = range.slice(idx + 1);

  // 순수 행 범위 (열 없음, 예: "3:3", "5:10") — 시트 전체 열에 걸친 행 지정
  const rowOnly = cellPart.match(/^(\d+)(?::(\d+))?$/);
  if (rowOnly) {
    const [, rowA, rowB] = rowOnly;
    return {
      sheetName,
      startCol: null,
      startRow: parseInt(rowA, 10),
      endCol: null,
      endRow: rowB ? parseInt(rowB, 10) : parseInt(rowA, 10),
    };
  }

  const m = cellPart.match(/^([A-Za-z]+)(\d+)?(?::([A-Za-z]+)(\d+)?)?$/);
  if (!m) fail(`범위 형식을 이해 못함: ${cellPart}`);
  const [, colA, rowA, colB, rowB] = m;
  return {
    sheetName,
    startCol: colLetterToIndex(colA),
    startRow: rowA ? parseInt(rowA, 10) : null,
    endCol: colB ? colLetterToIndex(colB) : colLetterToIndex(colA),
    endRow: rowB ? parseInt(rowB, 10) : rowA ? parseInt(rowA, 10) : null,
  };
}

// CLI에서 사용자가 직접 "rows"/"cols" 등으로 입력해도 API가 요구하는 정확한 대문자 enum으로 정규화
function normalizeDimension(s) {
  const up = (s || "").toUpperCase();
  if (up === "ROWS" || up === "ROW") return "ROWS";
  if (up === "COLUMNS" || up === "COL" || up === "COLS" || up === "COLUMN") return "COLUMNS";
  fail(`ROWS 또는 COLUMNS만 가능함 (받은 값: "${s}")`);
}

function quoteSheetName(name) {
  return /[^A-Za-z0-9_]/.test(name) ? `'${name}'` : name;
}

// ---------- 시트 이름 -> sheetId 조회 (batchUpdate용 GridRange 만들 때 필요) ----------

const sheetIdCache = new Map(); // spreadsheetId -> Map(title -> sheetId)

async function getSheetId(sheets, spreadsheetId, sheetName) {
  if (!sheetIdCache.has(spreadsheetId)) {
    const res = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets(properties(sheetId,title))",
    });
    const m = new Map(res.data.sheets.map((s) => [s.properties.title, s.properties.sheetId]));
    sheetIdCache.set(spreadsheetId, m);
  }
  const m = sheetIdCache.get(spreadsheetId);
  if (!m.has(sheetName)) fail(`시트 이름을 못 찾음: "${sheetName}" (실제 탭 이름과 정확히 일치해야 함, tabs 명령으로 확인)`);
  return m.get(sheetName);
}

function rangeToGridRange(sheetId, r) {
  return {
    sheetId,
    startRowIndex: r.startRow ? r.startRow - 1 : undefined,
    endRowIndex: r.endRow ? r.endRow : undefined,
    // startCol/endCol이 null이면(순수 행 범위, 예: "3:3") 열 인덱스 자체를 생략 — 시트 전체 열을 의미
    startColumnIndex: r.startCol === null || r.startCol === undefined ? undefined : r.startCol,
    endColumnIndex: r.endCol === null || r.endCol === undefined ? undefined : r.endCol + 1,
  };
}

async function copyFormat(sheets, spreadsheetId, sourceRangeStr, destRangeStr) {
  const src = parseRange(sourceRangeStr);
  const dst = parseRange(destRangeStr);
  const srcSheetId = await getSheetId(sheets, spreadsheetId, src.sheetName);
  const dstSheetId = await getSheetId(sheets, spreadsheetId, dst.sheetName);
  if (!src.startRow || !src.endRow || !dst.startRow || !dst.endRow)
    fail("copyFormat은 행 번호가 명시된 범위가 필요함 (예: Sheet1!A4:C4)");

  const req = {
    copyPaste: {
      source: rangeToGridRange(srcSheetId, src),
      destination: rangeToGridRange(dstSheetId, dst),
      pasteType: "PASTE_FORMAT",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// 열 폭을 고정 픽셀값으로 지정 (autoFit이 너무 넓게 잡았을 때 등)
async function setColumnWidth(sheets, spreadsheetId, rangeStr, pixelSize) {
  const r = parseRange(rangeStr);
  if (r.startCol === null) fail('setWidth는 열 범위가 필요함 (예: Sheet1!G:G) — 순수 행 범위("3:3")는 안 됨');
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    updateDimensionProperties: {
      range: { sheetId, dimension: "COLUMNS", startIndex: r.startCol, endIndex: r.endCol + 1 },
      properties: { pixelSize },
      fields: "pixelSize",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// 긴 텍스트를 옆으로 넓히지 않고 셀 안에서 줄바꿈(세로로 길어지게) — WRAP/CLIP/OVERFLOW_CELL
async function setWrap(sheets, spreadsheetId, rangeStr, strategy = "WRAP") {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("wrap은 행 번호가 명시된 범위가 필요함 (예: Sheet1!G1:G388)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    repeatCell: {
      range: rangeToGridRange(sheetId, r),
      cell: { userEnteredFormat: { wrapStrategy: strategy } },
      fields: "userEnteredFormat.wrapStrategy",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ⚠️ autoFit(ROWS)는 WRAP된 셀의 실제 줄바꿈 높이를 계산 못 함(실측: 260px 폭에 141자 텍스트도 21px 그대로).
// UI에서 행 경계 더블클릭하는 것과 달리 API의 autoResizeDimensions(ROWS)는 한 줄 기준으로만 계산하는 걸로 보임.
// 그래서 WRAP 걸린 긴 텍스트 열은 이 함수로 글자수 기반 추정 높이를 직접 지정해야 함.
async function fitWrapRowHeights(sheets, spreadsheetId, rangeStr, opts = {}) {
  const { colWidth = 260, charWidth = 6, lineHeight = 21, padding = 8 } = opts;
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("fitWrap은 행 번호가 명시된 범위가 필요함 (예: Sheet1!G1:G388)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const charsPerLine = Math.max(1, Math.floor((colWidth - padding) / charWidth));

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: rangeStr,
  });
  const values = res.data.values || [];

  const requests = [];
  values.forEach((row, i) => {
    const text = row[0] || "";
    if (!text) return;
    const lines = Math.max(1, Math.ceil(text.length / charsPerLine));
    if (lines <= 1) return; // 기본 한 줄 높이면 건드릴 필요 없음
    const rowIndex0 = r.startRow - 1 + i;
    requests.push({
      updateDimensionProperties: {
        range: { sheetId, dimension: "ROWS", startIndex: rowIndex0, endIndex: rowIndex0 + 1 },
        properties: { pixelSize: lines * lineHeight },
        fields: "pixelSize",
      },
    });
  });

  if (!requests.length) return { adjustedRows: 0 };
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  return { adjustedRows: requests.length, charsPerLine };
}

async function autoFit(sheets, spreadsheetId, rangeStr, dimension = "BOTH") {
  const r = parseRange(rangeStr);
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const requests = [];
  const wantCols = dimension === "BOTH" || dimension === "COLUMNS";
  const wantRows = dimension === "BOTH" || dimension === "ROWS";

  if (wantCols) {
    requests.push({
      autoResizeDimensions: {
        dimensions: {
          sheetId,
          dimension: "COLUMNS",
          startIndex: r.startCol,
          endIndex: r.endCol + 1,
        },
      },
    });
  }
  if (wantRows) {
    if (dimension === "ROWS" && !r.startRow)
      fail('행 자동맞춤은 행 번호가 있는 범위가 필요함 (예: Sheet1!A1:I388) — 열만 원하면 --dimension=COLUMNS');
    if (r.startRow && r.endRow) {
      requests.push({
        autoResizeDimensions: {
          dimensions: {
            sheetId,
            dimension: "ROWS",
            startIndex: r.startRow - 1,
            endIndex: r.endRow,
          },
        },
      });
    }
  }
  if (!requests.length) fail("autoFit: 적용할 범위가 없음 (열/행 범위를 확인)");
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  return res.data;
}

async function mergeCells(sheets, spreadsheetId, rangeStrs, mergeType = "MERGE_ALL") {
  const requests = [];
  for (const rangeStr of rangeStrs) {
    const r = parseRange(rangeStr);
    if (!r.startRow || !r.endRow) fail(`merge는 행 번호가 명시된 범위가 필요함 (예: Sheet1!B4:D4) — 받은 값: "${rangeStr}"`);
    const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
    requests.push({
      mergeCells: {
        range: rangeToGridRange(sheetId, r),
        mergeType,
      },
    });
  }
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests } });
  return res.data;
}

// ---------- 시트 관리: 추가/삭제/이름변경/복제 ----------

function invalidateSheetIdCache(spreadsheetId) {
  sheetIdCache.delete(spreadsheetId);
}

async function addSheet(sheets, spreadsheetId, title, index) {
  const props = { title };
  if (index !== undefined) props.index = index;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ addSheet: { properties: props } }] },
  });
  invalidateSheetIdCache(spreadsheetId);
  return res.data.replies[0].addSheet.properties;
}

async function deleteSheet(sheets, spreadsheetId, sheetName) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [{ deleteSheet: { sheetId } }] },
  });
  invalidateSheetIdCache(spreadsheetId);
  return res.data;
}

async function renameSheet(sheets, spreadsheetId, oldName, newName) {
  const sheetId = await getSheetId(sheets, spreadsheetId, oldName);
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          updateSheetProperties: {
            properties: { sheetId, title: newName },
            fields: "title",
          },
        },
      ],
    },
  });
  invalidateSheetIdCache(spreadsheetId);
  return res.data;
}

async function duplicateSheet(sheets, spreadsheetId, sourceName, newName, index) {
  const sourceSheetId = await getSheetId(sheets, spreadsheetId, sourceName);
  const req = { duplicateSheet: { sourceSheetId } };
  if (newName) req.duplicateSheet.newSheetName = newName;
  if (index !== undefined) req.duplicateSheet.insertSheetIndex = index;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests: [req] },
  });
  invalidateSheetIdCache(spreadsheetId);
  return res.data.replies[0].duplicateSheet.properties;
}

async function setFrozen(sheets, spreadsheetId, sheetName, frozenRows, frozenCols) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const props = { sheetId, gridProperties: {} };
  const fields = [];
  if (frozenRows !== undefined) {
    props.gridProperties.frozenRowCount = frozenRows;
    fields.push("gridProperties.frozenRowCount");
  }
  if (frozenCols !== undefined) {
    props.gridProperties.frozenColumnCount = frozenCols;
    fields.push("gridProperties.frozenColumnCount");
  }
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ updateSheetProperties: { properties: props, fields: fields.join(",") } }],
    },
  });
  return res.data;
}

// ---------- 행/열 삽입·삭제 ----------

async function insertRowsOrCols(sheets, spreadsheetId, sheetName, dimension, beforeIndex1based, count, inheritBefore) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const startIndex = beforeIndex1based - 1;
  const req = {
    insertDimension: {
      range: { sheetId, dimension, startIndex, endIndex: startIndex + count },
      inheritFromBefore: !!inheritBefore,
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function deleteRowsOrCols(sheets, spreadsheetId, sheetName, dimension, start1based, end1based) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    deleteDimension: {
      range: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 정렬 ----------

// sortSpecs: [{col:"D", order:"ASC"|"DESC"}, ...]
async function sortRange(sheets, spreadsheetId, rangeStr, sortSpecs) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("sort는 행 번호가 명시된 범위가 필요함 (예: Sheet1!A2:F6, 헤더 제외)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    sortRange: {
      range: rangeToGridRange(sheetId, r),
      sortSpecs: sortSpecs.map((s) => ({
        dimensionIndex: colLetterToIndex(s.col),
        sortOrder: (s.order || "ASC").toUpperCase() === "DESC" ? "DESCENDING" : "ASCENDING",
      })),
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 데이터 검증 (드롭다운 등) ----------

async function setValidation(sheets, spreadsheetId, rangeStr, listValues, opts = {}) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("validate는 행 번호가 명시된 범위가 필요함 (예: Sheet1!D2:D100)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    setDataValidation: {
      range: rangeToGridRange(sheetId, r),
      rule: {
        condition: {
          type: "ONE_OF_LIST",
          values: listValues.map((v) => ({ userEnteredValue: v })),
        },
        strict: opts.strict !== false,
        showCustomUi: opts.showDropdown !== false,
      },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function clearValidation(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { setDataValidation: { range: rangeToGridRange(sheetId, r) } }; // rule 없으면 해제
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 이름 지정 범위 ----------

async function addNamedRange(sheets, spreadsheetId, name, rangeStr) {
  const r = parseRange(rangeStr);
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { addNamedRange: { namedRange: { name, range: rangeToGridRange(sheetId, r) } } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].addNamedRange.namedRange;
}

async function deleteNamedRange(sheets, spreadsheetId, namedRangeId) {
  const req = { deleteNamedRange: { namedRangeId } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 찾기/바꾸기 ----------

async function findReplace(sheets, spreadsheetId, sheetName, find, replacement, opts = {}) {
  const req = {
    findReplace: {
      find,
      replacement,
      matchCase: !!opts.matchCase,
      matchEntireCell: !!opts.entireCell,
      searchByRegex: !!opts.regex,
    },
  };
  if (sheetName && sheetName !== "ALL") {
    req.findReplace.sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  } else {
    req.findReplace.allSheets = true;
  }
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].findReplace;
}

// ---------- 표시 형식(숫자/통화/퍼센트/날짜 등) ----------

const NUMBER_FORMAT_PRESETS = {
  NUMBER: "#,##0.##",
  INTEGER: "#,##0",
  CURRENCY: "₩#,##0",
  PERCENT: "0.00%",
  DATE: "yyyy-mm-dd",
  TIME: "hh:mm:ss",
  DATE_TIME: "yyyy-mm-dd hh:mm:ss",
  SCIENTIFIC: "0.00E+00",
};

// 프리셋 키 -> 실제 Sheets API NumberFormatType. "INTEGER"는 API에 없는 타입이라 NUMBER로 매핑.
const NUMBER_FORMAT_API_TYPES = {
  NUMBER: "NUMBER",
  INTEGER: "NUMBER",
  CURRENCY: "CURRENCY",
  PERCENT: "PERCENT",
  DATE: "DATE",
  TIME: "TIME",
  DATE_TIME: "DATE_TIME",
  SCIENTIFIC: "SCIENTIFIC",
};

async function setNumberFormat(sheets, spreadsheetId, rangeStr, typeOrPattern) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("numberFormat은 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const pattern = NUMBER_FORMAT_PRESETS[typeOrPattern] || typeOrPattern;
  const apiType = NUMBER_FORMAT_API_TYPES[typeOrPattern] || "NUMBER"; // 프리셋이 아니라 커스텀 패턴이면 NUMBER 기준
  const req = {
    repeatCell: {
      range: rangeToGridRange(sheetId, r),
      cell: { userEnteredFormat: { numberFormat: { type: apiType, pattern } } },
      fields: "userEnteredFormat.numberFormat",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 셀 메모(note) ----------

async function setNote(sheets, spreadsheetId, rangeStr, note) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("note는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    repeatCell: {
      range: rangeToGridRange(sheetId, r),
      cell: { note },
      fields: "note",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 보호된 범위 ----------

async function protectRange(sheets, spreadsheetId, rangeStr, opts = {}) {
  const r = parseRange(rangeStr);
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const protectedRange = { range: rangeToGridRange(sheetId, r) };
  if (opts.description) protectedRange.description = opts.description;
  if (opts.warningOnly) protectedRange.warningOnly = true;
  if (opts.editors && opts.editors.length) protectedRange.editors = { users: opts.editors };
  const req = { addProtectedRange: { protectedRange } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].addProtectedRange.protectedRange;
}

// ---------- 줄무늬(밴딩) 서식 ----------

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return {
    red: parseInt(h.slice(0, 2), 16) / 255,
    green: parseInt(h.slice(2, 4), 16) / 255,
    blue: parseInt(h.slice(4, 6), 16) / 255,
  };
}

async function addBanding(sheets, spreadsheetId, rangeStr, opts = {}) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("addBanding은 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const bandedRange = {
    range: rangeToGridRange(sheetId, r),
    rowProperties: {
      firstBandColorStyle: { rgbColor: hexToRgb(opts.firstColor || "#FFFFFF") },
      secondBandColorStyle: { rgbColor: hexToRgb(opts.secondColor || "#F3F3F3") },
    },
  };
  if (opts.headerColor) {
    bandedRange.rowProperties.headerColorStyle = { rgbColor: hexToRgb(opts.headerColor) };
  }
  const req = { addBanding: { bandedRange } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].addBanding.bandedRange;
}

// ---------- 차트 ----------

async function addChart(sheets, spreadsheetId, sheetName, chartType, dataRangeStr, opts = {}) {
  const dr = parseRange(dataRangeStr);
  const sheetId = await getSheetId(sheets, spreadsheetId, dr.sheetName);
  const domain = {
    domain: {
      sourceRange: {
        sources: [
          {
            sheetId,
            startRowIndex: dr.startRow - 1,
            endRowIndex: dr.endRow,
            startColumnIndex: dr.startCol,
            endColumnIndex: dr.startCol + 1,
          },
        ],
      },
    },
  };
  // ⚠️ 실측 확인된 API 제약: series 하나의 sourceRange는 반드시 열 폭이 1이어야 함
  // ("ChartSourceRange ranges require all rows or all columns to have length of 1" 400 에러).
  // 데이터가 3열 이상(도메인 1열 + 값 2열 이상)이면 값 열마다 series를 하나씩 따로 만들어야 함 —
  // 전에는 값 열 전체를 series 하나로 뭉쳐서 넘겨서 3열 이상일 때 항상 에러가 났었음.
  const series = [];
  for (let col = dr.startCol + 1; col <= dr.endCol; col++) {
    series.push({
      series: {
        sourceRange: {
          sources: [
            {
              sheetId,
              startRowIndex: dr.startRow - 1,
              endRowIndex: dr.endRow,
              startColumnIndex: col,
              endColumnIndex: col + 1,
            },
          ],
        },
      },
      targetAxis: "LEFT_AXIS",
    });
  }
  const anchorCell = opts.anchorCell ? parseRange(`${sheetName}!${opts.anchorCell}`) : null;

  const req = {
    addChart: {
      chart: {
        spec: {
          title: opts.title || "",
          basicChart: {
            chartType, // COLUMN, LINE, BAR, AREA, SCATTER
            legendPosition: "BOTTOM_LEGEND",
            axis: [{ position: "BOTTOM_AXIS" }, { position: "LEFT_AXIS" }],
            domains: [domain],
            series,
            headerCount: 1,
          },
        },
        position: {
          overlayPosition: {
            anchorCell: {
              sheetId, // 데이터가 있는 시트에 앵커
              rowIndex: anchorCell ? anchorCell.startRow - 1 : dr.endRow,
              columnIndex: anchorCell ? anchorCell.startCol : dr.endCol + 2,
            },
            widthPixels: opts.width || 600,
            heightPixels: opts.height || 350,
          },
        },
      },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].addChart.chart;
}

// ---------- 기본 필터 ----------

async function setBasicFilter(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("filter는 행 번호가 명시된 범위가 필요함 (헤더 행 포함, 예: Sheet1!A1:F6)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { setBasicFilter: { filter: { range: rangeToGridRange(sheetId, r) } } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function clearBasicFilter(sheets, spreadsheetId, sheetName) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = { clearBasicFilter: { sheetId } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 텍스트 나누기 (구분자 기준 열 분리) ----------

const DELIMITER_TYPES = { ",": "COMMA", "\t": "TAB", ";": "SEMICOLON", " ": "SPACE" };

async function splitTextToColumns(sheets, spreadsheetId, rangeStr, delimiter) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("splitText는 행 번호가 명시된 범위가 필요함 (예: Sheet1!A2:A10)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const delimiterType = DELIMITER_TYPES[delimiter];
  const req = {
    textToColumns: {
      source: rangeToGridRange(sheetId, r),
      delimiterType: delimiterType || "CUSTOM",
      ...(delimiterType ? {} : { delimiter }),
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 피벗 테이블 ----------

// rowsSpec/valuesSpec: [{col:"B", label?:"..."}] / valuesSpec: [{col:"D", fn:"SUM"|"COUNTA"|"AVERAGE"|...}]
async function addPivotTable(sheets, spreadsheetId, sourceRangeStr, anchorSheetName, anchorCellStr, rowsSpec, valuesSpec, colsSpec = []) {
  const src = parseRange(sourceRangeStr);
  if (!src.startRow || !src.endRow) fail("pivot 소스는 행 번호가 명시된 범위가 필요함 (헤더 포함)");
  const srcSheetId = await getSheetId(sheets, spreadsheetId, src.sheetName);
  const anchorSheetId = await getSheetId(sheets, spreadsheetId, anchorSheetName);
  const anchor = parseRange(`${anchorSheetName}!${anchorCellStr}`);

  const pivotTable = {
    source: rangeToGridRange(srcSheetId, src),
    rows: rowsSpec.map((s) => ({
      sourceColumnOffset: colLetterToIndex(s.col) - src.startCol,
      showTotals: true,
      sortOrder: "ASCENDING",
    })),
    columns: colsSpec.map((s) => ({
      sourceColumnOffset: colLetterToIndex(s.col) - src.startCol,
      showTotals: true,
      sortOrder: "ASCENDING",
    })),
    values: valuesSpec.map((s) => ({
      sourceColumnOffset: colLetterToIndex(s.col) - src.startCol,
      summarizeFunction: s.fn || "SUM",
    })),
  };

  const req = {
    updateCells: {
      rows: [{ values: [{ pivotTable }] }],
      start: { sheetId: anchorSheetId, rowIndex: anchor.startRow - 1, columnIndex: anchor.startCol },
      fields: "pivotTable",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 조건부 서식 ----------

// mode "colorScale": {minColor,midColor,maxColor} (hex, mid 생략 가능)
// mode "formula": {formula, backgroundColor} — formula는 "=" 포함, 범위 첫 셀 기준 상대참조
async function addConditionalFormat(sheets, spreadsheetId, rangeStr, mode, opts = {}) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("conditionalFormat은 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const gridRange = rangeToGridRange(sheetId, r);

  let rule;
  if (mode === "colorScale") {
    const points = [{ color: hexToRgb(opts.minColor || "#F4C7C3"), type: "MIN" }];
    if (opts.midColor) points.push({ color: hexToRgb(opts.midColor), type: "PERCENTILE", value: "50" });
    points.push({ color: hexToRgb(opts.maxColor || "#B7E1CD"), type: "MAX" });
    rule = {
      ranges: [gridRange],
      gradientRule: {
        minpoint: { colorStyle: { rgbColor: points[0].color }, type: points[0].type },
        ...(opts.midColor
          ? { midpoint: { colorStyle: { rgbColor: points[1].color }, type: points[1].type, value: points[1].value } }
          : {}),
        maxpoint: {
          colorStyle: { rgbColor: points[points.length - 1].color },
          type: points[points.length - 1].type,
        },
      },
    };
  } else if (mode === "formula") {
    if (!opts.formula) fail('formula 모드는 --formula="=..." 필요');
    rule = {
      ranges: [gridRange],
      booleanRule: {
        condition: { type: "CUSTOM_FORMULA", values: [{ userEnteredValue: opts.formula }] },
        format: { backgroundColorStyle: { rgbColor: hexToRgb(opts.backgroundColor || "#FFF2CC") } },
      },
    };
  } else {
    fail('mode는 "colorScale" 또는 "formula"만 지원');
  }

  const req = { addConditionalFormatRule: { rule, index: 0 } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 시트를 다른(또는 같은) 스프레드시트로 복사 ----------

async function copySheetTo(sheets, spreadsheetId, sheetName, destSpreadsheetId) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const res = await sheets.spreadsheets.sheets.copyTo({
    spreadsheetId,
    sheetId,
    requestBody: { destinationSpreadsheetId: destSpreadsheetId },
  });
  // 대상(destSpreadsheetId) 쪽에 새 시트가 생기므로 그 스프레드시트의 캐시를 무효화해야 함
  // (원본 spreadsheetId는 시트 목록이 안 바뀌니 그대로 둔다).
  invalidateSheetIdCache(destSpreadsheetId);
  return res.data;
}

// ---------- 스프레드시트 제목 변경 ----------

async function renameSpreadsheet(sheets, spreadsheetId, newTitle) {
  const req = { updateSpreadsheetProperties: { properties: { title: newTitle }, fields: "title" } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 테두리 ----------

// sides: "all" | "outer" | "inner" | "top,bottom,left,right,innerHorizontal,innerVertical" 콤마 조합 (풀네임만 인식, 축약형 없음)
async function setBorder(sheets, spreadsheetId, rangeStr, sides, opts = {}) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("border는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const border = {
    style: opts.style || "SOLID",
    colorStyle: { rgbColor: hexToRgb(opts.color || "#000000") },
  };
  const wanted = new Set(
    sides === "all"
      ? ["top", "bottom", "left", "right", "innerHorizontal", "innerVertical"]
      : sides === "outer"
      ? ["top", "bottom", "left", "right"]
      : sides === "inner"
      ? ["innerHorizontal", "innerVertical"]
      : sides.split(",").map((s) => s.trim())
  );
  const req = { updateBorders: { range: rangeToGridRange(sheetId, r) } };
  for (const side of ["top", "bottom", "left", "right", "innerHorizontal", "innerVertical"]) {
    if (wanted.has(side)) req.updateBorders[side] = border;
  }
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 병합 해제 / 행높이 / 행·열 이동 ----------

async function unmergeCells(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("unmerge는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { unmergeCells: { range: rangeToGridRange(sheetId, r) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function setRowHeight(sheets, spreadsheetId, rangeStr, pixelSize) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("setRowHeight는 행 번호가 명시된 범위가 필요함 (예: Sheet1!3:3)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    updateDimensionProperties: {
      range: { sheetId, dimension: "ROWS", startIndex: r.startRow - 1, endIndex: r.endRow },
      properties: { pixelSize },
      fields: "pixelSize",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// dimension: ROWS|COLUMNS. 1-based 번호로 [start,end] 구간을 destinationIndex 앞으로 옮김
async function moveRowsOrCols(sheets, spreadsheetId, sheetName, dimension, start1based, end1based, destBefore1based) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    moveDimension: {
      source: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
      destinationIndex: destBefore1based - 1,
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 삭제 계열 (보호범위/밴딩/조건부서식/차트) ----------

async function deleteProtectedRange(sheets, spreadsheetId, protectedRangeId) {
  const req = { deleteProtectedRange: { protectedRangeId: parseInt(protectedRangeId, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function deleteBandingRange(sheets, spreadsheetId, bandedRangeId) {
  const req = { deleteBanding: { bandedRangeId: parseInt(bandedRangeId, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function deleteConditionalFormat(sheets, spreadsheetId, sheetName, index) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = { deleteConditionalFormatRule: { sheetId, index: parseInt(index, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function deleteChart(sheets, spreadsheetId, chartId) {
  const req = { deleteEmbeddedObject: { objectId: parseInt(chartId, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function moveChart(sheets, spreadsheetId, chartId, sheetName, anchorCellStr, width, height) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const anchor = parseRange(`${sheetName}!${anchorCellStr}`);
  const req = {
    updateEmbeddedObjectPosition: {
      objectId: parseInt(chartId, 10),
      newPosition: {
        overlayPosition: {
          anchorCell: { sheetId, rowIndex: anchor.startRow - 1, columnIndex: anchor.startCol },
          ...(width ? { widthPixels: width } : {}),
          ...(height ? { heightPixels: height } : {}),
        },
      },
      fields: "*",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 시트 시각 속성: 탭 색상 / 숨기기 / 그리드라인 ----------

async function setTabColor(sheets, spreadsheetId, sheetName, hex) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    updateSheetProperties: {
      properties: { sheetId, tabColorStyle: { rgbColor: hexToRgb(hex) } },
      fields: "tabColorStyle",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function setSheetHidden(sheets, spreadsheetId, sheetName, hidden) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = { updateSheetProperties: { properties: { sheetId, hidden }, fields: "hidden" } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ⚠️ API 필드는 "showGrid"가 아니라 "hideGridlines"(반전된 boolean) — 실측으로 확인함(showGrid는 400 에러)
async function setGridlines(sheets, spreadsheetId, sheetName, show) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    updateSheetProperties: {
      properties: { sheetId, gridProperties: { hideGridlines: !show } },
      fields: "gridProperties.hideGridlines",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// dimension: ROWS|COLUMNS, 1-based 구간을 숨기거나 다시 보이게
async function setDimensionHidden(sheets, spreadsheetId, sheetName, dimension, start1based, end1based, hidden) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    updateDimensionProperties: {
      range: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
      properties: { hiddenByUser: hidden },
      fields: "hiddenByUser",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 행/열 그룹핑(접기/펼치기, 아웃라인) ----------

async function addDimensionGroup(sheets, spreadsheetId, sheetName, dimension, start1based, end1based) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    addDimensionGroup: {
      range: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function deleteDimensionGroup(sheets, spreadsheetId, sheetName, dimension, start1based, end1based) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    deleteDimensionGroup: {
      range: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 필터 뷰 (여러 개 저장 가능, 기본 필터와 별개) ----------

async function addFilterView(sheets, spreadsheetId, rangeStr, title) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("filterView는 행 번호가 명시된 범위가 필요함 (헤더 포함)");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    addFilterView: {
      filter: { title, range: rangeToGridRange(sheetId, r) },
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].addFilterView.filter;
}

async function deleteFilterView(sheets, spreadsheetId, filterId) {
  const req = { deleteFilterView: { filterId: parseInt(filterId, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

async function duplicateFilterView(sheets, spreadsheetId, filterId) {
  const req = { duplicateFilterView: { filterId: parseInt(filterId, 10) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].duplicateFilterView.filter;
}

// ---------- 개발자 메타데이터 (키-값을 시트/범위/스프레드시트에 붙여서 나중에 검색) ----------

async function addDeveloperMetadata(sheets, spreadsheetId, key, value, opts = {}) {
  const metadata = {
    metadataKey: key,
    metadataValue: value,
    visibility: opts.visibility || "DOCUMENT",
  };
  if (opts.sheetName) {
    const sheetId = await getSheetId(sheets, spreadsheetId, opts.sheetName);
    if (opts.range) {
      // 행 범위 기준으로만 지원 (예: "3:5") — 열 범위/셀 범위가 필요하면 dimension을 COLUMNS로 바꿔서 확장
      const r = parseRange(`${opts.sheetName}!${opts.range}`);
      if (!r.startRow) fail('developer metadata의 range는 행 범위여야 함 (예: --range="3:5")');
      metadata.location = {
        dimensionRange: { sheetId, dimension: "ROWS", startIndex: r.startRow - 1, endIndex: r.endRow },
      };
    } else {
      metadata.location = { sheetId };
    }
  } else {
    metadata.location = { spreadsheet: true };
  }
  const req = { createDeveloperMetadata: { developerMetadata: metadata } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].createDeveloperMetadata.developerMetadata;
}

async function searchDeveloperMetadata(sheets, spreadsheetId, key) {
  const res = await sheets.spreadsheets.developerMetadata.search({
    spreadsheetId,
    requestBody: { dataFilters: [{ developerMetadataLookup: { metadataKey: key } }] },
  });
  return res.data.matchedDeveloperMetadata || [];
}

async function deleteDeveloperMetadata(sheets, spreadsheetId, key) {
  const req = { deleteDeveloperMetadata: { dataFilter: { developerMetadataLookup: { metadataKey: key } } } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 데이터 정리: 공백 제거 / 중복 제거 / 자동 채우기 ----------

async function trimWhitespace(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("trimWhitespace는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { trimWhitespace: { range: rangeToGridRange(sheetId, r) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].trimWhitespace;
}

// compareCols: 중복 판단 기준 열문자 배열(생략하면 전체 열 기준)
async function deleteDuplicateRows(sheets, spreadsheetId, rangeStr, compareCols = []) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("deleteDuplicates는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = {
    deleteDuplicates: {
      range: rangeToGridRange(sheetId, r),
      ...(compareCols.length
        ? { comparisonColumns: compareCols.map((c) => ({ sheetId, dimension: "COLUMNS", startIndex: colLetterToIndex(c), endIndex: colLetterToIndex(c) + 1 })) }
        : {}),
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data.replies[0].deleteDuplicates;
}

// 시트 UI에서 채우기 핸들을 드래그하는 것과 동일 — range 안에 이미 있는 패턴(수식, 순번, 요일 등)을
// 감지해서 range의 나머지 빈 칸까지 자동 확장한다. range = 패턴이 있는 셀 + 채울 빈 셀 전체.
async function autoFillRange(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("autoFillRange는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { autoFill: { range: rangeToGridRange(sheetId, r), useAlternateSeries: false } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 잘라내기(이동, 원본은 지워짐) ----------

async function cutPaste(sheets, spreadsheetId, sourceRangeStr, destCellStr) {
  const src = parseRange(sourceRangeStr);
  const srcSheetId = await getSheetId(sheets, spreadsheetId, src.sheetName);
  const idx = destCellStr.lastIndexOf("!");
  const destSheetName = idx === -1 ? src.sheetName : destCellStr.slice(0, idx).replace(/^'|'$/g, "");
  const destCellOnly = idx === -1 ? destCellStr : destCellStr.slice(idx + 1);
  const destSheetId = await getSheetId(sheets, spreadsheetId, destSheetName);
  const destParsed = parseRange(`${destSheetName}!${destCellOnly}`);
  const req = {
    cutPaste: {
      source: rangeToGridRange(srcSheetId, src),
      destination: {
        sheetId: destSheetId,
        rowIndex: destParsed.startRow - 1,
        columnIndex: destParsed.startCol,
      },
      pasteType: "PASTE_NORMAL",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 셀 단위 삽입/삭제 (행/열 통째가 아니라 부분 시프트) ----------

// shiftDim: "ROWS"(아래로 밀림) | "COLUMNS"(오른쪽으로 밀림)
async function insertCellRange(sheets, spreadsheetId, rangeStr, shiftDim) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("insertCells는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { insertRange: { range: rangeToGridRange(sheetId, r), shiftDimension: shiftDim } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// shiftDim: 삭제 후 채워질 방향 — "ROWS"(아래→위로 당김) | "COLUMNS"(오른쪽→왼쪽으로 당김)
async function deleteCellRange(sheets, spreadsheetId, rangeStr, shiftDim) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("deleteCells는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { deleteRange: { range: rangeToGridRange(sheetId, r), shiftDimension: shiftDim } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 행/열 그룹 접기/펼치기 (그룹 자체는 addDimensionGroup으로 이미 생성돼 있어야 함) ----------

// ⚠️ depth를 안 주면 "depth must be > 0" 400 에러남(실측) — addDimensionGroup으로 만든 중첩 없는
// 그룹은 항상 depth 1이므로 기본값 1을 준다. 그룹을 중첩해서 만들었으면 depth를 맞게 지정할 것.
async function setDimensionGroupCollapsed(sheets, spreadsheetId, sheetName, dimension, start1based, end1based, collapsed, depth = 1) {
  const sheetId = await getSheetId(sheets, spreadsheetId, sheetName);
  const req = {
    updateDimensionGroup: {
      dimensionGroup: {
        range: { sheetId, dimension, startIndex: start1based - 1, endIndex: end1based },
        depth,
        collapsed,
      },
      fields: "collapsed",
    },
  };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 행 순서 무작위 섞기 ----------

async function randomizeRows(sheets, spreadsheetId, rangeStr) {
  const r = parseRange(rangeStr);
  if (!r.startRow || !r.endRow) fail("randomize는 행 번호가 명시된 범위가 필요함");
  const sheetId = await getSheetId(sheets, spreadsheetId, r.sheetName);
  const req = { randomizeRange: { range: rangeToGridRange(sheetId, r) } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- 스프레드시트 로케일 / 시간대 ----------

async function setSpreadsheetLocale(sheets, spreadsheetId, locale, timeZone) {
  const props = {};
  const fields = [];
  if (locale) {
    props.locale = locale;
    fields.push("locale");
  }
  if (timeZone) {
    props.timeZone = timeZone;
    fields.push("timeZone");
  }
  const req = { updateSpreadsheetProperties: { properties: props, fields: fields.join(",") } };
  const res = await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [req] } });
  return res.data;
}

// ---------- ARRAYFORMULA 안전장치 ----------

// 시트 전체를 FORMULA 렌더링으로 한 번 긁어옴 — appendRowSafe가 "마지막 데이터 행 찾기"에도
// 이 결과를 재사용해서(같은 열의 데이터가 이미 여기 들어있음) 별도로 다시 fetch하지 않는다.
// (최적화: appendRow 한 번에 API 호출 2번 하던 걸 1번으로 줄임 — 예전엔 키열 스캔 fetch와
// ARRAYFORMULA 전체 스캔 fetch가 따로였는데, 후자가 전자를 포함하는 상위집합이라 합칠 수 있었음.)
async function scanSheetFormulaRows(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${quoteSheetName(sheetName)}!A1:ZZ2000`,
    valueRenderOption: "FORMULA",
  });
  return res.data.values || [];
}

function findArrayFormulaColumnsInRows(rows) {
  const cols = new Map(); // colIndex -> {row, formula}
  rows.forEach((row, rowIdx) => {
    row.forEach((cell, colIdx) => {
      if (typeof cell === "string" && /ARRAYFORMULA/i.test(cell) && !cols.has(colIdx)) {
        cols.set(colIdx, { row: rowIdx + 1, formula: cell });
      }
    });
  });
  return cols;
}

// rows를 미리 넘기면(appendRowSafe처럼 이미 스캔해둔 게 있으면) 재사용, 없으면 새로 fetch
async function guardArrayFormulaCollision(sheets, spreadsheetId, range, force, preScannedRows) {
  const { sheetName, startCol, endCol, startRow } = parseRange(range);
  const rows = preScannedRows || (await scanSheetFormulaRows(sheets, spreadsheetId, sheetName));
  const formulaCols = findArrayFormulaColumnsInRows(rows);
  if (formulaCols.size === 0) return;

  const hit = [];
  if (startCol === null || endCol === null) {
    // 순수 행 범위("5:5" 같은, 열 지정이 없는 range) - startCol/endCol이 null이라
    // 고정폭 for 루프(`c <= endCol`)를 돌리면 null이 0으로 강제형변환되어 A열(인덱스 0)
    // 딱 하나만 검사하고 끝나버리는 버그가 있었다(실측 확인). 열 범위가 없는 거니까
    // 감지된 ARRAYFORMULA 열 전부를 대상으로 검사한다.
    for (const [c, info] of formulaCols) {
      if (!startRow || startRow > info.row) {
        hit.push({ col: indexToColLetter(c), ...info });
      }
    }
  } else {
    for (let c = startCol; c <= endCol; c++) {
      if (formulaCols.has(c) && (!startRow || startRow > formulaCols.get(c).row)) {
        hit.push({ col: indexToColLetter(c), ...formulaCols.get(c) });
      }
    }
  }
  if (hit.length === 0) return;

  const detail = hit
    .map((h) => `  ${h.col}열 (${h.row}행에 ${h.formula.slice(0, 60)}${h.formula.length > 60 ? "..." : ""})`)
    .join("\n");

  if (!force) {
    fail(
      `이 범위가 자동계산(ARRAYFORMULA) 열과 겹쳐요. 직접 쓰면 수식이 깨질 수 있어요:\n${detail}\n` +
        `정말 이 열에 값을 쓰고 싶으면 명령 끝에 --force를 붙이세요. 아니면 다른 열만 써서 쓰기 범위를 좁히세요.`
    );
  } else {
    console.error(`⚠️  경고: --force로 자동계산 열에 직접 씁니다. 아래 열의 수식이 깨질 수 있어요:\n${detail}`);
  }
}

// ---------- appendRow: 실제 마지막 데이터 행을 스캔해서 정확한 위치에 쓰기 ----------

async function appendRowSafe(sheets, spreadsheetId, sheetName, keyColLetter, values, force, formatFrom, startColLetter = "A") {
  const keyColIdx = colLetterToIndex(keyColLetter);
  const scannedRows = await scanSheetFormulaRows(sheets, spreadsheetId, sheetName);
  let lastRow = 0;
  scannedRows.forEach((row, i) => {
    const cell = row[keyColIdx];
    if (cell !== undefined && cell !== "") lastRow = i + 1;
  });
  const targetRow = lastRow + 1;
  const startColIdx = colLetterToIndex(startColLetter);
  const endColLetter = indexToColLetter(startColIdx + values[0].length - 1);
  const range = `${quoteSheetName(sheetName)}!${startColLetter}${targetRow}:${endColLetter}${targetRow}`;

  await guardArrayFormulaCollision(sheets, spreadsheetId, range, force, scannedRows);

  const writeRes = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values },
  });

  let formatCopied = null;
  // formatFrom이 명시적으로 false가 아니면(기본 동작) 서식을 복사한다.
  // formatFrom이 숫자면 그 행에서, 안 주면 바로 위(직전 데이터) 행에서 복사.
  if (formatFrom !== false && lastRow > 0) {
    const sourceRow = typeof formatFrom === "number" ? formatFrom : lastRow;
    const sourceRange = `${quoteSheetName(sheetName)}!${startColLetter}${sourceRow}:${endColLetter}${sourceRow}`;
    await copyFormat(sheets, spreadsheetId, sourceRange, range);
    formatCopied = sourceRow;
  }

  return { targetRow, range, result: writeRes.data, formatCopiedFromRow: formatCopied };
}

// ---------- main ----------

async function main() {
  const rawArgs = process.argv.slice(2);
  const force = rawArgs.includes("--force");
  const formulas = rawArgs.includes("--formulas");
  const noFormat = rawArgs.includes("--noFormat");
  const formatFromArg = rawArgs.find((a) => a.startsWith("--formatFrom="));
  const formatFrom = noFormat ? false : formatFromArg ? parseInt(formatFromArg.split("=")[1], 10) : undefined;
  const startColArg = rawArgs.find((a) => a.startsWith("--startCol="));
  const startCol = startColArg ? startColArg.split("=")[1] : "A";
  const mergeTypeArg = rawArgs.find((a) => a.startsWith("--mergeType="));
  const mergeType = mergeTypeArg ? mergeTypeArg.split("=")[1] : "MERGE_ALL";
  const dimensionArg = rawArgs.find((a) => a.startsWith("--dimension="));
  const dimension = dimensionArg ? dimensionArg.split("=")[1] : "BOTH";
  const indexArg = rawArgs.find((a) => a.startsWith("--index="));
  const indexOpt = indexArg ? parseInt(indexArg.split("=")[1], 10) : undefined;
  const inheritBefore = rawArgs.includes("--inheritBefore");
  const strict = !rawArgs.includes("--noStrict");
  const showDropdown = !rawArgs.includes("--noDropdown");
  const matchCase = rawArgs.includes("--matchCase");
  const entireCell = rawArgs.includes("--entireCell");
  const regexFlag = rawArgs.includes("--regex");
  const warningOnly = rawArgs.includes("--warningOnly");
  const descArg = rawArgs.find((a) => a.startsWith("--description="));
  const editorsArg = rawArgs.find((a) => a.startsWith("--editors="));
  const titleArg = rawArgs.find((a) => a.startsWith("--title="));
  const anchorCellArg = rawArgs.find((a) => a.startsWith("--anchorCell="));
  const widthArg = rawArgs.find((a) => a.startsWith("--width="));
  const heightArg = rawArgs.find((a) => a.startsWith("--height="));
  const firstColorArg = rawArgs.find((a) => a.startsWith("--firstColor="));
  const secondColorArg = rawArgs.find((a) => a.startsWith("--secondColor="));
  const headerColorArg = rawArgs.find((a) => a.startsWith("--headerColor="));
  const minColorArg = rawArgs.find((a) => a.startsWith("--minColor="));
  const midColorArg = rawArgs.find((a) => a.startsWith("--midColor="));
  const maxColorArg = rawArgs.find((a) => a.startsWith("--maxColor="));
  const formulaArg = rawArgs.find((a) => a.startsWith("--formula="));
  const backgroundColorArg = rawArgs.find((a) => a.startsWith("--backgroundColor="));
  const colorArg = rawArgs.find((a) => a.startsWith("--color="));
  const styleArg = rawArgs.find((a) => a.startsWith("--style="));
  const KNOWN_FLAGS = [
    "--force", "--formulas", "--noFormat", "--inheritBefore", "--noStrict", "--noDropdown",
    "--matchCase", "--entireCell", "--regex", "--warningOnly",
  ];
  const KNOWN_PREFIXES = [
    "--formatFrom=", "--startCol=", "--mergeType=", "--dimension=", "--colWidth=", "--charWidth=",
    "--lineHeight=", "--index=", "--description=", "--editors=", "--title=", "--anchorCell=",
    "--width=", "--height=", "--firstColor=", "--secondColor=", "--headerColor=",
    "--minColor=", "--midColor=", "--maxColor=", "--formula=", "--backgroundColor=",
    "--color=", "--style=", "--sheetTarget=", "--range=", "--visibility=",
  ];
  // 알려진 플래그를 rawArgs 전체에서 무조건 걸러내면, findReplace의 찾을값처럼 사용자가
  // 넘긴 실제 데이터가 우연히 "--force" 같은 플래그 패턴과 겹칠 때 플래그로 오인돼서
  // 위치인자가 통째로 밀리는 문제가 있었다(실측 확인 가능). 이 CLI의 모든 사용법이
  // "명령 <위치인자...> [--플래그...]" 식으로 플래그를 항상 맨 뒤에 붙이는 관례라서,
  // 끝에서부터 훑다가 플래그가 아닌 첫 인자를 만나면 그 앞쪽은 위치인자로 보고 그대로 둔다.
  const isKnownFlag = (a) => KNOWN_FLAGS.includes(a) || KNOWN_PREFIXES.some((p) => a.startsWith(p));
  let argsEnd = rawArgs.length;
  while (argsEnd > 0 && isKnownFlag(rawArgs[argsEnd - 1])) argsEnd--;
  const args = rawArgs.slice(0, argsEnd);
  const [cmd, ...rest] = args;
  const sheets = getClient();

  switch (cmd) {
    case "tabs": {
      const [spreadsheetId] = rest;
      if (!spreadsheetId) fail("사용법: tabs <spreadsheetId>");
      const res = await sheets.spreadsheets.get({ spreadsheetId });
      const list = res.data.sheets.map((s) => ({
        sheetId: s.properties.sheetId,
        title: s.properties.title,
        rows: s.properties.gridProperties?.rowCount,
        cols: s.properties.gridProperties?.columnCount,
      }));
      printResult(list);
      break;
    }

    case "get": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail("사용법: get <spreadsheetId> <range> [--formulas]");
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range,
        valueRenderOption: formulas ? "FORMULA" : "FORMATTED_VALUE",
      });
      printResult(res.data);
      break;
    }

    case "batchGet": {
      const [spreadsheetId, rangesJson] = rest;
      if (!spreadsheetId || !rangesJson) fail("사용법: batchGet <spreadsheetId> <rangesJSON>");
      const ranges = parseJsonArg(rangesJson, '\'["Sheet1!A1:B2","Sheet2!A1:B2"]\'');
      const res = await sheets.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges,
        valueRenderOption: formulas ? "FORMULA" : "FORMATTED_VALUE",
      });
      printResult(res.data);
      break;
    }

    case "update": {
      const [spreadsheetId, range, valuesJson] = rest;
      if (!spreadsheetId || !range || !valuesJson) fail("사용법: update <spreadsheetId> <range> <valuesJSON> [--force]");
      const values = parseJsonArg(valuesJson, '\'[["a","b"],["c","d"]]\'');
      await guardArrayFormulaCollision(sheets, spreadsheetId, range, force);
      const res = await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });
      printResult(res.data);
      break;
    }

    case "append": {
      const [spreadsheetId, range, valuesJson] = rest;
      if (!spreadsheetId || !range || !valuesJson) fail("사용법: append <spreadsheetId> <range> <valuesJSON> [--force]");
      console.error("참고: append는 Sheets API의 '표 인식' 방식이라 빈 칸이 섞여 있으면 예상 밖 위치에 써질 수 있어요. 정확한 위치가 중요하면 appendRow를 쓰세요.");
      const values = parseJsonArg(valuesJson, '\'[["a","b"],["c","d"]]\'');
      await guardArrayFormulaCollision(sheets, spreadsheetId, range, force);
      const res = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });
      printResult(res.data);
      break;
    }

    case "appendRow": {
      const [spreadsheetId, sheetName, keyColLetter, valuesJson] = rest;
      if (!spreadsheetId || !sheetName || !keyColLetter || !valuesJson)
        fail(
          '사용법: appendRow <spreadsheetId> "SheetName" <키열문자예:C> <valuesJSON — 2차원, 1행> [--force] [--formatFrom=행번호] [--noFormat] [--startCol=열문자]'
        );
      const values = parseJsonArg(valuesJson, '\'[["a","b","c"]]\'');
      if (!Array.isArray(values) || !Array.isArray(values[0]))
        fail("valuesJSON은 2차원 배열이어야 함, 예: [[\"a\",\"b\",\"c\"]]");
      const out = await appendRowSafe(
        sheets,
        spreadsheetId,
        sheetName,
        keyColLetter,
        values,
        force,
        formatFrom,
        startCol
      );
      console.error(
        `→ ${out.targetRow}행에 씀 (${out.range})` +
          (out.formatCopiedFromRow ? `, 서식은 ${out.formatCopiedFromRow}행에서 복사함` : ", 서식 복사 안 함")
      );
      printResult(out.result);
      break;
    }

    case "copyFormat": {
      const [spreadsheetId, sourceRange, destRange] = rest;
      if (!spreadsheetId || !sourceRange || !destRange)
        fail('사용법: copyFormat <spreadsheetId> "Sheet1!A4:C4" "Sheet1!A5:C5"  (값은 안 바뀌고 서식만 복사됨)');
      const res = await copyFormat(sheets, spreadsheetId, sourceRange, destRange);
      printResult(res);
      break;
    }

    case "merge": {
      const [spreadsheetId, rangesJson] = rest;
      if (!spreadsheetId || !rangesJson)
        fail(
          '사용법: merge <spreadsheetId> <rangesJSON> [--mergeType=MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS]  예: merge id \'["Sheet1!B1:D1","Sheet1!F1:I1"]\''
        );
      let rangeList;
      try {
        rangeList = JSON.parse(rangesJson);
      } catch {
        fail('rangesJSON 파싱 실패 — 예: \'["Sheet1!B1:D1","Sheet1!F1:I1"]\'');
      }
      if (!Array.isArray(rangeList)) fail("rangesJSON은 문자열 배열이어야 함");
      const res = await mergeCells(sheets, spreadsheetId, rangeList, mergeType);
      printResult(res);
      break;
    }

    case "autoFit": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range)
        fail(
          '사용법: autoFit <spreadsheetId> "Sheet1!A1:I388" [--dimension=BOTH|COLUMNS|ROWS]  열/행 경계 더블클릭한 것과 동일 (내용 길이에 맞춰 자동조정)'
        );
      const res = await autoFit(sheets, spreadsheetId, range, dimension);
      printResult(res);
      break;
    }

    case "setWidth": {
      const [spreadsheetId, range, pixelsStr] = rest;
      if (!spreadsheetId || !range || !pixelsStr)
        fail('사용법: setWidth <spreadsheetId> "Sheet1!G:G" <pixels>  예: setWidth id "Sheet1!G:G" 260');
      const pixels = parseInt(pixelsStr, 10);
      if (Number.isNaN(pixels)) fail("pixels는 숫자여야 함");
      const res = await setColumnWidth(sheets, spreadsheetId, range, pixels);
      printResult(res);
      break;
    }

    case "wrap": {
      const [spreadsheetId, range, strategy] = rest;
      if (!spreadsheetId || !range)
        fail('사용법: wrap <spreadsheetId> "Sheet1!G1:G388" [WRAP|CLIP|OVERFLOW_CELL]  기본 WRAP (셀 안에서 줄바꿈, 세로로 길어짐)');
      const res = await setWrap(sheets, spreadsheetId, range, strategy || "WRAP");
      printResult(res);
      break;
    }

    case "fitWrap": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range)
        fail(
          '사용법: fitWrap <spreadsheetId> "Sheet1!G1:G388" [--colWidth=260] [--charWidth=6] [--lineHeight=21]  ' +
            "autoFit(ROWS)가 WRAP 높이를 못 잡을 때 글자수 기반으로 행 높이 직접 계산/적용 (wrap 명령을 먼저 걸어둘 것)"
        );
      const colWidthArg = rawArgs.find((a) => a.startsWith("--colWidth="));
      const charWidthArg = rawArgs.find((a) => a.startsWith("--charWidth="));
      const lineHeightArg = rawArgs.find((a) => a.startsWith("--lineHeight="));
      const opts = {};
      if (colWidthArg) opts.colWidth = parseInt(colWidthArg.split("=")[1], 10);
      if (charWidthArg) opts.charWidth = parseInt(charWidthArg.split("=")[1], 10);
      if (lineHeightArg) opts.lineHeight = parseInt(lineHeightArg.split("=")[1], 10);
      const res = await fitWrapRowHeights(sheets, spreadsheetId, range, opts);
      printResult(res);
      break;
    }

    case "addSheet": {
      const [spreadsheetId, title] = rest;
      if (!spreadsheetId || !title) fail('사용법: addSheet <spreadsheetId> "새 시트 이름" [--index=N]');
      const props = await addSheet(sheets, spreadsheetId, title, indexOpt);
      printResult(props);
      break;
    }

    case "deleteSheet": {
      const [spreadsheetId, sheetName] = rest;
      if (!spreadsheetId || !sheetName) fail('사용법: deleteSheet <spreadsheetId> "시트이름"  ⚠️ 되돌릴 수 없음');
      const res = await deleteSheet(sheets, spreadsheetId, sheetName);
      printResult(res);
      break;
    }

    case "renameSheet": {
      const [spreadsheetId, oldName, newName] = rest;
      if (!spreadsheetId || !oldName || !newName) fail('사용법: renameSheet <spreadsheetId> "기존이름" "새이름"');
      const res = await renameSheet(sheets, spreadsheetId, oldName, newName);
      printResult(res);
      break;
    }

    case "duplicateSheet": {
      const [spreadsheetId, sourceName, newName] = rest;
      if (!spreadsheetId || !sourceName) fail('사용법: duplicateSheet <spreadsheetId> "원본시트" [새이름] [--index=N]');
      const props = await duplicateSheet(sheets, spreadsheetId, sourceName, newName, indexOpt);
      printResult(props);
      break;
    }

    case "freeze": {
      const [spreadsheetId, sheetName, rowsStr, colsStr] = rest;
      if (!spreadsheetId || !sheetName)
        fail('사용법: freeze <spreadsheetId> "SheetName" [고정행수] [고정열수]  예: freeze id Sheet1 1 0');
      const frozenRows = rowsStr !== undefined ? parseInt(rowsStr, 10) : undefined;
      const frozenCols = colsStr !== undefined ? parseInt(colsStr, 10) : undefined;
      const res = await setFrozen(sheets, spreadsheetId, sheetName, frozenRows, frozenCols);
      printResult(res);
      break;
    }

    case "insertRows":
    case "insertCols": {
      const [spreadsheetId, sheetName, beforeStr, countStr] = rest;
      if (!spreadsheetId || !sheetName || !beforeStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <기준행/열번호(그 앞에 삽입)> [개수=1] [--inheritBefore]`);
      const before = parseInt(beforeStr, 10);
      const count = countStr ? parseInt(countStr, 10) : 1;
      const dim = cmd === "insertRows" ? "ROWS" : "COLUMNS";
      const res = await insertRowsOrCols(sheets, spreadsheetId, sheetName, dim, before, count, inheritBefore);
      printResult(res);
      break;
    }

    case "deleteRows":
    case "deleteCols": {
      const [spreadsheetId, sheetName, startStr, endStr] = rest;
      if (!spreadsheetId || !sheetName || !startStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <시작번호> [끝번호=시작과동일]  ⚠️ 되돌릴 수 없음`);
      const start = parseInt(startStr, 10);
      const end = endStr ? parseInt(endStr, 10) : start;
      const dim = cmd === "deleteRows" ? "ROWS" : "COLUMNS";
      const res = await deleteRowsOrCols(sheets, spreadsheetId, sheetName, dim, start, end);
      printResult(res);
      break;
    }

    case "sort": {
      const [spreadsheetId, range, sortSpecJson] = rest;
      if (!spreadsheetId || !range || !sortSpecJson)
        fail(
          '사용법: sort <spreadsheetId> "Sheet1!A2:F6" <sortSpecJSON>  헤더 제외한 범위. 예: \'[{"col":"D","order":"DESC"}]\''
        );
      const sortSpecs = parseJsonArg(sortSpecJson, '\'[{"col":"D","order":"DESC"}]\'');
      const res = await sortRange(sheets, spreadsheetId, range, sortSpecs);
      printResult(res);
      break;
    }

    case "validate": {
      const [spreadsheetId, range, listJson] = rest;
      if (!spreadsheetId || !range || !listJson)
        fail(
          '사용법: validate <spreadsheetId> "Sheet1!D2:D100" <valuesJSON> [--noStrict] [--noDropdown]  예: validate id "Sheet1!D2:D100" \'["대기","진행중","완료"]\''
        );
      const listValues = parseJsonArg(listJson, '\'["대기","진행중","완료"]\'');
      const res = await setValidation(sheets, spreadsheetId, range, listValues, { strict, showDropdown });
      printResult(res);
      break;
    }

    case "clearValidation": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail('사용법: clearValidation <spreadsheetId> "Sheet1!D2:D100"');
      const res = await clearValidation(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "addNamedRange": {
      const [spreadsheetId, name, range] = rest;
      if (!spreadsheetId || !name || !range) fail('사용법: addNamedRange <spreadsheetId> <이름> "Sheet1!A1:F6"');
      const res = await addNamedRange(sheets, spreadsheetId, name, range);
      printResult(res);
      break;
    }

    case "deleteNamedRange": {
      const [spreadsheetId, namedRangeId] = rest;
      if (!spreadsheetId || !namedRangeId)
        fail("사용법: deleteNamedRange <spreadsheetId> <namedRangeId>  (id는 metadata로 확인)");
      const res = await deleteNamedRange(sheets, spreadsheetId, namedRangeId);
      printResult(res);
      break;
    }

    case "findReplace": {
      const [spreadsheetId, sheetName, find, replacement] = rest;
      if (!spreadsheetId || !sheetName || find === undefined || replacement === undefined)
        fail(
          '사용법: findReplace <spreadsheetId> <"SheetName"|ALL> <찾을값> <바꿀값> [--matchCase] [--entireCell] [--regex]'
        );
      const res = await findReplace(sheets, spreadsheetId, sheetName, find, replacement, {
        matchCase,
        entireCell,
        regex: regexFlag,
      });
      printResult(res);
      break;
    }

    case "numberFormat": {
      const [spreadsheetId, range, typeOrPattern] = rest;
      if (!spreadsheetId || !range || !typeOrPattern)
        fail(
          '사용법: numberFormat <spreadsheetId> <range> <타입|패턴>  타입: NUMBER,INTEGER,CURRENCY,PERCENT,DATE,TIME,DATE_TIME,SCIENTIFIC  또는 커스텀 패턴(예: "0.0%")'
        );
      const res = await setNumberFormat(sheets, spreadsheetId, range, typeOrPattern);
      printResult(res);
      break;
    }

    case "note": {
      const [spreadsheetId, range, ...noteParts] = rest;
      if (!spreadsheetId || !range || noteParts.length === 0)
        fail('사용법: note <spreadsheetId> <range> <메모내용>  빈 문자열("")로 주면 메모 삭제');
      const res = await setNote(sheets, spreadsheetId, range, noteParts.join(" "));
      printResult(res);
      break;
    }

    case "protect": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range)
        fail(
          '사용법: protect <spreadsheetId> <range> [--description=설명] [--warningOnly] [--editors=email1,email2]  editors 안 주면 본인만 편집 가능'
        );
      const res = await protectRange(sheets, spreadsheetId, range, {
        description: descArg ? descArg.split("=").slice(1).join("=") : undefined,
        warningOnly,
        editors: editorsArg ? editorsArg.split("=").slice(1).join("=").split(",") : undefined,
      });
      printResult(res);
      break;
    }

    case "addBanding": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range)
        fail(
          '사용법: addBanding <spreadsheetId> <range> [--firstColor=#RRGGBB] [--secondColor=#RRGGBB] [--headerColor=#RRGGBB]'
        );
      const res = await addBanding(sheets, spreadsheetId, range, {
        firstColor: firstColorArg ? firstColorArg.split("=")[1] : undefined,
        secondColor: secondColorArg ? secondColorArg.split("=")[1] : undefined,
        headerColor: headerColorArg ? headerColorArg.split("=")[1] : undefined,
      });
      printResult(res);
      break;
    }

    case "addChart": {
      const [spreadsheetId, sheetName, chartType, dataRange] = rest;
      if (!spreadsheetId || !sheetName || !chartType || !dataRange)
        fail(
          '사용법: addChart <spreadsheetId> "SheetName" <COLUMN|LINE|BAR|AREA|SCATTER> "Sheet1!A1:B6" [--title=] [--anchorCell=D1] [--width=] [--height=]  ' +
            "데이터 첫 열=X축(도메인), 나머지 열=시리즈. 헤더 행 포함해서 범위 지정할 것."
        );
      const res = await addChart(sheets, spreadsheetId, sheetName, chartType, dataRange, {
        title: titleArg ? titleArg.split("=").slice(1).join("=") : undefined,
        anchorCell: anchorCellArg ? anchorCellArg.split("=")[1] : undefined,
        width: widthArg ? parseInt(widthArg.split("=")[1], 10) : undefined,
        height: heightArg ? parseInt(heightArg.split("=")[1], 10) : undefined,
      });
      printResult(res);
      break;
    }

    case "setFilter": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail('사용법: setFilter <spreadsheetId> "Sheet1!A1:F6"  (헤더 행 포함 범위)');
      const res = await setBasicFilter(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "clearFilter": {
      const [spreadsheetId, sheetName] = rest;
      if (!spreadsheetId || !sheetName) fail('사용법: clearFilter <spreadsheetId> "SheetName"');
      const res = await clearBasicFilter(sheets, spreadsheetId, sheetName);
      printResult(res);
      break;
    }

    case "splitText": {
      const [spreadsheetId, range, delimiter] = rest;
      if (!spreadsheetId || !range || delimiter === undefined)
        fail('사용법: splitText <spreadsheetId> "Sheet1!A2:A10" <구분자>  예: 쉼표는 ","  탭은 "\\t"');
      const res = await splitTextToColumns(sheets, spreadsheetId, range, delimiter);
      printResult(res);
      break;
    }

    case "addPivot": {
      const [spreadsheetId, sourceRange, anchorSheetName, anchorCell, rowsJson, valuesJson, colsJson] = rest;
      if (!spreadsheetId || !sourceRange || !anchorSheetName || !anchorCell || !rowsJson || !valuesJson)
        fail(
          '사용법: addPivot <spreadsheetId> <sourceRange(헤더포함)> <대상시트> <앵커셀예:H1> <rowsJSON> <valuesJSON> [colsJSON]\n' +
            '  예: addPivot id "Sheet1!A1:F6" Sheet1 H1 \'[{"col":"C"}]\' \'[{"col":"D","fn":"COUNTA"}]\''
        );
      const rowsSpec = parseJsonArg(rowsJson, '\'[{"col":"C"}]\'');
      const valuesSpec = parseJsonArg(valuesJson, '\'[{"col":"D","fn":"COUNTA"}]\'');
      const colsSpec = colsJson ? parseJsonArg(colsJson, '\'[{"col":"E"}]\'') : [];
      const res = await addPivotTable(sheets, spreadsheetId, sourceRange, anchorSheetName, anchorCell, rowsSpec, valuesSpec, colsSpec);
      printResult(res);
      break;
    }

    case "condFormat": {
      const [spreadsheetId, range, mode] = rest;
      if (!spreadsheetId || !range || !mode)
        fail(
          '사용법: condFormat <spreadsheetId> <range> <colorScale|formula> [--minColor=] [--midColor=] [--maxColor=] [--formula="=..."] [--backgroundColor=]\n' +
            '  예1: condFormat id "Sheet1!D2:D6" colorScale --minColor=#F4C7C3 --maxColor=#B7E1CD\n' +
            '  예2: condFormat id "Sheet1!A2:F6" formula --formula=\'=$D2="완료"\' --backgroundColor=#D9EAD3'
        );
      const res = await addConditionalFormat(sheets, spreadsheetId, range, mode, {
        minColor: minColorArg ? minColorArg.split("=")[1] : undefined,
        midColor: midColorArg ? midColorArg.split("=")[1] : undefined,
        maxColor: maxColorArg ? maxColorArg.split("=")[1] : undefined,
        formula: formulaArg ? formulaArg.split("=").slice(1).join("=") : undefined,
        backgroundColor: backgroundColorArg ? backgroundColorArg.split("=")[1] : undefined,
      });
      printResult(res);
      break;
    }

    case "copySheetTo": {
      const [spreadsheetId, sheetName, destSpreadsheetId] = rest;
      if (!spreadsheetId || !sheetName || !destSpreadsheetId)
        fail('사용법: copySheetTo <spreadsheetId> "SheetName" <대상spreadsheetId>');
      const res = await copySheetTo(sheets, spreadsheetId, sheetName, destSpreadsheetId);
      printResult(res);
      break;
    }

    case "renameSpreadsheet": {
      const [spreadsheetId, newTitle] = rest;
      if (!spreadsheetId || !newTitle) fail('사용법: renameSpreadsheet <spreadsheetId> "새 제목"');
      const res = await renameSpreadsheet(sheets, spreadsheetId, newTitle);
      printResult(res);
      break;
    }

    case "setBorder": {
      const [spreadsheetId, range, sides] = rest;
      if (!spreadsheetId || !range || !sides)
        fail(
          '사용법: setBorder <spreadsheetId> <range> <all|outer|inner|top,bottom,left,right,innerHorizontal,innerVertical> [--color=#000000] [--style=SOLID|DASHED|DOTTED|DOUBLE|SOLID_MEDIUM|SOLID_THICK]'
        );
      const res = await setBorder(sheets, spreadsheetId, range, sides, {
        color: colorArg ? colorArg.split("=")[1] : undefined,
        style: styleArg ? styleArg.split("=")[1] : undefined,
      });
      printResult(res);
      break;
    }

    case "unmerge": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail('사용법: unmerge <spreadsheetId> <range>');
      const res = await unmergeCells(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "setRowHeight": {
      const [spreadsheetId, range, pixelsStr] = rest;
      if (!spreadsheetId || !range || !pixelsStr)
        fail('사용법: setRowHeight <spreadsheetId> "Sheet1!3:3" <pixels>');
      const pixels = parseInt(pixelsStr, 10);
      if (Number.isNaN(pixels)) fail("pixels는 숫자여야 함");
      const res = await setRowHeight(sheets, spreadsheetId, range, pixels);
      printResult(res);
      break;
    }

    case "moveRows":
    case "moveCols": {
      const [spreadsheetId, sheetName, startStr, endStr, destStr] = rest;
      if (!spreadsheetId || !sheetName || !startStr || !endStr || !destStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <시작번호> <끝번호> <이동목적지(그 앞으로)>`);
      const dim = cmd === "moveRows" ? "ROWS" : "COLUMNS";
      const res = await moveRowsOrCols(
        sheets,
        spreadsheetId,
        sheetName,
        dim,
        parseInt(startStr, 10),
        parseInt(endStr, 10),
        parseInt(destStr, 10)
      );
      printResult(res);
      break;
    }

    case "deleteProtect": {
      const [spreadsheetId, protectedRangeId] = rest;
      if (!spreadsheetId || !protectedRangeId) fail("사용법: deleteProtect <spreadsheetId> <protectedRangeId>");
      const res = await deleteProtectedRange(sheets, spreadsheetId, protectedRangeId);
      printResult(res);
      break;
    }

    case "deleteBanding": {
      const [spreadsheetId, bandedRangeId] = rest;
      if (!spreadsheetId || !bandedRangeId) fail("사용법: deleteBanding <spreadsheetId> <bandedRangeId>");
      const res = await deleteBandingRange(sheets, spreadsheetId, bandedRangeId);
      printResult(res);
      break;
    }

    case "deleteCondFormat": {
      const [spreadsheetId, sheetName, index] = rest;
      if (!spreadsheetId || !sheetName || index === undefined)
        fail('사용법: deleteCondFormat <spreadsheetId> "SheetName" <index>  (0-based, metadata로 확인)');
      const res = await deleteConditionalFormat(sheets, spreadsheetId, sheetName, index);
      printResult(res);
      break;
    }

    case "deleteChart": {
      const [spreadsheetId, chartId] = rest;
      if (!spreadsheetId || !chartId) fail("사용법: deleteChart <spreadsheetId> <chartId>");
      const res = await deleteChart(sheets, spreadsheetId, chartId);
      printResult(res);
      break;
    }

    case "moveChart": {
      const [spreadsheetId, chartId, sheetName, anchorCell] = rest;
      if (!spreadsheetId || !chartId || !sheetName || !anchorCell)
        fail('사용법: moveChart <spreadsheetId> <chartId> "SheetName" <앵커셀예:D1> [--width=] [--height=]');
      const res = await moveChart(
        sheets,
        spreadsheetId,
        chartId,
        sheetName,
        anchorCell,
        widthArg ? parseInt(widthArg.split("=")[1], 10) : undefined,
        heightArg ? parseInt(heightArg.split("=")[1], 10) : undefined
      );
      printResult(res);
      break;
    }

    case "setTabColor": {
      const [spreadsheetId, sheetName, hex] = rest;
      if (!spreadsheetId || !sheetName || !hex) fail('사용법: setTabColor <spreadsheetId> "SheetName" <#RRGGBB>');
      const res = await setTabColor(sheets, spreadsheetId, sheetName, hex);
      printResult(res);
      break;
    }

    case "hideSheet":
    case "showSheet": {
      const [spreadsheetId, sheetName] = rest;
      if (!spreadsheetId || !sheetName) fail(`사용법: ${cmd} <spreadsheetId> "SheetName"`);
      const res = await setSheetHidden(sheets, spreadsheetId, sheetName, cmd === "hideSheet");
      printResult(res);
      break;
    }

    case "setGridlines": {
      const [spreadsheetId, sheetName, onOff] = rest;
      if (!spreadsheetId || !sheetName || !onOff)
        fail('사용법: setGridlines <spreadsheetId> "SheetName" <on|off>');
      const res = await setGridlines(sheets, spreadsheetId, sheetName, onOff === "on");
      printResult(res);
      break;
    }

    case "hideRows":
    case "hideCols":
    case "showRows":
    case "showCols": {
      const [spreadsheetId, sheetName, startStr, endStr] = rest;
      if (!spreadsheetId || !sheetName || !startStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <시작번호> [끝번호=시작과동일]`);
      const start = parseInt(startStr, 10);
      const end = endStr ? parseInt(endStr, 10) : start;
      const dim = cmd.endsWith("Rows") ? "ROWS" : "COLUMNS";
      const hidden = cmd.startsWith("hide");
      const res = await setDimensionHidden(sheets, spreadsheetId, sheetName, dim, start, end, hidden);
      printResult(res);
      break;
    }

    case "groupRows":
    case "groupCols": {
      const [spreadsheetId, sheetName, startStr, endStr] = rest;
      if (!spreadsheetId || !sheetName || !startStr || !endStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <시작번호> <끝번호>  (접기/펼치기 그룹 생성)`);
      const dim = cmd === "groupRows" ? "ROWS" : "COLUMNS";
      const res = await addDimensionGroup(sheets, spreadsheetId, sheetName, dim, parseInt(startStr, 10), parseInt(endStr, 10));
      printResult(res);
      break;
    }

    case "ungroupRows":
    case "ungroupCols": {
      const [spreadsheetId, sheetName, startStr, endStr] = rest;
      if (!spreadsheetId || !sheetName || !startStr || !endStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <시작번호> <끝번호>`);
      const dim = cmd === "ungroupRows" ? "ROWS" : "COLUMNS";
      const res = await deleteDimensionGroup(sheets, spreadsheetId, sheetName, dim, parseInt(startStr, 10), parseInt(endStr, 10));
      printResult(res);
      break;
    }

    case "addFilterView": {
      const [spreadsheetId, range, title] = rest;
      if (!spreadsheetId || !range || !title)
        fail('사용법: addFilterView <spreadsheetId> <range(헤더포함)> <제목>');
      const res = await addFilterView(sheets, spreadsheetId, range, title);
      printResult(res);
      break;
    }

    case "deleteFilterView": {
      const [spreadsheetId, filterId] = rest;
      if (!spreadsheetId || !filterId) fail("사용법: deleteFilterView <spreadsheetId> <filterId>");
      const res = await deleteFilterView(sheets, spreadsheetId, filterId);
      printResult(res);
      break;
    }

    case "duplicateFilterView": {
      const [spreadsheetId, filterId] = rest;
      if (!spreadsheetId || !filterId) fail("사용법: duplicateFilterView <spreadsheetId> <filterId>");
      const res = await duplicateFilterView(sheets, spreadsheetId, filterId);
      printResult(res);
      break;
    }

    case "addMetadata": {
      const [spreadsheetId, key, value] = rest;
      if (!spreadsheetId || !key || value === undefined)
        fail(
          '사용법: addMetadata <spreadsheetId> <key> <value> [--sheetTarget="SheetName"] [--range="3:5"] [--visibility=DOCUMENT|PROJECT]'
        );
      const sheetTargetArg = rawArgs.find((a) => a.startsWith("--sheetTarget="));
      const rangeTargetArg = rawArgs.find((a) => a.startsWith("--range="));
      const visibilityArg = rawArgs.find((a) => a.startsWith("--visibility="));
      const res = await addDeveloperMetadata(sheets, spreadsheetId, key, value, {
        sheetName: sheetTargetArg ? sheetTargetArg.split("=").slice(1).join("=") : undefined,
        range: rangeTargetArg ? rangeTargetArg.split("=").slice(1).join("=") : undefined,
        visibility: visibilityArg ? visibilityArg.split("=")[1] : undefined,
      });
      printResult(res);
      break;
    }

    case "findMetadata": {
      const [spreadsheetId, key] = rest;
      if (!spreadsheetId || !key) fail("사용법: findMetadata <spreadsheetId> <key>");
      const res = await searchDeveloperMetadata(sheets, spreadsheetId, key);
      printResult(res);
      break;
    }

    case "deleteMetadata": {
      const [spreadsheetId, key] = rest;
      if (!spreadsheetId || !key) fail("사용법: deleteMetadata <spreadsheetId> <key>");
      const res = await deleteDeveloperMetadata(sheets, spreadsheetId, key);
      printResult(res);
      break;
    }

    case "trimWhitespace": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail('사용법: trimWhitespace <spreadsheetId> <range>');
      const res = await trimWhitespace(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "deleteDuplicates": {
      const [spreadsheetId, range, colsJson] = rest;
      if (!spreadsheetId || !range) fail('사용법: deleteDuplicates <spreadsheetId> <range> [비교열JSON예:\'["B","C"]\']');
      const compareCols = colsJson ? parseJsonArg(colsJson, '\'["B","C"]\'') : [];
      const res = await deleteDuplicateRows(sheets, spreadsheetId, range, compareCols);
      printResult(res);
      break;
    }

    case "autoFillRange": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range)
        fail('사용법: autoFillRange <spreadsheetId> <range>  패턴 있는 셀+빈 셀 전체 범위 (드래그 채우기와 동일)');
      const res = await autoFillRange(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "cutPaste": {
      const [spreadsheetId, sourceRange, destCell] = rest;
      if (!spreadsheetId || !sourceRange || !destCell)
        fail(
          '사용법: cutPaste <spreadsheetId> <소스range> <목적지시작셀>  예: cutPaste id "Sheet1!A1:B5" "Sheet1!D1"  ⚠️ 원본은 지워짐(잘라내기)'
        );
      const res = await cutPaste(sheets, spreadsheetId, sourceRange, destCell);
      printResult(res);
      break;
    }

    case "insertCells": {
      const [spreadsheetId, range, shiftDim] = rest;
      if (!spreadsheetId || !range || !shiftDim)
        fail('사용법: insertCells <spreadsheetId> <range> <ROWS|COLUMNS>  (부분 범위만 밀어냄, 행/열 통째 아님)');
      const res = await insertCellRange(sheets, spreadsheetId, range, normalizeDimension(shiftDim));
      printResult(res);
      break;
    }

    case "deleteCells": {
      const [spreadsheetId, range, shiftDim] = rest;
      if (!spreadsheetId || !range || !shiftDim)
        fail('사용법: deleteCells <spreadsheetId> <range> <ROWS|COLUMNS>  ⚠️ 되돌릴 수 없음, 부분 범위만 당겨짐');
      const res = await deleteCellRange(sheets, spreadsheetId, range, normalizeDimension(shiftDim));
      printResult(res);
      break;
    }

    case "collapseGroup":
    case "expandGroup": {
      const [spreadsheetId, sheetName, dimStr, startStr, endStr] = rest;
      if (!spreadsheetId || !sheetName || !dimStr || !startStr || !endStr)
        fail(`사용법: ${cmd} <spreadsheetId> "SheetName" <ROWS|COLUMNS> <시작번호> <끝번호>  (addDimensionGroup으로 그룹이 먼저 있어야 함)`);
      const res = await setDimensionGroupCollapsed(
        sheets,
        spreadsheetId,
        sheetName,
        normalizeDimension(dimStr),
        parseInt(startStr, 10),
        parseInt(endStr, 10),
        cmd === "collapseGroup"
      );
      printResult(res);
      break;
    }

    case "randomize": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail("사용법: randomize <spreadsheetId> <range>  행 순서를 무작위로 섞음");
      const res = await randomizeRows(sheets, spreadsheetId, range);
      printResult(res);
      break;
    }

    case "setLocale": {
      const [spreadsheetId, locale, timeZone] = rest;
      if (!spreadsheetId || (!locale && !timeZone))
        fail('사용법: setLocale <spreadsheetId> <locale예:ko_KR> [timeZone예:Asia/Seoul]  둘 중 하나만 줘도 됨(""로 생략)');
      const res = await setSpreadsheetLocale(
        sheets,
        spreadsheetId,
        locale && locale !== "" ? locale : undefined,
        timeZone && timeZone !== "" ? timeZone : undefined
      );
      printResult(res);
      break;
    }

    case "clear": {
      const [spreadsheetId, range] = rest;
      if (!spreadsheetId || !range) fail("사용법: clear <spreadsheetId> <range>");
      const res = await sheets.spreadsheets.values.clear({ spreadsheetId, range });
      printResult(res.data);
      break;
    }

    case "metadata": {
      const [spreadsheetId, flag] = rest;
      if (!spreadsheetId) fail("사용법: metadata <spreadsheetId> [--grid]");
      const res = await sheets.spreadsheets.get({
        spreadsheetId,
        includeGridData: flag === "--grid",
      });
      printResult(res.data);
      break;
    }

    case "batchUpdate": {
      const [spreadsheetId, requestsJson] = rest;
      if (!spreadsheetId || !requestsJson) fail("사용법: batchUpdate <spreadsheetId> <requestsJSON>");
      const requests = parseJsonArg(requestsJson, '\'[{"updateCells": {...}}]\'');
      const res = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests },
      });
      printResult(res.data);
      break;
    }

    default:
      console.log(`알 수 없는 명령: ${cmd}
사용 가능: tabs, get, batchGet, update, append, appendRow, copyFormat, merge, autoFit, setWidth, wrap, fitWrap,
  addSheet, deleteSheet, renameSheet, duplicateSheet, freeze, insertRows, insertCols, deleteRows, deleteCols, sort,
  validate, clearValidation, findReplace, numberFormat, note, addNamedRange, deleteNamedRange, protect, addBanding,
  addChart, setFilter, clearFilter, splitText, addPivot, condFormat, copySheetTo, renameSpreadsheet, setBorder,
  unmerge, setRowHeight, moveRows, moveCols,
  deleteProtect, deleteBanding, deleteCondFormat, deleteChart, moveChart, setTabColor, hideSheet, showSheet,
  setGridlines, hideRows, hideCols, showRows, showCols, groupRows, groupCols, ungroupRows, ungroupCols,
  addFilterView, deleteFilterView, duplicateFilterView, addMetadata, findMetadata, deleteMetadata,
  trimWhitespace, deleteDuplicates, autoFillRange,
  cutPaste, insertCells, deleteCells, collapseGroup, expandGroup, randomize, setLocale,
  clear, metadata, batchUpdate
자세한 사용법은 sheets.js 상단 주석 참고.`);
      process.exit(1);
  }
}

main().catch((err) => {
  fail(err.response?.data ? JSON.stringify(err.response.data) : err.message);
});
