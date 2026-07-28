// 최초 1회 실행용: Google OAuth 동의 후 refresh token을 로컬에 저장한다.
// 사용법: node auth.js
//
// 사전 준비 (사용자가 Google Cloud Console에서 직접 해야 함):
// 1. https://console.cloud.google.com 에서 프로젝트 생성
// 2. "API 및 서비스 > 라이브러리"에서 Google Sheets API 활성화
// 3. "API 및 서비스 > 사용자 인증 정보"에서 OAuth 클라이언트 ID 생성
//    - 애플리케이션 유형: "데스크톱 앱"
//    - 생성된 클라이언트 ID / 클라이언트 보안 비밀번호를 config.json에 입력
// 4. OAuth 동의 화면에서 테스트 사용자로 본인 계정 추가(앱이 "테스트" 상태인 경우)

const fs = require("fs");
const path = require("path");
const http = require("http");
const { google } = require("googleapis");

const CONFIG_PATH = path.join(__dirname, "config.json");
const TOKEN_PATH = path.join(__dirname, "token.json");
const REDIRECT_PORT = 51789;
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/callback`;

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive.readonly",
];

function loadConfig() {
  if (!fs.existsSync(CONFIG_PATH)) {
    console.error(`config.json이 없습니다. config.example.json을 복사해서 client_id/client_secret을 채워주세요.`);
    process.exit(1);
  }
  return JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));
}

async function main() {
  const { client_id, client_secret } = loadConfig();
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, REDIRECT_URI);

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent", // refresh token을 매번 확실히 받기 위해
    scope: SCOPES,
  });

  console.log("\n아래 URL을 브라우저로 열어서 Google 계정으로 로그인/동의해주세요:\n");
  console.log(authUrl, "\n");

  const code = await new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const url = new URL(req.url, REDIRECT_URI);
      const code = url.searchParams.get("code");
      const error = url.searchParams.get("error");
      if (error) {
        res.end("인증 실패, 터미널을 확인하세요. 창을 닫아도 됩니다.");
        server.close();
        reject(new Error(error));
        return;
      }
      res.end("인증 완료! 이 창은 닫으셔도 됩니다.");
      server.close();
      resolve(code);
    });
    server.listen(REDIRECT_PORT, () => {
      console.log(`(로컬에서 인증 코드를 기다리는 중: ${REDIRECT_URI})`);
    });
  });

  const { tokens } = await oAuth2Client.getToken(code);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2), { mode: 0o600 });
  console.log(`\n토큰 저장 완료: ${TOKEN_PATH}`);
  console.log("이제 sheets.js로 스프레드시트를 읽고 쓸 수 있습니다.");
}

main().catch((err) => {
  console.error("인증 실패:", err.message);
  process.exit(1);
});
