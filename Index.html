
/**
 * 數位科技概論｜版本四
 * Google 帳號辨識＋每位使用者獨立試算表
 *
 * 部署時必須選：
 * 執行身分：使用網頁應用程式的使用者
 */

const SHEET_NAME = "遊戲紀錄";
const PROPERTY_KEY = "MY_GAME_SPREADSHEET_ID";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("四位數幾A幾B｜版本四")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * 使用者按下同意說明後才呼叫。
 * 每個 Google 使用者只會建立自己的試算表一次。
 */
function getOrCreateProfile() {
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    throw new Error(
      "無法辨識目前的 Google 帳號。請確認部署設定為「使用網頁應用程式的使用者」並重新授權。"
    );
  }

  const userProperties = PropertiesService.getUserProperties();
  let spreadsheetId = userProperties.getProperty(PROPERTY_KEY);
  let spreadsheet;

  if (spreadsheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      spreadsheetId = null;
    }
  }

  // 第一次使用才建立試算表
  if (!spreadsheetId) {
    spreadsheet = SpreadsheetApp.create("幾A幾B_版本四_我的遊戲紀錄");

    const oldSheet = spreadsheet.getSheets()[0];
    oldSheet.setName(SHEET_NAME);

    oldSheet.getRange(1, 1, 1, 2).setValues([
      ["遊戲資料(JSON)", "最後更新時間"]
    ]);

    oldSheet.setFrozenRows(1);

    userProperties.setProperty(
      PROPERTY_KEY,
      spreadsheet.getId()
    );
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("找不到遊戲紀錄工作表。");
  }

  let games = [];

  if (sheet.getLastRow() >= 2) {
    const json = sheet.getRange(2, 1).getValue();

    if (json) {
      try {
        games = JSON.parse(json);
      } catch (error) {
        throw new Error("遊戲資料讀取失敗，請檢查試算表內容。");
      }
    }
  }

  if (!Array.isArray(games)) {
    games = [];
  }

  return {
    success: true,
    email: email,
    spreadsheetUrl: spreadsheet.getUrl(),
    games: games
  };
}

/**
 * 儲存目前登入者自己的遊戲資料。
 */
function saveMyGames(games) {
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    throw new Error("無法確認 Google 帳號，請重新授權。");
  }

  if (!Array.isArray(games)) {
    throw new Error("遊戲資料格式錯誤。");
  }

  if (games.length > 200) {
    throw new Error("遊戲紀錄數量超過上限。");
  }

  const json = JSON.stringify(games);

  if (json.length > 40000) {
    throw new Error("遊戲資料太大，請減少紀錄後再試。");
  }

  const userProperties = PropertiesService.getUserProperties();
  const spreadsheetId = userProperties.getProperty(PROPERTY_KEY);

  if (!spreadsheetId) {
    throw new Error("尚未建立個人試算表，請重新進入遊戲。");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("找不到遊戲紀錄工作表。");
  }

  // 每位使用者只會寫入自己的試算表
  sheet.getRange(2, 1, 1, 2).setValues([
    [json, new Date()]
  ]);

  return {
    success: true,
    savedAt: new Date().toISOString()
  };
}

