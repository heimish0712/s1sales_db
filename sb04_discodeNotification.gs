/****************************************************
 * 영업지원요청 G열 변경 감지 → Discord 알림
 *
 * 감시 시트: 영업지원요청
 * 기준 ID: A열 접수번호
 * 감시 값: G열
 *
 * 사용 순서:
 * 1. 이 코드 전체 붙여넣기
 * 2. setSalesSupportWebhookUrlOnce_() 안에 웹훅 URL 넣기
 * 3. setSalesSupportWebhookUrlOnce_() 최초 1회 실행
 * 4. saveSalesSupportCurrentState() 최초 1회 실행
 * 5. installSalesSupportAlertTrigger() 최초 1회 실행
 ****************************************************/

var SALES_SUPPORT_ALERT_CONFIG = {
  SHEET_NAME: "영업지원요청",
  START_ROW: 2,
  END_ROW: 1000,

  ID_COL: 1,       // A열: 접수번호
  WATCH_COL: 7,    // G열: 감시할 내용

  STATE_PROP_KEY: "SALES_SUPPORT_LAST_ID_G_MAP",
  WEBHOOK_PROP_KEY: "SALES_SUPPORT_DISCORD_WEBHOOK_URL",

  MAX_ALERT_ITEMS: 20
};


/**
 * 최초 1회만 실행.
 * 네 Discord 웹훅 URL을 Script Properties에 저장함.
 *
 * ⚠️ 여기 문자열 안에 웹훅 URL 넣고 실행한 뒤,
 * 보안 신경 쓰이면 이 함수 안 URL은 다시 지워도 됨.
 */
function setSalesSupportWebhookUrlOnce() {
  var webhookUrl = "https://discord.com/api/webhooks/1516614305305329785/7SVOGX2cOSDIyetVyPF9SGlCFj5HRNVFYASGhUvKWakGWn8xGmNt7hUfpfDWL94zcaUF";

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SALES_SUPPORT_ALERT_CONFIG.WEBHOOK_PROP_KEY,
      webhookUrl
    );

  Logger.log("Discord 웹훅 URL 저장 완료");
}


/**
 * 현재 상태를 저장함.
 * 최초 1회 실행해야 기존 G열 값들이 한꺼번에 알림으로 터지는 참사를 막음.
 */
function saveSalesSupportCurrentState() {
  var currentMap = getSalesSupportCurrentIdGMap_();

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY,
      JSON.stringify(currentMap)
    );

  Logger.log("영업지원요청 현재 상태 저장 완료");
}


/**
 * 실제 감시 함수.
 * 시간 기반 트리거가 이 함수를 반복 실행함.
 */
function checkSalesSupportNewValues() {
  var currentMap = getSalesSupportCurrentIdGMap_();

  var props = PropertiesService.getScriptProperties();
  var oldRaw = props.getProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);
  var oldMap = oldRaw ? JSON.parse(oldRaw) : {};

  var changedItems = [];

  Object.keys(currentMap).forEach(function(id) {
    var currentValue = currentMap[id] || "";
    var oldValue = oldMap[id] || "";

    // G열이 비어 있으면 알림 안 함
    if (!currentValue) return;

    // 기존 값과 현재 값이 다르면 알림
    if (currentValue !== oldValue) {
      changedItems.push({
        id: id,
        oldValue: oldValue,
        newValue: currentValue
      });
    }
  });

  if (changedItems.length > 0) {
    sendSalesSupportDiscordAlert_(changedItems);
  }

  // 변경 여부와 상관없이 현재 상태 저장
  // 그래야 G열을 지웠다가 다시 넣는 경우도 상태 꼬임이 덜함
  props.setProperty(
    SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY,
    JSON.stringify(currentMap)
  );
}


/**
 * 영업지원요청 시트에서
 * A열 접수번호와 G열 값을 맵으로 가져옴.
 */
function getSalesSupportCurrentIdGMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SALES_SUPPORT_ALERT_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "시트를 찾을 수 없음: " + SALES_SUPPORT_ALERT_CONFIG.SHEET_NAME
    );
  }

  var startRow = SALES_SUPPORT_ALERT_CONFIG.START_ROW;
  var endRow = SALES_SUPPORT_ALERT_CONFIG.END_ROW;
  var watchCol = SALES_SUPPORT_ALERT_CONFIG.WATCH_COL;
  var idCol = SALES_SUPPORT_ALERT_CONFIG.ID_COL;

  var numRows = endRow - startRow + 1;

  var values = sheet
    .getRange(startRow, 1, numRows, watchCol)
    .getDisplayValues();

  var map = {};

  values.forEach(function(row) {
    var id = String(row[idCol - 1] || "").trim();
    var gValue = String(row[watchCol - 1] || "").trim();

    if (!id) return;

    map[id] = gValue;
  });

  return map;
}


/**
 * Discord로 알림 발송.
 */
function sendSalesSupportDiscordAlert_(items) {
  var webhookUrl = PropertiesService
    .getScriptProperties()
    .getProperty(SALES_SUPPORT_ALERT_CONFIG.WEBHOOK_PROP_KEY);

  if (!webhookUrl) {
    throw new Error(
      "Discord 웹훅 URL이 저장되어 있지 않음. setSalesSupportWebhookUrlOnce_() 먼저 실행해."
    );
  }

  var shownItems = items.slice(0, SALES_SUPPORT_ALERT_CONFIG.MAX_ALERT_ITEMS);

  var lines = shownItems
    .map(function(item) {
      var before = item.oldValue ? item.oldValue : "빈칸";

      return [
        "• 접수번호 " + item.id,
        "  기존: " + before,
        "  신규: " + item.newValue
      ].join("\n");
    })
    .join("\n\n");

  var extraCount = items.length - shownItems.length;
  var extraText = extraCount > 0
    ? "\n\n외 " + extraCount + "건 더 있음. 시트 가서 봐. 디스코드에 장편소설 쓰지 말고."
    : "";

  var content =
    "🔔 새로운 영업지원요청이 들어왔어여 ㅋㅋ\n\n" +
    lines +
    extraText;

  UrlFetchApp.fetch(webhookUrl, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      content: content
    }),
    muteHttpExceptions: true
  });
}


/**
 * 시간 기반 트리거 설치.
 * 1분마다 checkSalesSupportNewValues() 실행.
 *
 * 이미 같은 함수 트리거가 있으면 삭제 후 다시 설치함.
 */
function installSalesSupportAlertTrigger() {
  var targetFunctionName = "checkSalesSupportNewValues";

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === targetFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(targetFunctionName)
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("영업지원요청 Discord 알림 트리거 설치 완료");
}


/**
 * 테스트용.
 * Discord 알림이 실제로 가는지 확인하고 싶을 때 실행.
 */
function testSalesSupportDiscordAlert() {
  sendSalesSupportDiscordAlert_([
    {
      id: "TEST-001",
      oldValue: "빈칸",
      newValue: "테스트 알림입니다"
    }
  ]);
}
