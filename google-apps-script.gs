// 粘贴到 Google 表格的「扩展程序 → Apps 脚本」里，然后部署为网页应用
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["received_at","session_id","nickname","lang","started_at","finished_at",
        "q_index","sticker_id","original_file","options_shown","answer","is_custom","ms","user_agent"]);
    }
    var s = JSON.parse(e.postData.contents);
    var now = new Date();
    var rows = (s.answers || []).map(function (a, i) {
      return [now, s.id, s.nick || "", s.lang || "", s.startedAt, s.finishedAt,
        i + 1, a.sticker, a.orig || "", (a.shown || []).join("/"), a.choice, a.custom ? 1 : 0, a.ms, s.ua || ""];
    });
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
