// 整段替换 Apps 脚本里原来的代码。数据会写进表格里名为 chat 的工作表（没有会自动新建）。
// 以后网页多传了新字段，这个脚本会自动在表头后面加新列，不用再改。
var SHEET_ID = "1htgwkGz6i_287M-QuQPdb1NPY82zWr29NhizGSIn6Q0";
var TAB = "chat";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(TAB) || ss.insertSheet(TAB);
    var s = JSON.parse(e.postData.contents);
    var base = {
      received_at: new Date(), session_id: s.id, nickname: s.nick || "", lang: s.lang || "",
      started_at: s.startedAt, finished_at: s.finishedAt, user_agent: s.ua || ""
    };
    var rows = (s.answers || []).map(function (a, i) {
      var r = {};
      for (var k in base) r[k] = base[k];
      r.q_index = i + 1;
      for (var key in a) {
        var v = a[key];
        r[key] = Array.isArray(v) ? v.join("/") : v;
      }
      return r;
    });
    if (!rows.length) return ContentService.createTextOutput("empty");

    var lastCol = sheet.getLastColumn();
    var header = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    rows.forEach(function (r) {
      for (var k in r) if (header.indexOf(k) < 0) header.push(k);
    });
    sheet.getRange(1, 1, 1, header.length).setValues([header]);

    var values = rows.map(function (r) {
      return header.map(function (h) { return r[h] === undefined ? "" : r[h]; });
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, values.length, header.length).setValues(values);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
