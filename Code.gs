/**
 * 皇家大衛成語集點站 —— Google Apps Script（成績紀錄表）
 * 部署方式見「老師設定說明.md」。
 * TEACHER_CODE 必須和 config.js 的 teacherCode 一樣。
 */
var TEACHER_CODE = "TEACHER2026";
var SHEET_NAME = "成語集點紀錄";
var HEAD = ["time", "name", "classCode", "level", "unitId", "unit", "theme", "score", "total", "points"];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}
function rows_() {
  var v = sheet_().getDataRange().getValues();
  var head = v.shift();
  return v.map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], "Asia/Taipei", "yyyy-MM-dd HH:mm") : r[i]; });
    return o;
  });
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var r = JSON.parse(e.postData.contents);
    if (r.action !== "submit" || !r.name || !r.unitId) return json_({ ok: false });
    // 每位學生每單元只記第一次
    var dup = rows_().some(function (x) { return x.name === r.name && x.unitId === r.unitId; });
    if (!dup) sheet_().appendRow(HEAD.map(function (h) { return r[h] == null ? "" : r[h]; }));
    return json_({ ok: true, duplicate: dup });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = e.parameter || {};
  if (p.action === "mine" && p.name) {
    return json_({ records: rows_().filter(function (x) { return x.name === p.name; }) });
  }
  if (p.action === "all" && p.teacher === TEACHER_CODE) {
    return json_({ records: rows_() });
  }
  return json_({ ok: true, msg: "皇家大衛成語集點站 API" });
}
