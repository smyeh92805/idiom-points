/**
 * 皇家大衛成語集點站 —— Google 試算表後台（Apps Script）
 *
 * 會自動建立兩個分頁：
 *   「作答紀錄」：每一次交卷都記錄（第一次＝測驗、計點；之後＝練習、不計點）
 *   「點數總表」：每位學生的完成單元數、全對單元數、總點數、最後作答時間（自動更新、依點數排序）
 *
 * TEACHER_CODE 必須和網站 config.js 裡的 teacherCode 一樣。
 */
var TEACHER_CODE = "David2026";
var LOG_SHEET = "作答紀錄";
var SUM_SHEET = "點數總表";
var HEAD = ["time", "name", "classCode", "level", "unitId", "unit", "theme", "score", "total", "points", "mode", "counted"];
var HEAD_ZH = ["作答時間", "姓名", "班級代碼", "年級", "單元代號", "單元", "主題", "答對", "題數", "點數", "類型", "是否計點"];
var SUM_HEAD = ["姓名", "班級代碼", "完成單元", "全對單元", "總點數", "最後作答時間"];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function logSheet_() {
  var sh = ss_().getSheetByName(LOG_SHEET) || ss_().insertSheet(LOG_SHEET);
  if (sh.getLastRow() > 0 && sh.getLastColumn() < HEAD_ZH.length) {
    // 舊版表頭（10 欄）自動補上「類型」「是否計點」
    sh.getRange(1, 1, 1, HEAD_ZH.length).setValues([HEAD_ZH]);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD_ZH);
    sh.getRange(1, 1, 1, HEAD_ZH.length).setFontWeight("bold").setBackground("#1f3f8f").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_() {
  var sh = logSheet_();
  if (sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  return v.map(function (r) {
    var o = {};
    HEAD.forEach(function (h, i) {
      o[h] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], "Asia/Taipei", "yyyy-MM-dd HH:mm") : r[i];
    });
    return o;
  });
}

function updateSummary_() {
  var by = {};
  rows_().forEach(function (r) {
    if (r.counted === "否") return;   // 練習不計點
    var k = r.name;
    if (!by[k]) by[k] = { name: r.name, code: r.classCode, units: 0, perfect: 0, points: 0, last: "" };
    by[k].units++;
    if (+r.score === +r.total) by[k].perfect++;
    by[k].points += (+r.points || 0);
    if (String(r.time) > by[k].last) by[k].last = String(r.time);
  });
  var list = Object.keys(by).map(function (k) { return by[k]; })
    .sort(function (a, b) { return b.points - a.points; });
  var sh = ss_().getSheetByName(SUM_SHEET) || ss_().insertSheet(SUM_SHEET);
  sh.clearContents();
  sh.getRange(1, 1, 1, SUM_HEAD.length).setValues([SUM_HEAD])
    .setFontWeight("bold").setBackground("#d9a82e").setFontColor("#2b2105");
  sh.setFrozenRows(1);
  if (list.length) {
    sh.getRange(2, 1, list.length, SUM_HEAD.length).setValues(list.map(function (s) {
      return [s.name, s.code, s.units, s.perfect, s.points, s.last];
    }));
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 學生交卷時網站送來的成績 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = JSON.parse(e.postData.contents);
    if (r.action !== "submit" || !r.name || !r.unitId) return json_({ ok: false });
    // 以試算表為準：同一學生同一單元已有計點紀錄 → 這次一律算練習、不計點
    var already = rows_().some(function (x) {
      return x.name === r.name && x.unitId === r.unitId && x.counted !== "否";
    });
    r.mode = already ? "練習" : "測驗";
    r.counted = already ? "否" : "是";
    if (already) r.points = 0;
    logSheet_().appendRow(HEAD.map(function (h) { return r[h] == null ? "" : r[h]; }));
    if (!already) updateSummary_();
    return json_({ ok: true, counted: !already });
  } finally {
    lock.releaseLock();
  }
}

/** 學生登入時取回自己的紀錄；老師後台取回全部紀錄 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === "mine" && p.name) {
    return json_({ records: rows_().filter(function (x) { return x.name === p.name && x.counted !== "否"; }) });
  }
  if (p.action === "all" && p.teacher === TEACHER_CODE) {
    return json_({ records: rows_() });
  }
  return json_({ ok: true, msg: "皇家大衛成語集點站 API 運作中" });
}

/** 第一次設定時，在編輯器選這個函式按「執行」，會建好兩個分頁並要求授權 */
function setup() {
  logSheet_();
  updateSummary_();
}
