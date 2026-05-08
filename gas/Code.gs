/**
 * 修信念恩日記 — Google Apps Script 後端
 *
 * 部署步驟：
 *   1. 開啟試算表 → 擴充功能 → Apps Script
 *   2. 貼上此程式碼
 *   3. 專案設定 → 指令碼屬性 → 新增：
 *        API_KEY = （自訂一組密碼，與前端一致）
 *   4. 部署 → 新增部署 → 網路應用程式
 *        執行身分：我
 *        存取權：任何人
 *   5. 複製部署網址，貼到前端 HTML 的 GAS_URL
 *
 * 每位使用者有自己的工作表（以使用者名稱命名）
 * 欄位：日期 | 時間 | 聽聞前行 | 聽修信念恩 | 寫日記 | 字數 | 聽聞段落 | 念恩段落 | 日記內容
 */

var TZ = "Asia/Taipei";
var HEADER_ROW = 2;
var DATA_START = 3;
var HEADERS = ["日期", "時間", "聽聞前行", "聽修信念恩", "寫日記", "字數", "聽聞段落", "念恩段落", "聽聞段落內容", "念恩段落內容", "日記內容"];
var COL_WIDTHS = [90, 70, 75, 90, 65, 60, 80, 80, 250, 250, 500];

/* ── Entry Points ── */

function doGet(e) {
  if (e && e.parameter && e.parameter.ping === "1") {
    return out_({ ok: true, data: { pong: true } });
  }
  return ContentService.createTextOutput("修信念恩日記 API")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) return out_({ ok: false, error: "empty body" });
    var body = JSON.parse(e.postData.contents);
    if (!verifyKey_(body)) return out_({ ok: false, error: "unauthorized" });

    switch (body.action) {
      case "listUsers":         return out_({ ok: true, data: { users: listUsers_() } });
      case "ensureUser":        return out_(ensureUser_(body.userName, body.code));
      case "loginUser":         return out_(loginUser_(body.userName, body.code));
      case "upgradeLegacyUser": return out_(upgradeLegacyUser_(body.userName));
      case "saveEntry":         return out_(saveEntry_(body));
      case "getTodayEntry":     return out_(getTodayEntry_(body.userName, body.code));
      case "getEntries":        return out_(getEntries_(body.userName, body.code, body.limit));
      case "getStats":          return out_(getStats_(body.userName, body.code));
      case "getHomeData":       return out_(getHomeData_(body.userName, body.code, body.limit));
      default:                  return out_({ ok: false, error: "unknown action: " + body.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyKey_(body) {
  var key = PropertiesService.getScriptProperties().getProperty("API_KEY");
  if (!key) return false;
  return body.apiKey === key;
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ── Users (= Sheet Tabs) ── */
/*
 * 工作表命名規則：
 *   新版：「{name}-{code}」，code 為 4 碼大寫 hex（例：小明-A3F2）
 *   舊版（legacy）：「{name}」，無短碼（為相容既有使用者保留）
 *   系統表：以「_」開頭（不顯示給使用者）
 */

var CODE_LEN = 4;

function listUsers_() {
  var sheets = ss_().getSheets();
  var users = [];
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.indexOf("_") === 0) continue;
    var p = parseSheetName_(n);
    users.push({ name: p.name, code: p.code, full: n, isLegacy: !p.code });
  }
  return users;
}

/**
 * 把時間欄位讀回的值統一成 "HH:mm" 字串。
 * Sheet 可能把 "11:29" 自動轉成 Date 物件（epoch 1899-12-30），讀回會是長字串，
 * 此處統一處理新舊兩種情況。
 */
function formatTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "HH:mm");
  return String(v == null ? "" : v).trim();
}

function sanitizeName_(name) {
  if (!name || typeof name !== "string") throw new Error("invalid userName");
  var s = name.trim();
  if (!s) throw new Error("empty userName");
  if (/[:\\/?*[\]]/.test(s)) throw new Error("名稱含有不允許的字元");
  if (s.length > 30) throw new Error("名稱太長（最多 30 字）");
  return s;
}

function sanitizeCode_(code) {
  if (code == null || code === "") return "";
  var s = String(code).trim().toUpperCase();
  if (!/^[0-9A-F]+$/.test(s)) throw new Error("識別碼格式不正確");
  if (s.length < 3 || s.length > 8) throw new Error("識別碼長度不正確");
  return s;
}

function genCode_() {
  // 4 hex chars, uppercase. Avoids ambiguity by using crypto-quality randomness.
  var hex = Utilities.getUuid().replace(/-/g, "").toUpperCase();
  return hex.substring(0, CODE_LEN);
}

function parseSheetName_(sheetName) {
  // 「name-CODE」→ {name, code}；無 dash 或尾段非 hex → legacy {name, code:""}
  var m = /^(.+)-([0-9A-F]{3,8})$/.exec(sheetName);
  if (m) return { name: m[1], code: m[2] };
  return { name: sheetName, code: "" };
}

function formatSheetName_(name, code) {
  return code ? (name + "-" + code) : name;
}

/**
 * 找到使用者的工作表。
 *   - 有 code：只找「name-code」（不會 fallback 到 legacy）
 *   - 沒 code：找純「name」（legacy 帳號）
 */
function findUserSheet_(name, code) {
  return ss_().getSheetByName(formatSheetName_(name, code));
}

var DEFAULT_ROWS = 2000;

function initSheet_(sheet, displayName) {
  sheet.getRange(1, 1, 1, HEADERS.length).merge();
  sheet.getRange(1, 1).setValue(displayName + " 的修信念恩日記")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(12);
  for (var i = 0; i < HEADERS.length; i++) {
    sheet.getRange(HEADER_ROW, i + 1).setValue(HEADERS[i]).setFontWeight("bold");
    sheet.setColumnWidth(i + 1, COL_WIDTHS[i]);
  }
  sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setBackground("#f0e6d8");
  sheet.setFrozenRows(HEADER_ROW);

  // 預設 2000 列（Google Sheet 新增工作表預設 1000 列）
  var maxRows = sheet.getMaxRows();
  if (maxRows < DEFAULT_ROWS) {
    sheet.insertRowsAfter(maxRows, DEFAULT_ROWS - maxRows);
  } else if (maxRows > DEFAULT_ROWS) {
    sheet.deleteRows(DEFAULT_ROWS + 1, maxRows - DEFAULT_ROWS);
  }

  // 預設隱藏新工作表（避免共用試算表時其他人直接看到）
  sheet.hideSheet();
}

/**
 * 建立或確認使用者。
 *   - 有 code：找「name-code」，找不到就建立。用於既有裝置回傳已知短碼。
 *   - 沒 code：建立全新使用者（自動產生短碼，避免同名衝突）。
 */
function ensureUser_(userName, code) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);

  if (c) {
    var existing = ss_().getSheetByName(formatSheetName_(name, c));
    if (existing) return { ok: true, data: { userName: name, code: c, created: false } };
    var s1 = ss_().insertSheet(formatSheetName_(name, c));
    initSheet_(s1, name);
    SpreadsheetApp.flush();
    return { ok: true, data: { userName: name, code: c, created: true } };
  }

  // 新使用者：自動產生不重複的短碼
  var newCode = genCode_();
  for (var tries = 0; tries < 10 && ss_().getSheetByName(formatSheetName_(name, newCode)); tries++) {
    newCode = genCode_();
  }
  var s2 = ss_().insertSheet(formatSheetName_(name, newCode));
  initSheet_(s2, name);
  SpreadsheetApp.flush();
  return { ok: true, data: { userName: name, code: newCode, created: true } };
}

/**
 * 從其他裝置登入。
 *   - 有 code：驗證「name-code」是否存在
 *   - 沒 code：驗證 legacy 純「name」是否存在
 */
function loginUser_(userName, code) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);
  var sheet = findUserSheet_(name, c);
  if (!sheet) return { ok: true, data: { found: false } };
  return { ok: true, data: { found: true, userName: name, code: c, isLegacy: !c } };
}

/**
 * 升級 legacy 帳號：把純「name」的工作表 rename 為「name-code」並回傳新短碼。
 */
function upgradeLegacyUser_(userName) {
  var name = sanitizeName_(userName);
  var sheet = ss_().getSheetByName(name);
  if (!sheet) return { ok: false, error: "找不到舊版帳號：" + name };

  var newCode = genCode_();
  for (var tries = 0; tries < 10 && ss_().getSheetByName(formatSheetName_(name, newCode)); tries++) {
    newCode = genCode_();
  }
  sheet.setName(formatSheetName_(name, newCode));
  SpreadsheetApp.flush();
  return { ok: true, data: { userName: name, code: newCode } };
}

/* ── Save Entry ── */

function saveEntry_(body) {
  var name = sanitizeName_(body.userName);
  var c = sanitizeCode_(body.code);
  var sheet = findUserSheet_(name, c);
  if (!sheet) throw new Error("找不到工作表：" + formatSheetName_(name, c));

  var now = new Date();
  var dateStr = Utilities.formatDate(now, TZ, "yyyy/MM/dd");
  var timeStr = Utilities.formatDate(now, TZ, "HH:mm");
  var didHearing = body.didHearing ? "V" : "";
  var didFaith   = body.didFaith   ? "V" : "";
  var didJournal = body.didJournal ? "V" : "";
  var text       = body.journalText || "";
  var charCount  = text.length;
  var hearingId    = body.hearingId || "";
  var faithId      = body.faithId || "";
  var hearingQuote = body.hearingQuote || "";
  var faithQuote   = body.faithQuote || "";

  // Check if today already has an entry — overwrite if so
  var lastRow = sheet.getLastRow();
  var targetRow = -1;
  if (lastRow >= DATA_START) {
    var dates = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 1).getValues();
    for (var i = dates.length - 1; i >= 0; i--) {
      if (String(dates[i][0]).trim() === dateStr) {
        targetRow = DATA_START + i;
        break;
      }
    }
  }
  if (targetRow === -1) {
    targetRow = lastRow < DATA_START ? DATA_START : lastRow + 1;
  }

  var rowData = [dateStr, timeStr, didHearing, didFaith, didJournal,
                 charCount, hearingId, faithId, hearingQuote, faithQuote, text];
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  sheet.getRange(targetRow, 1).setNumberFormat("@");
  sheet.getRange(targetRow, 2).setNumberFormat("@"); // 時間欄強制文字，避免被 Sheet 自動轉成時間值
  sheet.getRange(targetRow, 6).setNumberFormat("0");
  sheet.getRange(targetRow, 3, 1, 3).setHorizontalAlignment("center");

  SpreadsheetApp.flush();
  return { ok: true, data: { row: targetRow, overwritten: targetRow <= lastRow } };
}

/* ── Get Today's Entry ── */

function getTodayEntry_(userName, code) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);
  var sheet = findUserSheet_(name, c);
  if (!sheet) return { ok: true, data: { found: false } };

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { ok: true, data: { found: false } };

  var todayStr = Utilities.formatDate(new Date(), TZ, "yyyy/MM/dd");
  var data = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, HEADERS.length).getValues();

  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === todayStr) {
      return { ok: true, data: {
        found: true,
        date: todayStr,
        time: formatTime_(data[i][1]),
        didHearing: String(data[i][2]).trim() === "V",
        didFaith: String(data[i][3]).trim() === "V",
        didJournal: String(data[i][4]).trim() === "V",
        charCount: parseInt(data[i][5], 10) || 0,
        hearingId: String(data[i][6]).trim(),
        faithId: String(data[i][7]).trim(),
        hearingQuote: String(data[i][8] || "").trim(),
        faithQuote: String(data[i][9] || "").trim(),
        text: String(data[i][10] || "").trim()
      }};
    }
  }
  return { ok: true, data: { found: false } };
}

/* ── Get Entries ── */

function getEntries_(userName, code, limit) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);
  var sheet = findUserSheet_(name, c);
  if (!sheet) return { ok: true, data: { entries: [] } };

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { ok: true, data: { entries: [] } };

  var lim = limit || 30;
  var data = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, HEADERS.length).getValues();
  var entries = [];

  for (var i = data.length - 1; i >= 0 && entries.length < lim; i--) {
    var r = data[i];
    entries.push({
      date:         String(r[0]).trim(),
      time:         formatTime_(r[1]),
      didHearing:   String(r[2]).trim() === "V",
      didFaith:     String(r[3]).trim() === "V",
      didJournal:   String(r[4]).trim() === "V",
      charCount:    parseInt(r[5], 10) || 0,
      hearingId:    String(r[6]).trim(),
      faithId:      String(r[7]).trim(),
      hearingQuote: String(r[8] || "").trim(),
      faithQuote:   String(r[9] || "").trim(),
      text:         String(r[10] || "").trim()
    });
  }
  return { ok: true, data: { entries: entries } };
}

/* ── Stats ── */

function getStats_(userName, code) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);
  var sheet = findUserSheet_(name, c);
  if (!sheet) return { ok: true, data: emptyStats_() };

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { ok: true, data: emptyStats_() };

  var data = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 6).getValues();

  var totalEntries = 0;
  var totalChars = 0;
  var hearingCount = 0;
  var faithCount = 0;
  var journalCount = 0;
  var dateSet = {};        // unique dates
  var recentDates = {};    // last 30 days
  var weekDates = {};      // this week
  var monthDates = {};     // this month

  var now = new Date();
  var todayStr = Utilities.formatDate(now, TZ, "yyyy/MM/dd");
  var thirtyDaysAgo = new Date(now.getTime() - 30 * 86400000);
  var thirtyStr = Utilities.formatDate(thirtyDaysAgo, TZ, "yyyy/MM/dd");

  // This week (Monday start)
  var dayOfWeek = now.getDay(); // 0=Sun
  var mondayOffset = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  var monday = new Date(now.getTime() - mondayOffset * 86400000);
  var mondayStr = Utilities.formatDate(monday, TZ, "yyyy/MM/dd");

  // This month
  var monthStart = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), TZ, "yyyy/MM/dd");

  for (var i = 0; i < data.length; i++) {
    var dateVal = String(data[i][0]).trim();
    if (!dateVal) continue;
    totalEntries++;

    var chars = parseInt(data[i][5], 10) || 0;
    totalChars += chars;

    if (String(data[i][2]).trim() === "V") hearingCount++;
    if (String(data[i][3]).trim() === "V") faithCount++;
    if (String(data[i][4]).trim() === "V") journalCount++;

    dateSet[dateVal] = true;
    if (dateVal >= thirtyStr)  recentDates[dateVal] = true;
    if (dateVal >= mondayStr)  weekDates[dateVal] = true;
    if (dateVal >= monthStart) monthDates[dateVal] = true;
  }

  var uniqueDays = Object.keys(dateSet).length;

  // Streak: count consecutive days ending at today (or yesterday)
  var streak = 0;
  var checkDate = new Date(now.getTime());
  // If no entry today, start from yesterday
  if (!dateSet[todayStr]) {
    checkDate = new Date(now.getTime() - 86400000);
  }
  for (var s = 0; s < 365; s++) {
    var ds = Utilities.formatDate(checkDate, TZ, "yyyy/MM/dd");
    if (dateSet[ds]) { streak++; checkDate = new Date(checkDate.getTime() - 86400000); }
    else break;
  }

  return {
    ok: true,
    data: {
      totalEntries: totalEntries,
      uniqueDays: uniqueDays,
      totalChars: totalChars,
      avgChars: uniqueDays > 0 ? Math.round(totalChars / uniqueDays) : 0,
      hearingCount: hearingCount,
      faithCount: faithCount,
      journalCount: journalCount,
      streak: streak,
      weekDays: Object.keys(weekDates).length,
      monthDays: Object.keys(monthDates).length,
      recentDays: Object.keys(recentDates).length
    }
  };
}

function emptyStats_() {
  return {
    totalEntries: 0, uniqueDays: 0, totalChars: 0, avgChars: 0,
    hearingCount: 0, faithCount: 0, journalCount: 0,
    streak: 0, weekDays: 0, monthDays: 0, recentDays: 0
  };
}

/* ── Home data (today + entries + stats in one sheet read) ── */

function getHomeData_(userName, code, limit) {
  var name = sanitizeName_(userName);
  var c = sanitizeCode_(code);
  var sheet = findUserSheet_(name, c);
  var lim = limit || 30;
  if (!sheet) {
    return { ok: true, data: { today: { found: false }, entries: [], stats: emptyStats_() } };
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) {
    return { ok: true, data: { today: { found: false }, entries: [], stats: emptyStats_() } };
  }

  var data = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, HEADERS.length).getValues();

  var todayStr = Utilities.formatDate(new Date(), TZ, "yyyy/MM/dd");
  var today = { found: false };
  var entries = [];

  for (var i = data.length - 1; i >= 0; i--) {
    var r = data[i];
    var dateVal = String(r[0]).trim();
    if (!dateVal) continue;
    var rec = {
      date: dateVal,
      time: formatTime_(r[1]),
      didHearing: String(r[2]).trim() === "V",
      didFaith: String(r[3]).trim() === "V",
      didJournal: String(r[4]).trim() === "V",
      charCount: parseInt(r[5], 10) || 0,
      hearingId: String(r[6]).trim(),
      faithId: String(r[7]).trim(),
      hearingQuote: String(r[8] || "").trim(),
      faithQuote: String(r[9] || "").trim(),
      text: String(r[10] || "").trim()
    };
    if (!today.found && dateVal === todayStr) {
      today = {
        found: true,
        date: rec.date, time: rec.time,
        didHearing: rec.didHearing, didFaith: rec.didFaith, didJournal: rec.didJournal,
        charCount: rec.charCount,
        hearingId: rec.hearingId, faithId: rec.faithId,
        hearingQuote: rec.hearingQuote, faithQuote: rec.faithQuote,
        text: rec.text
      };
    }
    if (entries.length < lim) entries.push(rec);
  }

  var stats = computeStats_(data);
  return { ok: true, data: { today: today, entries: entries, stats: stats } };
}

function computeStats_(data) {
  var totalEntries = 0, totalChars = 0;
  var hearingCount = 0, faithCount = 0, journalCount = 0;
  var dateSet = {}, recentDates = {}, weekDates = {}, monthDates = {};

  var now = new Date();
  var todayStr = Utilities.formatDate(now, TZ, "yyyy/MM/dd");
  var thirtyStr = Utilities.formatDate(new Date(now.getTime() - 30 * 86400000), TZ, "yyyy/MM/dd");
  var dayOfWeek = now.getDay();
  var mondayOffset = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  var mondayStr = Utilities.formatDate(new Date(now.getTime() - mondayOffset * 86400000), TZ, "yyyy/MM/dd");
  var monthStart = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), TZ, "yyyy/MM/dd");

  for (var i = 0; i < data.length; i++) {
    var dateVal = String(data[i][0]).trim();
    if (!dateVal) continue;
    totalEntries++;
    totalChars += parseInt(data[i][5], 10) || 0;
    if (String(data[i][2]).trim() === "V") hearingCount++;
    if (String(data[i][3]).trim() === "V") faithCount++;
    if (String(data[i][4]).trim() === "V") journalCount++;
    dateSet[dateVal] = true;
    if (dateVal >= thirtyStr)  recentDates[dateVal] = true;
    if (dateVal >= mondayStr)  weekDates[dateVal] = true;
    if (dateVal >= monthStart) monthDates[dateVal] = true;
  }

  var uniqueDays = Object.keys(dateSet).length;
  var streak = 0;
  var checkDate = dateSet[todayStr] ? new Date(now.getTime()) : new Date(now.getTime() - 86400000);
  for (var s = 0; s < 365; s++) {
    var ds = Utilities.formatDate(checkDate, TZ, "yyyy/MM/dd");
    if (dateSet[ds]) { streak++; checkDate = new Date(checkDate.getTime() - 86400000); }
    else break;
  }

  return {
    totalEntries: totalEntries,
    uniqueDays: uniqueDays,
    totalChars: totalChars,
    avgChars: uniqueDays > 0 ? Math.round(totalChars / uniqueDays) : 0,
    hearingCount: hearingCount,
    faithCount: faithCount,
    journalCount: journalCount,
    streak: streak,
    weekDays: Object.keys(weekDates).length,
    monthDays: Object.keys(monthDates).length,
    recentDays: Object.keys(recentDates).length
  };
}
