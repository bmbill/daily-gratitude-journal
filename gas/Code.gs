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
      case "listUsers":    return out_({ ok: true, data: { users: listUsers_() } });
      case "ensureUser":   return out_(ensureUser_(body.userName));
      case "saveEntry":    return out_(saveEntry_(body));
      case "getTodayEntry":return out_(getTodayEntry_(body.userName));
      case "getEntries":   return out_(getEntries_(body.userName, body.limit));
      case "getStats":     return out_(getStats_(body.userName));
      default:             return out_({ ok: false, error: "unknown action: " + body.action });
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

function listUsers_() {
  var sheets = ss_().getSheets();
  var names = [];
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.indexOf("_") !== 0) names.push(n);   // skip _prefixed system sheets
  }
  return names;
}

function sanitizeName_(name) {
  if (!name || typeof name !== "string") throw new Error("invalid userName");
  var s = name.trim();
  if (!s) throw new Error("empty userName");
  if (/[:\\/?*[\]]/.test(s)) throw new Error("名稱含有不允許的字元");
  if (s.length > 30) throw new Error("名稱太長（最多 30 字）");
  return s;
}

function ensureUser_(userName) {
  var name = sanitizeName_(userName);
  var sheet = ss_().getSheetByName(name);
  if (sheet) return { ok: true, data: { userName: name, created: false } };

  sheet = ss_().insertSheet(name);
  // Title row
  sheet.getRange(1, 1, 1, HEADERS.length).merge();
  sheet.getRange(1, 1).setValue(name + " 的修信念恩日記")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontSize(12);

  // Header row
  for (var i = 0; i < HEADERS.length; i++) {
    sheet.getRange(HEADER_ROW, i + 1).setValue(HEADERS[i]).setFontWeight("bold");
    sheet.setColumnWidth(i + 1, COL_WIDTHS[i]);
  }
  sheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setBackground("#f0e6d8");
  sheet.setFrozenRows(HEADER_ROW);

  SpreadsheetApp.flush();
  return { ok: true, data: { userName: name, created: true } };
}

/* ── Save Entry ── */

function saveEntry_(body) {
  var name = sanitizeName_(body.userName);
  var sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error("找不到工作表：" + name);

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
  sheet.getRange(targetRow, 6).setNumberFormat("0");
  sheet.getRange(targetRow, 3, 1, 3).setHorizontalAlignment("center");

  SpreadsheetApp.flush();
  return { ok: true, data: { row: targetRow, overwritten: targetRow <= lastRow } };
}

/* ── Get Today's Entry ── */

function getTodayEntry_(userName) {
  var name = sanitizeName_(userName);
  var sheet = ss_().getSheetByName(name);
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
        time: String(data[i][1]).trim(),
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

function getEntries_(userName, limit) {
  var name = sanitizeName_(userName);
  var sheet = ss_().getSheetByName(name);
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
      time:         String(r[1]).trim(),
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

function getStats_(userName) {
  var name = sanitizeName_(userName);
  var sheet = ss_().getSheetByName(name);
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
