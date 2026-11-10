/**
 * Hope Kids 活動回覆：網站「活動專區」問卷的後台（Google Apps Script）
 *
 * 設定方式（只要做一次，之後每個活動共用）：
 *   1. 建一個 Google 試算表 → 擴充功能 → Apps Script → 把這整段貼上 → 儲存
 *   2. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *      執行身分：我　／　誰可以存取：所有人 → 部署 → 授權
 *   3. 把「網頁應用程式網址」填進各活動頁的 CONFIG.scriptUrl
 *
 * 每筆回覆寫進「回覆」分頁；同一個活動、同一個名字再送一次，會更新原本那一列。
 * 一次可以回覆好幾個活動（例如下午備課日＋晚上聚集），每個活動各寫一列。
 * 「統計」分頁用公式把各活動分開算人數，回覆一進來就自動更新。
 * 改了這份程式要「管理部署作業 → 編輯 → 版本：新版本」才會生效，網址不變。
 */
const SHEET_NAME = '回覆';
const STATS_NAME = '統計';
const HEADER = ['最後更新', '活動', '姓名', '出席', '請假原因', '人數', '備註'];
const STATUSES = ['出席', '請假', '參加', '不參加'];
const AWAY = ['請假', '不參加'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    // 新的活動頁送 { records: [...] }；舊的活動頁只送一筆，照樣收
    const records = Array.isArray(d.records) ? d.records : [d];
    if (!records.length || records.length > 5) return reply({ ok: false, error: '回覆筆數不正確' });
    const rows = records.map(toRow);

    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    rows.forEach((row) => {
      const i = values.findIndex((v, k) => k > 0 && v[1] === row[1] && v[2] === row[2]);
      if (i > 0) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      } else {
        sheet.appendRow(row);
        values.push(row);
      }
    });
    ensureStats();
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器直接打開網址會看到這行，代表部署成功
function doGet() {
  return ContentService.createTextOutput('Hope Kids 活動回覆：運作中');
}

// 檢查一筆回覆並轉成試算表的一列；請假、不參加算 0 人
function toRow(r) {
  const event = clean(r.event, 60);
  const name = clean(r.name, 30);
  const status = r.status;
  if (!event || !name) throw new Error('缺少活動或姓名');
  if (STATUSES.indexOf(status) < 0) throw new Error('出席狀態不正確');
  const away = AWAY.indexOf(status) >= 0;
  const people = away ? 0 : Math.min(Math.max(parseInt(r.people, 10) || 1, 1), 20);
  const reason = away ? clean(r.reason, 200) : '';
  const note = away ? '' : clean(r.note, 100);
  return [new Date(), event, name, status, reason, people, note];
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  }
  // 舊版只有前 5 欄：補上新欄位名稱，舊回覆的人數照出席狀態補 1 或 0
  if (sheet.getRange(1, HEADER.length).getValue() !== HEADER[HEADER.length - 1]) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    const n = sheet.getLastRow() - 1;
    if (n > 0) {
      const statuses = sheet.getRange(2, 4, n, 1).getValues();
      sheet.getRange(2, 6, n, 1).setValues(statuses.map(([s]) => [AWAY.indexOf(s) >= 0 ? 0 : 1]));
    }
  }
  return sheet;
}

// 「統計」分頁：左邊是各活動分開算的人數，右邊是聚集名單（誰來、幾個人、帶什麼菜）
function ensureStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(STATS_NAME)) return;
  const sheet = ss.insertSheet(STATS_NAME, 1);
  sheet.getRange('A1').setValue('各活動人數（自動計算）').setFontWeight('bold');
  sheet.getRange('A2').setFormula(
    "=IFERROR(QUERY('回覆'!A:G,\"select B, D, count(C), sum(F) where B is not null group by B, D order by B, D " +
    "label B '活動', D '回覆', count(C) '回覆數', sum(F) '總人數'\",1),\"還沒有回覆\")");
  sheet.getRange('G1').setValue('聚集名單（人數與帶的菜）').setFontWeight('bold');
  sheet.getRange('G2').setFormula(
    "=IFERROR(QUERY('回覆'!A:G,\"select B, C, F, G where D = '參加' order by B, A " +
    "label B '活動', C '姓名', F '人數', G '帶的菜'\",1),\"還沒有人報名\")");
  sheet.setFrozenRows(2);
}

// 去頭尾空白、限制長度；開頭是 = + - @ 的前面加 ' ，避免被試算表當成公式
function clean(value, max) {
  const s = String(value == null ? '' : value).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
