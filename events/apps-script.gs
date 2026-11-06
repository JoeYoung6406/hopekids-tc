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
 * 改了這份程式要「管理部署作業 → 編輯 → 版本：新版本」才會生效，網址不變。
 */
const SHEET_NAME = '回覆';
const HEADER = ['最後更新', '活動', '姓名', '出席', '請假原因'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    const event = clean(d.event, 60);
    const name = clean(d.name, 30);
    const status = d.status;
    if (!event || !name) return reply({ ok: false, error: '缺少活動或姓名' });
    if (status !== '出席' && status !== '請假') return reply({ ok: false, error: '出席狀態不正確' });
    const reason = status === '請假' ? clean(d.reason, 200) : '';

    const sheet = getSheet();
    const row = [new Date(), event, name, status, reason];
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][1] === event && values[i][2] === name) {
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return reply({ ok: true, updated: true });
      }
    }
    sheet.appendRow(row);
    return reply({ ok: true, updated: false });
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

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  }
  return sheet;
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
