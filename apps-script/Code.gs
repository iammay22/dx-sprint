/**
 * dx-sprint 進度同步後端
 * 分頁 progress：A 欄同步代碼、B 欄進度 JSON、C 欄最後更新時間
 */
const SHEET = 'progress';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['同步代碼', '進度', '最後更新']);
  }
  return sh;
}

function findRow_(sh, code) {
  const hit = sh.getRange('A:A').createTextFinder(code).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const code = e && e.parameter && e.parameter.code;
  if (!code) return ContentService.createTextOutput('alive');
  const sh = sheet_();
  const row = findRow_(sh, code);
  if (!row) return json_({ st: {} });
  return json_({ st: JSON.parse(sh.getRange(row, 2).getValue() || '{}') });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  if (!d.code || String(d.code).length < 6) return json_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const row = findRow_(sh, d.code);
    const old = row ? JSON.parse(sh.getRange(row, 2).getValue() || '{}') : {};
    const merged = Object.assign({}, old);
    Object.keys(d.st || {}).forEach(function (k) {
      if (!merged[k] || (d.st[k] && d.st[k].t > merged[k].t)) merged[k] = d.st[k];
    });
    const values = [[String(d.code), JSON.stringify(merged), new Date()]];
    if (row) sh.getRange(row, 1, 1, 3).setValues(values);
    else sh.appendRow(values[0]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
