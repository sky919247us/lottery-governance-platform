/**
 * 一次性綁定 token 管理
 *
 * 流程：
 *   1. Admin 從 Apps Script 編輯器跑 GenerateBindUrlForE001（或 generateBindUrl_(id)）
 *   2. 系統產生 24 小時內有效的綁定連結
 *   3. Admin 把連結貼給員工（LINE / 簡訊 / Email）
 *   4. 員工點開 → LINE 登入 → 系統綁定 line_user_id 到該員工
 *
 * BindTokens 分頁：token | employee_id | created_at | expires_at | used_at | used_by_line_id
 */

/** 為 E001（admin 自己）產生綁定連結 — 快速入口 */
function GenerateBindUrlForE001() {
  const url = generateBindUrl_('E001');
  const html =
    '<div style="font-family:sans-serif">' +
    '<p>綁定連結（24 小時有效）：</p>' +
    '<input type="text" value="' + url + '" style="width:100%;font-size:12px;padding:6px" readonly onclick="this.select()">' +
    '<p style="margin-top:16px">請複製此連結到瀏覽器開啟，並用 LINE 登入完成綁定。</p>' +
    '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(640).setHeight(220),
    '綁定連結 — E001'
  );
}

/** 通用：為指定員工產生綁定連結，回傳 URL */
function generateBindUrl_(employeeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  const empData = empSheet.getDataRange().getValues();
  const headers = empData[0];
  const idCol = headers.indexOf('employee_id');
  const lineCol = headers.indexOf('line_user_id');

  let foundIdx = -1;
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][idCol] === employeeId) { foundIdx = i; break; }
  }
  if (foundIdx < 0) throw new Error('員工 ' + employeeId + ' 不存在');
  if (empData[foundIdx][lineCol]) {
    throw new Error('員工 ' + employeeId + ' 已綁定 LINE。如需重綁，請先在 Employees 表清空該員工的 line_user_id 欄位。');
  }

  const tokenSheet = ss.getSheetByName(SHEET_NAMES.BIND_TOKENS);
  if (!tokenSheet) throw new Error('BindTokens 分頁不存在，請先執行 MigrateAddBindTokens');

  const token = Utilities.getUuid().replace(/-/g, '');
  const now = new Date();
  const expires = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  tokenSheet.appendRow([token, employeeId, now, expires, '', '']);

  return getWebAppUrl_() + '?action=bind&token=' + token;
}

/** 查詢 token，回傳完整資訊或 null */
function lookupBindToken_(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.BIND_TOKENS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
      const empData = empSheet.getDataRange().getValues();
      let empName = '';
      for (let j = 1; j < empData.length; j++) {
        if (empData[j][0] === data[i][1]) { empName = empData[j][1]; break; }
      }
      return {
        employee_id: data[i][1],
        employee_name: empName,
        created_at: data[i][2],
        expires_at: data[i][3],
        used_at: data[i][4],
        used_by_line_id: data[i][5]
      };
    }
  }
  return null;
}

/** 標記 token 已使用 */
function markBindTokenUsed_(token, lineUserId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BIND_TOKENS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      sheet.getRange(i + 1, 5).setValue(new Date());
      sheet.getRange(i + 1, 6).setValue(lineUserId);
      return;
    }
  }
}

/** Migration：手動執行一次，把 BindTokens 分頁加進已存在的 Sheet */
function MigrateAddBindTokens() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET_NAMES.BIND_TOKENS)) {
    SpreadsheetApp.getUi().alert('BindTokens 分頁已存在，無需 migration。');
    return;
  }
  const sheet = ss.insertSheet(SHEET_NAMES.BIND_TOKENS);
  const headers = ['token', 'employee_id', 'created_at', 'expires_at', 'used_at', 'used_by_line_id'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert('BindTokens 分頁已建立 ✓');
}
