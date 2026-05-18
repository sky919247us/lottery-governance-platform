/**
 * Settings 分頁讀寫 helpers + admin email 通知。
 * 共用 utility，所有模組都可呼叫。
 */

function getSetting_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setSetting_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('Settings 分頁不存在');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

function notifyAdminEmail_(subject, body) {
  try {
    const to = getSetting_('admin_email_for_alerts');
    if (!to) {
      console.warn('admin_email_for_alerts not set, skipping email');
      return;
    }
    MailApp.sendEmail(String(to), subject, body);
  } catch (e) {
    console.warn('notifyAdminEmail_ failed: ' + e.message);
  }
}
