/**
 * CoreClient — 給其他模組（Attendance / 訓練 / 考核 …）唯讀讀取 Core 資料的薄 helper。
 *
 * 使用：把這份檔案複製到其他模組的 GAS project，呼叫 CoreClient.getEmployee('E001') 等。
 * Core Sheet ID 寫死在此（部署到他模組時要記得改）。
 *
 * 回傳的物件 key 是中文欄位名（與 Sheet header 一致），請用 COL 常數存取以避免 typo。
 */

const CORE_SHEET_ID = '1h6u7RPinABlNdElGDXV_yqNDwpdE9IwAv7Bxu_jLAuM';

const CoreClient = {
  getEmployee: function (employeeId) {
    return findRow_(SHEET_NAMES.EMPLOYEES, COL.EMPLOYEES.ID, employeeId);
  },

  getEmployeeByLineId: function (lineUserId) {
    return findRow_(SHEET_NAMES.EMPLOYEES, COL.EMPLOYEES.LINE_USER_ID, lineUserId);
  },

  getStores: function () {
    return getAllRows_(SHEET_NAMES.STORES).filter(function (s) {
      return s[COL.STORES.STATUS] === 'active';
    });
  },

  getStore: function (storeId) {
    return findRow_(SHEET_NAMES.STORES, COL.STORES.ID, storeId);
  },

  checkPermission: function (employeeId, permissionKey) {
    const emp = this.getEmployee(employeeId);
    if (!emp) return false;
    const row = findRowByTwo_(
      SHEET_NAMES.PERMISSIONS,
      COL.PERMISSIONS.ROLE, emp[COL.EMPLOYEES.ROLE],
      COL.PERMISSIONS.PERMISSION_KEY, permissionKey
    );
    if (!row) return false;
    return row[COL.PERMISSIONS.GRANTED] === true ||
      String(row[COL.PERMISSIONS.GRANTED]).toUpperCase() === 'TRUE';
  },

  getSetting: function (key) {
    const row = findRow_(SHEET_NAMES.SETTINGS, COL.SETTINGS.KEY, key);
    return row ? row[COL.SETTINGS.VALUE] : null;
  }
};

function findRow_(sheetName, colName, value) {
  const sheet = SpreadsheetApp.openById(CORE_SHEET_ID).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const colIdx = headers.indexOf(colName);
  if (colIdx < 0) return null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIdx] === value) return zipRow_(headers, data[i]);
  }
  return null;
}

function findRowByTwo_(sheetName, c1, v1, c2, v2) {
  const sheet = SpreadsheetApp.openById(CORE_SHEET_ID).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const i1 = headers.indexOf(c1);
  const i2 = headers.indexOf(c2);
  for (let i = 1; i < data.length; i++) {
    if (data[i][i1] === v1 && data[i][i2] === v2) return zipRow_(headers, data[i]);
  }
  return null;
}

function getAllRows_(sheetName) {
  const sheet = SpreadsheetApp.openById(CORE_SHEET_ID).getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(function (row) { return zipRow_(headers, row); });
}

function zipRow_(headers, row) {
  const obj = {};
  headers.forEach(function (h, i) { obj[h] = row[i]; });
  return obj;
}
