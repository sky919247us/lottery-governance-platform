/**
 * Bootstrap — 建立 Core Sheet 所有分頁 + 預設資料
 *
 * 從 Apps Script 編輯器手動執行一次。重複執行安全（已存在分頁不會覆寫）。
 */
function Bootstrap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEET_NAMES.EMPLOYEES, [
    COL.EMPLOYEES.ID, COL.EMPLOYEES.NAME, COL.EMPLOYEES.LINE_USER_ID,
    COL.EMPLOYEES.HOME_STORE, COL.EMPLOYEES.ROLE,
    COL.EMPLOYEES.EMPLOYMENT_TYPE, COL.EMPLOYEES.MONTHLY_SALARY, COL.EMPLOYEES.HOURLY_RATE,
    COL.EMPLOYEES.SALARY_EFFECTIVE_DATE,
    COL.EMPLOYEES.HIRE_DATE, COL.EMPLOYEES.PROBATION_END_DATE, COL.EMPLOYEES.PROBATION_STATUS,
    COL.EMPLOYEES.STATUS, COL.EMPLOYEES.RESIGN_DATE,
    COL.EMPLOYEES.EMERGENCY_CONTACT, COL.EMPLOYEES.EMERGENCY_PHONE,
    COL.EMPLOYEES.CREATED_AT, COL.EMPLOYEES.UPDATED_AT
  ]);

  ensureSheet_(ss, SHEET_NAMES.EMPLOYEES_SENSITIVE, [
    COL.EMPLOYEES_SENSITIVE.ID, COL.EMPLOYEES_SENSITIVE.ID_NUMBER,
    COL.EMPLOYEES_SENSITIVE.BANK_ACCOUNT,
    COL.EMPLOYEES_SENSITIVE.UPDATED_BY, COL.EMPLOYEES_SENSITIVE.UPDATED_AT
  ]);

  ensureSheet_(ss, SHEET_NAMES.STORES, [
    COL.STORES.ID, COL.STORES.NAME, COL.STORES.ADDRESS,
    COL.STORES.LAT, COL.STORES.LNG, COL.STORES.RADIUS_M,
    COL.STORES.MORNING_START, COL.STORES.MORNING_END,
    COL.STORES.EVENING_START, COL.STORES.EVENING_END,
    COL.STORES.STATUS
  ]);

  ensureSheet_(ss, SHEET_NAMES.ROLES, [COL.ROLES.ROLE, COL.ROLES.DESCRIPTION]);

  ensureSheet_(ss, SHEET_NAMES.PERMISSIONS, [
    COL.PERMISSIONS.ROLE, COL.PERMISSIONS.PERMISSION_KEY, COL.PERMISSIONS.GRANTED
  ]);

  ensureSheet_(ss, SHEET_NAMES.SALARY_HISTORY, [
    COL.SALARY_HISTORY.EMPLOYEE_ID, COL.SALARY_HISTORY.FIELD,
    COL.SALARY_HISTORY.OLD_VALUE, COL.SALARY_HISTORY.NEW_VALUE,
    COL.SALARY_HISTORY.EFFECTIVE_DATE, COL.SALARY_HISTORY.MODIFIED_BY,
    COL.SALARY_HISTORY.MODIFIED_AT, COL.SALARY_HISTORY.REASON
  ]);

  ensureSheet_(ss, SHEET_NAMES.SETTINGS, [
    COL.SETTINGS.KEY, COL.SETTINGS.VALUE, COL.SETTINGS.DESCRIPTION
  ]);

  ensureSheet_(ss, SHEET_NAMES.BACKUP_LOG, [
    COL.BACKUP_LOG.BACKUP_AT, COL.BACKUP_LOG.FILE_ID,
    COL.BACKUP_LOG.STATUS, COL.BACKUP_LOG.NOTE
  ]);

  ensureSheet_(ss, SHEET_NAMES.BIND_TOKENS, [
    COL.BIND_TOKENS.TOKEN, COL.BIND_TOKENS.EMPLOYEE_ID,
    COL.BIND_TOKENS.CREATED_AT, COL.BIND_TOKENS.EXPIRES_AT,
    COL.BIND_TOKENS.USED_AT, COL.BIND_TOKENS.USED_BY_LINE_ID
  ]);

  ensureSheet_(ss, SHEET_NAMES.HOLIDAYS, [
    COL.HOLIDAYS.DATE, COL.HOLIDAYS.YEAR, COL.HOLIDAYS.NAME,
    COL.HOLIDAYS.IS_RED, COL.HOLIDAYS.LUNAR, COL.HOLIDAYS.NOTE, COL.HOLIDAYS.SOURCE
  ]);

  seedStores_(ss);
  seedRoles_(ss);
  seedPermissions_(ss);
  seedSettings_(ss);
  seedAdminEmployee_(ss);
  protectSensitive_(ss);

  SpreadsheetApp.getUi().alert(
    'Bootstrap 完成！\n\n請檢查左下分頁列，所有 10 張表已建立。\n' +
    '欄位中英對照可參考「欄位對照表」分頁（執行 BuildGlossary 產生）。'
  );
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

const STORES_SEED_ = [
  ['LUCKY', '好運彩券行', '台中市西屯區長安路二段159號',
    24.17397559, 120.6652575, 100, '06:45', '16:30', '16:15', '01:45', 'active'],
  ['BINGO', '賓果彩券行', '台中市北屯區昌平路二段5-11號',
    24.17941266, 120.6902831, 100, '06:45', '16:30', '16:15', '01:45', 'active'],
  ['MONEY', '撿到錢投注站', '台中市北屯區崇德二路一段146-2號',
    24.17619498, 120.6961428, 100, '06:45', '16:30', '16:15', '01:45', 'active']
];

function seedStores_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.STORES);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, STORES_SEED_.length, STORES_SEED_[0].length).setValues(STORES_SEED_);
}

/**
 * 一次性遷移：用真實 store_id + GPS 座標覆寫，並把 admin home_store 對齊 LUCKY。
 * 已執行過可不再執行。
 */
function MigrateStoresAndAdmin() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const storesSheet = ss.getSheetByName(SHEET_NAMES.STORES);
  if (storesSheet.getLastRow() > 1) {
    storesSheet.getRange(2, 1, storesSheet.getLastRow() - 1, storesSheet.getLastColumn())
      .clearContent();
  }
  storesSheet.getRange(2, 1, STORES_SEED_.length, STORES_SEED_[0].length)
    .setValues(STORES_SEED_);

  const empSheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  const empData = empSheet.getDataRange().getValues();
  const headers = empData[0];
  const idCol = headers.indexOf(COL.EMPLOYEES.ID);
  const homeCol = headers.indexOf(COL.EMPLOYEES.HOME_STORE);
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][idCol] === 'E001' && empData[i][homeCol] === 'S001') {
      empSheet.getRange(i + 1, homeCol + 1).setValue('LUCKY');
    }
  }

  SpreadsheetApp.getUi().alert('Stores 已更新 + Admin home_store 已對齊 LUCKY。');
}

function seedRoles_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ROLES);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 3, 2).setValues([
    ['admin', '管理者（系統管理員、老闆）'],
    ['manager', '店長'],
    ['staff', '一般員工']
  ]);
}

function seedPermissions_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.PERMISSIONS);
  if (sheet.getLastRow() > 1) return;
  const keys = [
    'employee.create',
    'employee.edit',
    'employee.delete',
    'employee_sensitive.view',
    'employee_sensitive.edit',
    'store.edit',
    'schedule.edit_any_store',
    'correction.approve_for_my_store',
    'salary.edit_for_main_store',
    'report.view_any_store',
    'report.view_my_store',
    'holidays.sync_toggle',
    'backup.run',
    'backup.restore'
  ];
  const rows = [];
  keys.forEach(function (k) {
    rows.push(['admin', k, true]);
    rows.push(['manager', k, false]);
    rows.push(['staff', k, false]);
  });
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

function seedSettings_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 9, 3).setValues([
    ['web_app_url', '',
      '【請填】Apps Script Web App production /exec URL（從「管理部署」對話框複製）'],
    ['attendance_sheet_id', '', 'Attendance Sheet ID（執行 CreateAttendanceSheet 後自動填入）'],
    ['holidays_sync_enabled', 'true', '是否自動同步 taiwan-holidays（連 3 次失敗自動切 false）'],
    ['holidays_sync_failures', '0', '同步失敗計數'],
    ['backup_folder_id', '', 'Google Drive 備份資料夾 ID（首次備份時自動建立並寫入）'],
    ['line_login_channel_id', '', '【請填】LINE Login channel ID'],
    ['line_login_channel_secret', '', '【請填】LINE Login channel secret'],
    ['line_messaging_access_token', '', '【請填】LINE Messaging API channel access token'],
    ['admin_email_for_alerts', Session.getEffectiveUser().getEmail(),
      '系統異常通知 email（備份失敗、同步失敗等）']
  ]);
}

function seedAdminEmployee_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES);
  if (sheet.getLastRow() > 1) return;
  const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  sheet.getRange(2, 1, 1, 18).setValues([[
    'E001', '系統管理員', '', 'LUCKY', 'admin',
    'monthly', 29500, 200, today,
    today, '', '通過',
    'active', '', '', '',
    new Date(), new Date()
  ]]);
}

function protectSensitive_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.EMPLOYEES_SENSITIVE);
  try {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (protections.length > 0) return;

    const protection = sheet.protect()
      .setDescription('敏感資料：只限 admin 開啟 — 由 Bootstrap 自動設定');
    const me = Session.getEffectiveUser();
    protection.addEditor(me);
    const others = protection.getEditors()
      .filter(function (e) { return e.getEmail() !== me.getEmail(); });
    if (others.length > 0) {
      protection.removeEditors(others);
    }
    if (protection.canDomainEdit && protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } catch (err) {
    console.warn('protectSensitive_ failed: ' + err);
  }
}
