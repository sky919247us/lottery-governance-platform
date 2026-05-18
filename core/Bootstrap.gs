/**
 * Bootstrap — 一鍵建立 Core Sheet 所有分頁 + 預設資料
 *
 * 從 Apps Script 編輯器手動執行一次即可。重複執行安全（已存在的分頁不會被覆蓋）。
 */
function Bootstrap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEET_NAMES.EMPLOYEES, [
    'employee_id', 'name', 'line_user_id', 'home_store', 'role',
    'employment_type', 'monthly_salary', 'hourly_rate', 'salary_effective_date',
    'hire_date', 'probation_end_date', 'probation_status',
    'status', 'resign_date', 'emergency_contact', 'emergency_phone',
    'created_at', 'updated_at'
  ]);

  ensureSheet_(ss, SHEET_NAMES.EMPLOYEES_SENSITIVE, [
    'employee_id', 'id_number', 'bank_account', 'updated_by', 'updated_at'
  ]);

  ensureSheet_(ss, SHEET_NAMES.STORES, [
    'store_id', 'name', 'address', 'lat', 'lng', 'radius_m',
    'default_morning_start', 'default_morning_end',
    'default_evening_start', 'default_evening_end',
    'status'
  ]);

  ensureSheet_(ss, SHEET_NAMES.ROLES, ['role', 'description']);

  ensureSheet_(ss, SHEET_NAMES.PERMISSIONS, ['role', 'permission_key', 'granted']);

  ensureSheet_(ss, SHEET_NAMES.SALARY_HISTORY, [
    'employee_id', 'field', 'old_value', 'new_value', 'effective_date',
    'modified_by', 'modified_at', 'reason'
  ]);

  ensureSheet_(ss, SHEET_NAMES.SETTINGS, ['key', 'value', 'description']);

  ensureSheet_(ss, SHEET_NAMES.BACKUP_LOG, ['backup_at', 'file_id', 'status', 'note']);

  ensureSheet_(ss, SHEET_NAMES.BIND_TOKENS, [
    'token', 'employee_id', 'created_at', 'expires_at', 'used_at', 'used_by_line_id'
  ]);

  seedStores_(ss);
  seedRoles_(ss);
  seedPermissions_(ss);
  seedSettings_(ss);
  seedAdminEmployee_(ss);
  protectSensitive_(ss);

  SpreadsheetApp.getUi().alert(
    'Bootstrap 完成！\n\n' +
    '請檢查左下分頁列：應有 Employees / EmployeesSensitive / Stores / Roles / ' +
    'Permissions / SalaryHistory / Settings / Backup_Log 共 8 張表，且各表已有預設資料。\n\n' +
    '下一步：到 Stores 填入三間店的 GPS 座標（lat/lng），到 Settings 填入 LINE channel 資料。'
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

function seedStores_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.STORES);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, STORES_SEED_.length, STORES_SEED_[0].length).setValues(STORES_SEED_);
}

const STORES_SEED_ = [
  ['LUCKY', '好運彩券行', '台中市西屯區長安路二段159號',
    24.17397559, 120.6652575, 100, '06:45', '16:30', '16:15', '01:45', 'active'],
  ['BINGO', '賓果彩券行', '台中市北屯區昌平路二段5-11號',
    24.17941266, 120.6902831, 100, '06:45', '16:30', '16:15', '01:45', 'active'],
  ['MONEY', '撿到錢投注站', '台中市北屯區崇德二路一段146-2號',
    24.17619498, 120.6961428, 100, '06:45', '16:30', '16:15', '01:45', 'active']
];

/**
 * 一次性遷移：用真實 store_id（BINGO/LUCKY/MONEY）和 GPS 座標覆蓋 Stores 表，
 * 並同步把 admin 員工 (E001) 的 home_store 從 S001 改成 LUCKY。
 *
 * 只在 Bootstrap 跑過、需要更新時手動執行一次即可。
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
  const idCol = headers.indexOf('employee_id');
  const homeCol = headers.indexOf('home_store');
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][idCol] === 'E001' && empData[i][homeCol] === 'S001') {
      empSheet.getRange(i + 1, homeCol + 1).setValue('LUCKY');
    }
  }

  SpreadsheetApp.getUi().alert(
    'Stores 已更新：\n' +
    '• LUCKY 好運彩券行\n' +
    '• BINGO 賓果彩券行\n' +
    '• MONEY 撿到錢投注站\n\n' +
    'Admin (E001) home_store 已從 S001 改為 LUCKY。'
  );
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
  sheet.getRange(2, 1, 7, 3).setValues([
    ['web_app_url', '',
      '【請填】Apps Script Web App production /exec URL（從「管理部署」對話框複製）'],
    ['holidays_sync_enabled', 'true', '是否自動同步 taiwan-holidays（連 3 次失敗自動切 false）'],
    ['holidays_sync_failures', '0', '同步失敗計數'],
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
