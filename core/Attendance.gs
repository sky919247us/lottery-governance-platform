/**
 * Attendance 模組 — 排班 / 打卡 / 出勤稽核
 *
 * 資料分離原則：
 *   Core Sheet（員工 / 門市 / 權限）獨立一份；
 *   Attendance 資料另開一份 Sheet（按年分頁 排班_2026 / 打卡_2026 等）。
 *
 * 從 Apps Script 編輯器執行 CreateAttendanceSheet 一次完成。
 */

/** 帶年份的分頁名稱：排班_2026、打卡_2026 等 */
function attendanceSheetName_(prefix, year) {
  return prefix + '_' + year;
}

/** 從 Settings 讀取 Attendance Sheet 物件，不存在則 throw */
function getAttendanceSpreadsheet_() {
  const id = getSetting_('attendance_sheet_id');
  if (!id) throw new Error('Settings.attendance_sheet_id 未設定。請先執行 CreateAttendanceSheet。');
  return SpreadsheetApp.openById(String(id));
}

/** 主入口：建立 Attendance Sheet + 全部 schema + 種子資料 */
function CreateAttendanceSheet() {
  const existingId = getSetting_('attendance_sheet_id');
  if (existingId) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      'Attendance Sheet 已存在',
      'Settings.attendance_sheet_id = ' + existingId + '\n\n' +
      '若該檔已被刪除，按「是」會建立新的並覆寫此 ID。\n' +
      '若該檔仍在，按「否」取消。',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  const coreSs = SpreadsheetApp.getActiveSpreadsheet();
  const coreFile = DriveApp.getFileById(coreSs.getId());
  const parents = coreFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const year = new Date().getFullYear();
  const newSs = SpreadsheetApp.create('[彩券行] Attendance');
  const newFile = DriveApp.getFileById(newSs.getId());
  parent.addFile(newFile);
  if (parent.getId() !== DriveApp.getRootFolder().getId()) {
    DriveApp.getRootFolder().removeFile(newFile);
  }

  bootstrapAttendanceSheet_(newSs, year);
  setSetting_('attendance_sheet_id', newSs.getId());

  SpreadsheetApp.getUi().alert(
    '✓ Attendance Sheet 建立完成\n\n' +
    'ID 已寫入 Core Settings.attendance_sheet_id\n\n' +
    'URL：\n' + newSs.getUrl() + '\n\n' +
    '已建立分頁：' + SHEET_NAMES.SHIFT_TYPES + ' 及 ' + year + ' 年度 5 張分頁。'
  );
}

/** 顯示 Attendance Sheet 的完整 URL（可選取複製） */
function ShowAttendanceSheetUrl() {
  const id = getSetting_('attendance_sheet_id');
  if (!id) {
    SpreadsheetApp.getUi().alert('Settings.attendance_sheet_id 未設定。請先執行 CreateAttendanceSheet。');
    return;
  }
  const ss = SpreadsheetApp.openById(String(id));
  const url = ss.getUrl();
  const html =
    '<div style="font-family:sans-serif">' +
    '<p>Attendance Sheet URL：</p>' +
    '<input type="text" value="' + url + '" style="width:100%;font-size:12px;padding:6px" readonly onclick="this.select()">' +
    '<p style="margin-top:16px;color:#555;font-size:12px">ID: ' + id + '</p>' +
    '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(200),
    'Attendance Sheet'
  );
}

/** 每年初手動執行：為新年度開啟一組分頁 */
function CreateAttendanceSheetsForNextYear() {
  const ss = getAttendanceSpreadsheet_();
  const year = new Date().getFullYear() + 1;
  createYearlySheets_(ss, year);
  SpreadsheetApp.getUi().alert('✓ 已為 ' + year + ' 年建立分頁');
}

function bootstrapAttendanceSheet_(ss, year) {
  const defaultSheet = ss.getSheets()[0];

  const shiftTypesSheet = ss.insertSheet(SHEET_NAMES.SHIFT_TYPES);
  setHeaderRow_(shiftTypesSheet, [
    COL.SHIFT_TYPES.SHIFT_CODE, COL.SHIFT_TYPES.NAME,
    COL.SHIFT_TYPES.IS_REGULAR, COL.SHIFT_TYPES.BREAK_MINUTES,
    COL.SHIFT_TYPES.BREAK_PERIODS
  ]);
  shiftTypesSheet.getRange(2, 1, 3, 5).setValues([
    ['MORNING', '早班', true, 60, '10:00-10:30,14:30-15:00'],
    ['EVENING', '晚班', true, 60, '20:00-20:30,00:30-01:00'],
    ['NIGHT', '大夜', false, 0, '']
  ]);

  createYearlySheets_(ss, year);

  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

function createYearlySheets_(ss, year) {
  const sched = insertSheetIfMissing_(ss, attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, year));
  setHeaderRow_(sched, [
    COL.SCHEDULE.SCHEDULE_ID, COL.SCHEDULE.DATE, COL.SCHEDULE.EMPLOYEE_ID,
    COL.SCHEDULE.STORE_ID, COL.SCHEDULE.SHIFT_CODE,
    COL.SCHEDULE.START_TIME, COL.SCHEDULE.END_TIME,
    COL.SCHEDULE.PAY_MODE_OVERRIDE, COL.SCHEDULE.HOURLY_RATE_OVERRIDE,
    COL.SCHEDULE.STATUS, COL.SCHEDULE.CREATED_BY,
    COL.SCHEDULE.CREATED_AT, COL.SCHEDULE.UPDATED_AT
  ]);

  const punches = insertSheetIfMissing_(ss, attendanceSheetName_(ATTENDANCE_PREFIX.PUNCHES, year));
  setHeaderRow_(punches, [
    COL.PUNCHES.PUNCH_ID, COL.PUNCHES.EMPLOYEE_ID, COL.PUNCHES.PUNCH_TYPE,
    COL.PUNCHES.PUNCH_TIME, COL.PUNCHES.STORE_ID,
    COL.PUNCHES.LAT, COL.PUNCHES.LNG, COL.PUNCHES.GPS_STATUS,
    COL.PUNCHES.DEVICE_INFO, COL.PUNCHES.CREATED_AT
  ]);

  const corrs = insertSheetIfMissing_(ss, attendanceSheetName_(ATTENDANCE_PREFIX.CORRECTIONS, year));
  setHeaderRow_(corrs, [
    COL.CORRECTIONS.CORRECTION_ID, COL.CORRECTIONS.EMPLOYEE_ID,
    COL.CORRECTIONS.TARGET_DATE, COL.CORRECTIONS.TARGET_SCHEDULE_ID,
    COL.CORRECTIONS.REASON, COL.CORRECTIONS.REQUESTED_IN_TIME,
    COL.CORRECTIONS.REQUESTED_OUT_TIME, COL.CORRECTIONS.STATUS,
    COL.CORRECTIONS.APPROVER_ID, COL.CORRECTIONS.APPROVED_AT,
    COL.CORRECTIONS.APPROVER_COMMENT
  ]);

  const swaps = insertSheetIfMissing_(ss, attendanceSheetName_(ATTENDANCE_PREFIX.SWAP_REQUESTS, year));
  setHeaderRow_(swaps, [
    COL.SWAP_REQUESTS.SWAP_ID,
    COL.SWAP_REQUESTS.FROM_EMPLOYEE_ID, COL.SWAP_REQUESTS.TO_EMPLOYEE_ID,
    COL.SWAP_REQUESTS.FROM_SCHEDULE_ID, COL.SWAP_REQUESTS.TO_SCHEDULE_ID,
    COL.SWAP_REQUESTS.STATUS,
    COL.SWAP_REQUESTS.REQUESTED_AT, COL.SWAP_REQUESTS.COUNTERPARTY_RESPONDED_AT,
    COL.SWAP_REQUESTS.APPROVER_ID, COL.SWAP_REQUESTS.APPROVED_AT
  ]);

  const audit = insertSheetIfMissing_(ss, attendanceSheetName_(ATTENDANCE_PREFIX.DAILY_AUDIT, year));
  setHeaderRow_(audit, [
    COL.DAILY_AUDIT.AUDIT_ID, COL.DAILY_AUDIT.DATE,
    COL.DAILY_AUDIT.EMPLOYEE_ID, COL.DAILY_AUDIT.SCHEDULE_ID, COL.DAILY_AUDIT.STORE_ID,
    COL.DAILY_AUDIT.EXPECTED_START, COL.DAILY_AUDIT.EXPECTED_END,
    COL.DAILY_AUDIT.ACTUAL_IN, COL.DAILY_AUDIT.ACTUAL_OUT,
    COL.DAILY_AUDIT.LATE_MINUTES, COL.DAILY_AUDIT.EARLY_LEAVE_MINUTES,
    COL.DAILY_AUDIT.MISSING_IN, COL.DAILY_AUDIT.MISSING_OUT,
    COL.DAILY_AUDIT.WORKED_MINUTES,
    COL.DAILY_AUDIT.OT_MINUTES_134, COL.DAILY_AUDIT.OT_MINUTES_167,
    COL.DAILY_AUDIT.IS_HOLIDAY, COL.DAILY_AUDIT.ANOMALY_CODES,
    COL.DAILY_AUDIT.STORE_ID_COST, COL.DAILY_AUDIT.COMPUTED_AT
  ]);
}

function insertSheetIfMissing_(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) return existing;
  return ss.insertSheet(name);
}

function setHeaderRow_(sheet, headers) {
  if (sheet.getLastRow() > 0) return;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}
