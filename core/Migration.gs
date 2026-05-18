/**
 * 一次性遷移：把現有 Core + Attendance Sheets 的分頁名稱與欄位名稱從英文改為中文。
 *
 * 用法：從 Apps Script 編輯器執行 MigrateRenameToChinese 一次即可。
 * 已遷移過的分頁會自動跳過，重複執行安全。
 */
function MigrateRenameToChinese() {
  migrateCoreSheet_();
  migrateAttendanceSheet_();
  SpreadsheetApp.getUi().alert(
    '中文化遷移完成 ✓\n\n' +
    'Core Sheet 和 Attendance Sheet 的分頁名稱與欄位名稱已全部改為中文。\n\n' +
    '接著建議執行 BuildGlossary 產生「欄位對照表」分頁。'
  );
}

function migrateCoreSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  CORE_MIGRATION_PLAN_.forEach(function (plan) {
    migrateOneSheet_(ss, plan);
  });
}

function migrateAttendanceSheet_() {
  let ss;
  try {
    ss = getAttendanceSpreadsheet_();
  } catch (e) {
    console.warn('Attendance Sheet 未設定，跳過');
    return;
  }
  ATTENDANCE_MIGRATION_PLAN_.forEach(function (plan) {
    migrateOneSheet_(ss, plan);
  });
  // 年度分頁
  ATTENDANCE_YEARLY_PLANS_.forEach(function (yearlyPlan) {
    for (let year = 2024; year <= 2030; year++) {
      const oldName = yearlyPlan.oldPrefix + '_' + year;
      const newName = yearlyPlan.newPrefix + '_' + year;
      if (ss.getSheetByName(oldName) || ss.getSheetByName(newName)) {
        migrateOneSheet_(ss, {
          oldName: oldName,
          newName: newName,
          cols: yearlyPlan.cols
        });
      }
    }
  });
}

function migrateOneSheet_(ss, plan) {
  let sheet = ss.getSheetByName(plan.oldName);
  if (!sheet) {
    sheet = ss.getSheetByName(plan.newName);
    if (!sheet) return; // 分頁不存在，跳過
  } else {
    sheet.setName(plan.newName);
  }

  if (sheet.getLastColumn() === 0) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const renamed = headers.map(function (h) {
    const found = plan.cols.find(function (c) { return c[0] === h; });
    return found ? found[1] : h;
  });
  sheet.getRange(1, 1, 1, renamed.length).setValues([renamed]);
}

// ============================================================================
// 遷移計畫：每個分頁 [oldName, newName] + 欄位 [[oldCol, newCol], ...]
// ============================================================================

const CORE_MIGRATION_PLAN_ = [
  {
    oldName: 'Employees', newName: SHEET_NAMES.EMPLOYEES,
    cols: [
      ['employee_id', COL.EMPLOYEES.ID],
      ['name', COL.EMPLOYEES.NAME],
      ['line_user_id', COL.EMPLOYEES.LINE_USER_ID],
      ['home_store', COL.EMPLOYEES.HOME_STORE],
      ['role', COL.EMPLOYEES.ROLE],
      ['employment_type', COL.EMPLOYEES.EMPLOYMENT_TYPE],
      ['monthly_salary', COL.EMPLOYEES.MONTHLY_SALARY],
      ['hourly_rate', COL.EMPLOYEES.HOURLY_RATE],
      ['salary_effective_date', COL.EMPLOYEES.SALARY_EFFECTIVE_DATE],
      ['hire_date', COL.EMPLOYEES.HIRE_DATE],
      ['probation_end_date', COL.EMPLOYEES.PROBATION_END_DATE],
      ['probation_status', COL.EMPLOYEES.PROBATION_STATUS],
      ['status', COL.EMPLOYEES.STATUS],
      ['resign_date', COL.EMPLOYEES.RESIGN_DATE],
      ['emergency_contact', COL.EMPLOYEES.EMERGENCY_CONTACT],
      ['emergency_phone', COL.EMPLOYEES.EMERGENCY_PHONE],
      ['created_at', COL.EMPLOYEES.CREATED_AT],
      ['updated_at', COL.EMPLOYEES.UPDATED_AT]
    ]
  },
  {
    oldName: 'EmployeesSensitive', newName: SHEET_NAMES.EMPLOYEES_SENSITIVE,
    cols: [
      ['employee_id', COL.EMPLOYEES_SENSITIVE.ID],
      ['id_number', COL.EMPLOYEES_SENSITIVE.ID_NUMBER],
      ['bank_account', COL.EMPLOYEES_SENSITIVE.BANK_ACCOUNT],
      ['updated_by', COL.EMPLOYEES_SENSITIVE.UPDATED_BY],
      ['updated_at', COL.EMPLOYEES_SENSITIVE.UPDATED_AT]
    ]
  },
  {
    oldName: 'Stores', newName: SHEET_NAMES.STORES,
    cols: [
      ['store_id', COL.STORES.ID],
      ['name', COL.STORES.NAME],
      ['address', COL.STORES.ADDRESS],
      ['lat', COL.STORES.LAT],
      ['lng', COL.STORES.LNG],
      ['radius_m', COL.STORES.RADIUS_M],
      ['default_morning_start', COL.STORES.MORNING_START],
      ['default_morning_end', COL.STORES.MORNING_END],
      ['default_evening_start', COL.STORES.EVENING_START],
      ['default_evening_end', COL.STORES.EVENING_END],
      ['status', COL.STORES.STATUS]
    ]
  },
  {
    oldName: 'Roles', newName: SHEET_NAMES.ROLES,
    cols: [
      ['role', COL.ROLES.ROLE],
      ['description', COL.ROLES.DESCRIPTION]
    ]
  },
  {
    oldName: 'Permissions', newName: SHEET_NAMES.PERMISSIONS,
    cols: [
      ['role', COL.PERMISSIONS.ROLE],
      ['permission_key', COL.PERMISSIONS.PERMISSION_KEY],
      ['granted', COL.PERMISSIONS.GRANTED]
    ]
  },
  {
    oldName: 'SalaryHistory', newName: SHEET_NAMES.SALARY_HISTORY,
    cols: [
      ['employee_id', COL.SALARY_HISTORY.EMPLOYEE_ID],
      ['field', COL.SALARY_HISTORY.FIELD],
      ['old_value', COL.SALARY_HISTORY.OLD_VALUE],
      ['new_value', COL.SALARY_HISTORY.NEW_VALUE],
      ['effective_date', COL.SALARY_HISTORY.EFFECTIVE_DATE],
      ['modified_by', COL.SALARY_HISTORY.MODIFIED_BY],
      ['modified_at', COL.SALARY_HISTORY.MODIFIED_AT],
      ['reason', COL.SALARY_HISTORY.REASON]
    ]
  },
  {
    oldName: 'Settings', newName: SHEET_NAMES.SETTINGS,
    cols: [
      ['key', COL.SETTINGS.KEY],
      ['value', COL.SETTINGS.VALUE],
      ['description', COL.SETTINGS.DESCRIPTION]
    ]
  },
  {
    oldName: 'Backup_Log', newName: SHEET_NAMES.BACKUP_LOG,
    cols: [
      ['backup_at', COL.BACKUP_LOG.BACKUP_AT],
      ['file_id', COL.BACKUP_LOG.FILE_ID],
      ['status', COL.BACKUP_LOG.STATUS],
      ['note', COL.BACKUP_LOG.NOTE]
    ]
  },
  {
    oldName: 'BindTokens', newName: SHEET_NAMES.BIND_TOKENS,
    cols: [
      ['token', COL.BIND_TOKENS.TOKEN],
      ['employee_id', COL.BIND_TOKENS.EMPLOYEE_ID],
      ['created_at', COL.BIND_TOKENS.CREATED_AT],
      ['expires_at', COL.BIND_TOKENS.EXPIRES_AT],
      ['used_at', COL.BIND_TOKENS.USED_AT],
      ['used_by_line_id', COL.BIND_TOKENS.USED_BY_LINE_ID]
    ]
  },
  {
    oldName: 'Holidays', newName: SHEET_NAMES.HOLIDAYS,
    cols: [
      ['date', COL.HOLIDAYS.DATE],
      ['year', COL.HOLIDAYS.YEAR],
      ['name', COL.HOLIDAYS.NAME],
      ['is_red', COL.HOLIDAYS.IS_RED],
      ['source', COL.HOLIDAYS.SOURCE]
    ]
  }
];

const ATTENDANCE_MIGRATION_PLAN_ = [
  {
    oldName: 'ShiftTypes', newName: SHEET_NAMES.SHIFT_TYPES,
    cols: [
      ['shift_code', COL.SHIFT_TYPES.SHIFT_CODE],
      ['name', COL.SHIFT_TYPES.NAME],
      ['is_regular', COL.SHIFT_TYPES.IS_REGULAR],
      ['break_minutes', COL.SHIFT_TYPES.BREAK_MINUTES],
      ['break_periods', COL.SHIFT_TYPES.BREAK_PERIODS]
    ]
  }
];

const ATTENDANCE_YEARLY_PLANS_ = [
  {
    oldPrefix: 'Schedule', newPrefix: ATTENDANCE_PREFIX.SCHEDULE,
    cols: [
      ['schedule_id', COL.SCHEDULE.SCHEDULE_ID],
      ['date', COL.SCHEDULE.DATE],
      ['employee_id', COL.SCHEDULE.EMPLOYEE_ID],
      ['store_id', COL.SCHEDULE.STORE_ID],
      ['shift_code', COL.SCHEDULE.SHIFT_CODE],
      ['start_time', COL.SCHEDULE.START_TIME],
      ['end_time', COL.SCHEDULE.END_TIME],
      ['pay_mode_override', COL.SCHEDULE.PAY_MODE_OVERRIDE],
      ['hourly_rate_override', COL.SCHEDULE.HOURLY_RATE_OVERRIDE],
      ['status', COL.SCHEDULE.STATUS],
      ['created_by', COL.SCHEDULE.CREATED_BY],
      ['created_at', COL.SCHEDULE.CREATED_AT],
      ['updated_at', COL.SCHEDULE.UPDATED_AT]
    ]
  },
  {
    oldPrefix: 'Punches', newPrefix: ATTENDANCE_PREFIX.PUNCHES,
    cols: [
      ['punch_id', COL.PUNCHES.PUNCH_ID],
      ['employee_id', COL.PUNCHES.EMPLOYEE_ID],
      ['punch_type', COL.PUNCHES.PUNCH_TYPE],
      ['punch_time', COL.PUNCHES.PUNCH_TIME],
      ['store_id', COL.PUNCHES.STORE_ID],
      ['lat', COL.PUNCHES.LAT],
      ['lng', COL.PUNCHES.LNG],
      ['gps_status', COL.PUNCHES.GPS_STATUS],
      ['device_info', COL.PUNCHES.DEVICE_INFO],
      ['created_at', COL.PUNCHES.CREATED_AT]
    ]
  },
  {
    oldPrefix: 'Corrections', newPrefix: ATTENDANCE_PREFIX.CORRECTIONS,
    cols: [
      ['correction_id', COL.CORRECTIONS.CORRECTION_ID],
      ['employee_id', COL.CORRECTIONS.EMPLOYEE_ID],
      ['target_date', COL.CORRECTIONS.TARGET_DATE],
      ['target_schedule_id', COL.CORRECTIONS.TARGET_SCHEDULE_ID],
      ['reason', COL.CORRECTIONS.REASON],
      ['requested_in_time', COL.CORRECTIONS.REQUESTED_IN_TIME],
      ['requested_out_time', COL.CORRECTIONS.REQUESTED_OUT_TIME],
      ['status', COL.CORRECTIONS.STATUS],
      ['approver_id', COL.CORRECTIONS.APPROVER_ID],
      ['approved_at', COL.CORRECTIONS.APPROVED_AT],
      ['approver_comment', COL.CORRECTIONS.APPROVER_COMMENT]
    ]
  },
  {
    oldPrefix: 'SwapRequests', newPrefix: ATTENDANCE_PREFIX.SWAP_REQUESTS,
    cols: [
      ['swap_id', COL.SWAP_REQUESTS.SWAP_ID],
      ['from_employee_id', COL.SWAP_REQUESTS.FROM_EMPLOYEE_ID],
      ['to_employee_id', COL.SWAP_REQUESTS.TO_EMPLOYEE_ID],
      ['from_schedule_id', COL.SWAP_REQUESTS.FROM_SCHEDULE_ID],
      ['to_schedule_id', COL.SWAP_REQUESTS.TO_SCHEDULE_ID],
      ['status', COL.SWAP_REQUESTS.STATUS],
      ['requested_at', COL.SWAP_REQUESTS.REQUESTED_AT],
      ['counterparty_responded_at', COL.SWAP_REQUESTS.COUNTERPARTY_RESPONDED_AT],
      ['approver_id', COL.SWAP_REQUESTS.APPROVER_ID],
      ['approved_at', COL.SWAP_REQUESTS.APPROVED_AT]
    ]
  },
  {
    oldPrefix: 'DailyAudit', newPrefix: ATTENDANCE_PREFIX.DAILY_AUDIT,
    cols: [
      ['audit_id', COL.DAILY_AUDIT.AUDIT_ID],
      ['date', COL.DAILY_AUDIT.DATE],
      ['employee_id', COL.DAILY_AUDIT.EMPLOYEE_ID],
      ['schedule_id', COL.DAILY_AUDIT.SCHEDULE_ID],
      ['store_id', COL.DAILY_AUDIT.STORE_ID],
      ['expected_start', COL.DAILY_AUDIT.EXPECTED_START],
      ['expected_end', COL.DAILY_AUDIT.EXPECTED_END],
      ['actual_in', COL.DAILY_AUDIT.ACTUAL_IN],
      ['actual_out', COL.DAILY_AUDIT.ACTUAL_OUT],
      ['late_minutes', COL.DAILY_AUDIT.LATE_MINUTES],
      ['early_leave_minutes', COL.DAILY_AUDIT.EARLY_LEAVE_MINUTES],
      ['missing_in', COL.DAILY_AUDIT.MISSING_IN],
      ['missing_out', COL.DAILY_AUDIT.MISSING_OUT],
      ['worked_minutes', COL.DAILY_AUDIT.WORKED_MINUTES],
      ['ot_minutes_134', COL.DAILY_AUDIT.OT_MINUTES_134],
      ['ot_minutes_167', COL.DAILY_AUDIT.OT_MINUTES_167],
      ['is_holiday', COL.DAILY_AUDIT.IS_HOLIDAY],
      ['anomaly_codes', COL.DAILY_AUDIT.ANOMALY_CODES],
      ['store_id_cost', COL.DAILY_AUDIT.STORE_ID_COST],
      ['computed_at', COL.DAILY_AUDIT.COMPUTED_AT]
    ]
  }
];
