const TIMEZONE = 'Asia/Taipei';

/**
 * 分頁名稱。所有 .gs 程式碼存取分頁一律透過此常數，
 * 任何重命名只要改這裡，不用全程式搜尋取代。
 */
const SHEET_NAMES = {
  // Core Sheet
  EMPLOYEES: '員工',
  EMPLOYEES_SENSITIVE: '員工敏感資料',
  STORES: '門市',
  ROLES: '角色',
  PERMISSIONS: '權限',
  SALARY_HISTORY: '薪資異動紀錄',
  SETTINGS: '系統設定',
  BACKUP_LOG: '備份紀錄',
  BIND_TOKENS: '綁定連結',
  HOLIDAYS: '假日表',
  GLOSSARY: '欄位對照表',
  // Attendance Sheet
  SHIFT_TYPES: '班別',
  DEFAULT_SCHEDULE: '預設排班'
  // 年度分頁透過 attendanceSheetName_(prefix, year) 動態組合
};

/** 年度分頁前綴（與年份組成 "排班_2026" 等） */
const ATTENDANCE_PREFIX = {
  SCHEDULE: '排班',
  PUNCHES: '打卡',
  CORRECTIONS: '補卡',
  SWAP_REQUESTS: '換班申請',
  DAILY_AUDIT: '出勤稽核'
};

/**
 * 各分頁的欄位名稱。所有 indexOf(headers, '...') 一律走此常數。
 */
const COL = {
  EMPLOYEES: {
    ID: '員工編號',
    NAME: '姓名',
    LINE_USER_ID: 'LINE使用者ID',
    HOME_STORE: '主店',
    ROLE: '角色',
    EMPLOYMENT_TYPE: '雇用類型',
    MONTHLY_SALARY: '月薪',
    HOURLY_RATE: '時薪',
    SALARY_EFFECTIVE_DATE: '薪資生效日',
    HIRE_DATE: '到職日',
    PROBATION_END_DATE: '試用期截止日',
    PROBATION_STATUS: '試用期狀態',
    STATUS: '在職狀態',
    RESIGN_DATE: '離職日',
    EMERGENCY_CONTACT: '緊急聯絡人',
    EMERGENCY_PHONE: '緊急聯絡電話',
    CREATED_AT: '建立時間',
    UPDATED_AT: '更新時間'
  },
  EMPLOYEES_SENSITIVE: {
    ID: '員工編號',
    ID_NUMBER: '身分證字號',
    BANK_ACCOUNT: '銀行帳號',
    UPDATED_BY: '更新人',
    UPDATED_AT: '更新時間'
  },
  STORES: {
    ID: '門市代號',
    NAME: '名稱',
    ADDRESS: '地址',
    LAT: '緯度',
    LNG: '經度',
    RADIUS_M: '打卡半徑',
    MORNING_START: '早班預設上班',
    MORNING_END: '早班預設下班',
    EVENING_START: '晚班預設上班',
    EVENING_END: '晚班預設下班',
    STATUS: '狀態'
  },
  ROLES: {
    ROLE: '角色代號',
    DESCRIPTION: '說明'
  },
  PERMISSIONS: {
    ROLE: '角色',
    PERMISSION_KEY: '權限代號',
    GRANTED: '啟用'
  },
  SALARY_HISTORY: {
    EMPLOYEE_ID: '員工編號',
    FIELD: '欄位',
    OLD_VALUE: '舊值',
    NEW_VALUE: '新值',
    EFFECTIVE_DATE: '生效日',
    MODIFIED_BY: '修改人',
    MODIFIED_AT: '修改時間',
    REASON: '原因'
  },
  SETTINGS: {
    KEY: '設定代號',
    VALUE: '值',
    DESCRIPTION: '說明'
  },
  BACKUP_LOG: {
    BACKUP_AT: '備份時間',
    FILE_ID: '檔案ID',
    STATUS: '狀態',
    NOTE: '備註'
  },
  BIND_TOKENS: {
    TOKEN: 'token',
    EMPLOYEE_ID: '員工編號',
    CREATED_AT: '建立時間',
    EXPIRES_AT: '過期時間',
    USED_AT: '使用時間',
    USED_BY_LINE_ID: '綁定LINE_ID'
  },
  HOLIDAYS: {
    DATE: '日期',
    YEAR: '年度',
    NAME: '名稱',
    IS_RED: '紅字',
    LUNAR: '農曆',
    NOTE: '備註',
    SOURCE: '資料來源'
  },
  SHIFT_TYPES: {
    SHIFT_CODE: '班別代號',
    NAME: '名稱',
    IS_REGULAR: '常態班別',
    BREAK_MINUTES: '休息分鐘',
    BREAK_PERIODS: '休息時段'
  },
  DEFAULT_SCHEDULE: {
    STORE_ID: '門市代號',
    SHIFT_CODE: '班別代號',
    EMPLOYEE_ID: '員工編號',
    NOTE: '備註'
  },
  SCHEDULE: {
    SCHEDULE_ID: '排班ID',
    DATE: '日期',
    EMPLOYEE_ID: '員工編號',
    STORE_ID: '門市代號',
    SHIFT_CODE: '班別代號',
    START_TIME: '開始時間',
    END_TIME: '結束時間',
    PAY_MODE_OVERRIDE: '薪資模式覆寫',
    HOURLY_RATE_OVERRIDE: '時薪覆寫',
    STATUS: '狀態',
    CREATED_BY: '建立人',
    CREATED_AT: '建立時間',
    UPDATED_AT: '更新時間'
  },
  PUNCHES: {
    PUNCH_ID: '打卡ID',
    EMPLOYEE_ID: '員工編號',
    PUNCH_TYPE: '打卡類型',
    PUNCH_TIME: '打卡時間',
    STORE_ID: '門市代號',
    LAT: '緯度',
    LNG: '經度',
    GPS_STATUS: 'GPS狀態',
    DEVICE_INFO: '裝置資訊',
    CREATED_AT: '建立時間'
  },
  CORRECTIONS: {
    CORRECTION_ID: '補卡ID',
    EMPLOYEE_ID: '員工編號',
    TARGET_DATE: '目標日期',
    TARGET_SCHEDULE_ID: '目標排班ID',
    REASON: '原因',
    REQUESTED_IN_TIME: '申請上班時間',
    REQUESTED_OUT_TIME: '申請下班時間',
    STATUS: '狀態',
    APPROVER_ID: '簽核人',
    APPROVED_AT: '簽核時間',
    APPROVER_COMMENT: '簽核備註'
  },
  SWAP_REQUESTS: {
    SWAP_ID: '換班ID',
    FROM_EMPLOYEE_ID: '申請人員工編號',
    TO_EMPLOYEE_ID: '對方員工編號',
    FROM_SCHEDULE_ID: '申請人班次ID',
    TO_SCHEDULE_ID: '對方班次ID',
    STATUS: '狀態',
    REQUESTED_AT: '申請時間',
    COUNTERPARTY_RESPONDED_AT: '對方回覆時間',
    APPROVER_ID: '簽核人',
    APPROVED_AT: '簽核時間'
  },
  DAILY_AUDIT: {
    AUDIT_ID: '稽核ID',
    DATE: '日期',
    EMPLOYEE_ID: '員工編號',
    SCHEDULE_ID: '排班ID',
    STORE_ID: '門市代號',
    EXPECTED_START: '排定上班',
    EXPECTED_END: '排定下班',
    ACTUAL_IN: '實際上班',
    ACTUAL_OUT: '實際下班',
    LATE_MINUTES: '遲到分鐘',
    EARLY_LEAVE_MINUTES: '早退分鐘',
    MISSING_IN: '漏打上班',
    MISSING_OUT: '漏打下班',
    WORKED_MINUTES: '實際工時分鐘',
    OT_MINUTES_134: '加班分鐘_134倍',
    OT_MINUTES_167: '加班分鐘_167倍',
    IS_HOLIDAY: '紅字日',
    ANOMALY_CODES: '異常碼',
    STORE_ID_COST: '成本歸屬門市',
    COMPUTED_AT: '計算時間'
  }
};
