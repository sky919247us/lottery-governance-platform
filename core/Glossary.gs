/**
 * 建立 / 重建「欄位對照表」分頁 — 列出所有分頁的欄位中英對照、型別、可填入值、說明。
 *
 * 從 Apps Script 編輯器執行 BuildGlossary 即可。重複執行會覆寫內容。
 */

function BuildGlossary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.GLOSSARY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.GLOSSARY);
  } else {
    sheet.clear();
  }

  const headers = ['分頁', '欄位（中文）', '欄位（英文代號）', '型別', '可填入值', '說明'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#fef3c7');
  sheet.setFrozenRows(1);

  const rows = GLOSSARY_DATA_;
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // 視覺：分頁區塊分隔
  let lastSheetName = null;
  let blockStart = 2;
  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    if (rows[i][0] !== lastSheetName) {
      if (lastSheetName !== null) {
        sheet.getRange(rowNum, 1).setBorder(true, false, false, false, false, false,
          '#666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
      lastSheetName = rows[i][0];
      // 每個分頁第一列加底色
      sheet.getRange(rowNum, 1, 1, headers.length).setBackground('#f1f5f9');
    }
  }

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 280);
  sheet.setColumnWidth(6, 360);

  SpreadsheetApp.getUi().alert(
    '✓ 欄位對照表已建立\n\n' +
    '共 ' + rows.length + ' 個欄位定義，分布於 ' + GLOSSARY_SHEET_COUNT_ + ' 個分頁。'
  );
}

// 計算分頁數（給 alert 用）
const GLOSSARY_SHEET_COUNT_ = 16; // Core 10 + Attendance 6（年度為 prefix）

/**
 * 對照表內容。每列：[分頁名, 欄位中文, 欄位英文, 型別, 可填入值, 說明]
 */
const GLOSSARY_DATA_ = [
  // ============= 員工 =============
  ['員工', '員工編號', 'employee_id', '字串', 'E001、E002…（自訂格式）', '主鍵，新增員工時遞增'],
  ['員工', '姓名', 'name', '字串', '中文姓名', ''],
  ['員工', 'LINE使用者ID', 'line_user_id', '字串', 'LINE OAuth 取得的 userId（U 開頭）', '空白代表尚未綁定 LINE'],
  ['員工', '主店', 'home_store', '字串（FK）', 'LUCKY / BINGO / MONEY', '對應「門市」分頁的「門市代號」'],
  ['員工', '角色', 'role', '列舉', 'admin / manager / staff', 'admin=管理員、manager=店長、staff=一般員工'],
  ['員工', '雇用類型', 'employment_type', '列舉', 'monthly / hourly', 'monthly=月薪制、hourly=時薪制'],
  ['員工', '月薪', 'monthly_salary', '數字', '例 29500', '月薪人員的當前月薪。可由 admin/店長隨時編輯'],
  ['員工', '時薪', 'hourly_rate', '數字', '例 200', '時薪人員的時薪 + 月薪人員「跑班」時的時薪'],
  ['員工', '薪資生效日', 'salary_effective_date', '日期', 'YYYY-MM-DD', '當前薪資生效日，用於跨期結算'],
  ['員工', '到職日', 'hire_date', '日期', 'YYYY-MM-DD', '影響特休年資 + 試用期計算'],
  ['員工', '試用期截止日', 'probation_end_date', '日期', 'YYYY-MM-DD', '通常 = hire_date + 30 天'],
  ['員工', '試用期狀態', 'probation_status', '列舉', '評估中 / 通過 / 未通過', '到期 7 天前 LINE 推播店長提醒考核'],
  ['員工', '在職狀態', 'status', '列舉', 'active / resigned / leave', 'active=在職、resigned=離職、leave=留職停薪'],
  ['員工', '離職日', 'resign_date', '日期', 'YYYY-MM-DD（離職才填）', ''],
  ['員工', '緊急聯絡人', 'emergency_contact', '字串', '姓名', ''],
  ['員工', '緊急聯絡電話', 'emergency_phone', '字串', '電話', ''],
  ['員工', '建立時間', 'created_at', '時間戳', '自動', ''],
  ['員工', '更新時間', 'updated_at', '時間戳', '自動', ''],

  // ============= 員工敏感資料 =============
  ['員工敏感資料', '員工編號', 'employee_id', '字串（FK）', '對應「員工」', '受 ACL 保護，僅 admin 可開啟此分頁'],
  ['員工敏感資料', '身分證字號', 'id_number', '字串', '10 字元', ''],
  ['員工敏感資料', '銀行帳號', 'bank_account', '字串', '中國信託帳號', ''],
  ['員工敏感資料', '更新人', 'updated_by', '字串', 'employee_id', ''],
  ['員工敏感資料', '更新時間', 'updated_at', '時間戳', '自動', ''],

  // ============= 門市 =============
  ['門市', '門市代號', 'store_id', '字串', 'LUCKY / BINGO / MONEY', '主鍵'],
  ['門市', '名稱', 'name', '字串', '好運彩券行 / 賓果彩券行 / 撿到錢投注站', ''],
  ['門市', '地址', 'address', '字串', '完整地址', ''],
  ['門市', '緯度', 'lat', '小數', '24.xxxxxxx', '打卡 GPS 比對'],
  ['門市', '經度', 'lng', '小數', '120.xxxxxxx', '打卡 GPS 比對'],
  ['門市', '打卡半徑', 'radius_m', '整數', '預設 100（公尺）', 'GPS 容許誤差範圍'],
  ['門市', '早班預設上班', 'default_morning_start', '時間', 'HH:mm', '早班排班預設起始時間'],
  ['門市', '早班預設下班', 'default_morning_end', '時間', 'HH:mm', ''],
  ['門市', '晚班預設上班', 'default_evening_start', '時間', 'HH:mm', ''],
  ['門市', '晚班預設下班', 'default_evening_end', '時間', 'HH:mm', ''],
  ['門市', '狀態', 'status', '列舉', 'active / closed', ''],

  // ============= 角色 =============
  ['角色', '角色代號', 'role', '字串', 'admin / manager / staff', ''],
  ['角色', '說明', 'description', '字串', '中文角色名稱', ''],

  // ============= 權限 =============
  ['權限', '角色', 'role', '字串（FK）', 'admin / manager / staff', ''],
  ['權限', '權限代號', 'permission_key', '字串', 'employee.create / schedule.edit_any_store …', '見系統權限清單'],
  ['權限', '啟用', 'granted', '布林', 'TRUE / FALSE', 'admin 預設全 TRUE，manager/staff 由 admin 勾選'],

  // ============= 薪資異動紀錄 =============
  ['薪資異動紀錄', '員工編號', 'employee_id', '字串（FK）', '', ''],
  ['薪資異動紀錄', '欄位', 'field', '列舉', 'monthly_salary / hourly_rate', '哪個欄位被改'],
  ['薪資異動紀錄', '舊值', 'old_value', '數字', '', ''],
  ['薪資異動紀錄', '新值', 'new_value', '數字', '', ''],
  ['薪資異動紀錄', '生效日', 'effective_date', '日期', 'YYYY-MM-DD', ''],
  ['薪資異動紀錄', '修改人', 'modified_by', '字串', 'admin 或店長的 employee_id', ''],
  ['薪資異動紀錄', '修改時間', 'modified_at', '時間戳', '自動', ''],
  ['薪資異動紀錄', '原因', 'reason', '字串', '可選填', ''],

  // ============= 系統設定 =============
  ['系統設定', '設定代號', 'key', '字串', 'web_app_url / line_login_channel_id …', ''],
  ['系統設定', '值', 'value', '字串', '', ''],
  ['系統設定', '說明', 'description', '字串', '', ''],

  // ============= 備份紀錄 =============
  ['備份紀錄', '備份時間', 'backup_at', '時間戳', '', ''],
  ['備份紀錄', '檔案ID', 'file_id', '字串', 'Drive file ID', ''],
  ['備份紀錄', '狀態', 'status', '列舉', 'success / failed', ''],
  ['備份紀錄', '備註', 'note', '字串', '失敗時記錯誤訊息', ''],

  // ============= 綁定連結 =============
  ['綁定連結', 'token', 'token', '字串', 'UUID（32 字元）', '一次性綁定 token'],
  ['綁定連結', '員工編號', 'employee_id', '字串（FK）', '', '這 token 綁哪個員工'],
  ['綁定連結', '建立時間', 'created_at', '時間戳', '', ''],
  ['綁定連結', '過期時間', 'expires_at', '時間戳', '建立時間 + 24 小時', ''],
  ['綁定連結', '使用時間', 'used_at', '時間戳', '空=未使用 / 有時間=已使用', ''],
  ['綁定連結', '綁定LINE_ID', 'used_by_line_id', '字串', '使用者 LINE userId', ''],

  // ============= 假日表 =============
  ['假日表', '日期', 'date', '日期', 'YYYY-MM-DD', ''],
  ['假日表', '年度', 'year', '整數', '2026', '方便按年篩選'],
  ['假日表', '名稱', 'name', '字串', '例「春節」、「228」', '純週末為空白'],
  ['假日表', '紅字', 'is_red', '布林', 'TRUE / FALSE', 'TRUE=紅字假日；補班日為 FALSE'],
  ['假日表', '資料來源', 'source', '列舉', 'auto / manual', 'auto=API 同步；manual=admin 覆寫，永不被覆蓋'],

  // ============= 班別（Attendance Sheet）=============
  ['班別', '班別代號', 'shift_code', '列舉', 'MORNING / EVENING / NIGHT', ''],
  ['班別', '名稱', 'name', '字串', '早班 / 晚班 / 大夜', ''],
  ['班別', '常態班別', 'is_regular', '布林', 'TRUE / FALSE', '早晚為 TRUE，大夜 FALSE（彈性班）'],
  ['班別', '休息分鐘', 'break_minutes', '整數', '早晚 60 / 大夜 0', '工時計算扣除'],
  ['班別', '休息時段', 'break_periods', '字串', '"10:00-10:30,14:30-15:00"', '逗號分隔多段，僅供顯示'],

  // ============= 排班_YYYY =============
  ['排班_YYYY', '排班ID', 'schedule_id', '字串', 'SCH-YYYYMMDD-員工ID-班別', '主鍵'],
  ['排班_YYYY', '日期', 'date', '日期', 'YYYY-MM-DD', '跨日大夜歸上班那天'],
  ['排班_YYYY', '員工編號', 'employee_id', '字串（FK）', '', ''],
  ['排班_YYYY', '門市代號', 'store_id', '字串（FK）', 'LUCKY / BINGO / MONEY', '實際上班店（可跨店）'],
  ['排班_YYYY', '班別代號', 'shift_code', '字串（FK）', 'MORNING / EVENING / NIGHT', ''],
  ['排班_YYYY', '開始時間', 'start_time', '時間', 'HH:mm', '可覆寫該店預設'],
  ['排班_YYYY', '結束時間', 'end_time', '時間', 'HH:mm', '大夜可跨日'],
  ['排班_YYYY', '薪資模式覆寫', 'pay_mode_override', '列舉', '空白 / hourly', 'hourly=這班按時薪結算（跑班）'],
  ['排班_YYYY', '時薪覆寫', 'hourly_rate_override', '數字', '預設帶當前時薪（200）', '跑班時可單筆改'],
  ['排班_YYYY', '狀態', 'status', '列舉', '草稿 / 已確認 / 已換班 / 已取消', ''],
  ['排班_YYYY', '建立人', 'created_by', '字串', 'employee_id', '員工自填或店長代排'],
  ['排班_YYYY', '建立時間', 'created_at', '時間戳', '', ''],
  ['排班_YYYY', '更新時間', 'updated_at', '時間戳', '', ''],

  // ============= 打卡_YYYY =============
  ['打卡_YYYY', '打卡ID', 'punch_id', '字串', 'PCH-YYYYMMDDHHMMSS-員工ID', '主鍵'],
  ['打卡_YYYY', '員工編號', 'employee_id', '字串（FK）', '', ''],
  ['打卡_YYYY', '打卡類型', 'punch_type', '列舉', 'IN / OUT', 'IN=上班 / OUT=下班'],
  ['打卡_YYYY', '打卡時間', 'punch_time', '時間戳', '', '伺服器收到時的時間'],
  ['打卡_YYYY', '門市代號', 'store_id', '字串（FK）', '', '系統依當日排班自動帶入'],
  ['打卡_YYYY', '緯度', 'lat', '小數', '', '員工手機 GPS'],
  ['打卡_YYYY', '經度', 'lng', '小數', '', '員工手機 GPS'],
  ['打卡_YYYY', 'GPS狀態', 'gps_status', '列舉', 'ok / outlier / out_of_range', 'ok=範圍內、outlier=飄移收下、out_of_range=不符'],
  ['打卡_YYYY', '裝置資訊', 'device_info', '字串', 'UA + IP', '稽核用'],
  ['打卡_YYYY', '建立時間', 'created_at', '時間戳', '自動', ''],

  // ============= 補卡_YYYY =============
  ['補卡_YYYY', '補卡ID', 'correction_id', '字串', 'COR-流水號', ''],
  ['補卡_YYYY', '員工編號', 'employee_id', '字串', '', '申請人'],
  ['補卡_YYYY', '目標日期', 'target_date', '日期', '', '要補哪一天'],
  ['補卡_YYYY', '目標排班ID', 'target_schedule_id', '字串（FK）', '', ''],
  ['補卡_YYYY', '原因', 'reason', '字串', '', ''],
  ['補卡_YYYY', '申請上班時間', 'requested_in_time', '時間戳', '', ''],
  ['補卡_YYYY', '申請下班時間', 'requested_out_time', '時間戳', '', ''],
  ['補卡_YYYY', '狀態', 'status', '列舉', '待簽核 / 已核准 / 已駁回', ''],
  ['補卡_YYYY', '簽核人', 'approver_id', '字串', '店長 employee_id', ''],
  ['補卡_YYYY', '簽核時間', 'approved_at', '時間戳', '', ''],
  ['補卡_YYYY', '簽核備註', 'approver_comment', '字串', '可選填', ''],

  // ============= 換班申請_YYYY =============
  ['換班申請_YYYY', '換班ID', 'swap_id', '字串', 'SWP-流水號', ''],
  ['換班申請_YYYY', '申請人員工編號', 'from_employee_id', '字串', '', '發起換班的人'],
  ['換班申請_YYYY', '對方員工編號', 'to_employee_id', '字串', '', '要交換的對象'],
  ['換班申請_YYYY', '申請人班次ID', 'from_schedule_id', '字串（FK）', '', '申請人要換掉的班'],
  ['換班申請_YYYY', '對方班次ID', 'to_schedule_id', '字串（FK）', '', '對方原本的班'],
  ['換班申請_YYYY', '狀態', 'status', '列舉', '待對方同意 / 待店長核可 / 已生效 / 駁回', ''],
  ['換班申請_YYYY', '申請時間', 'requested_at', '時間戳', '', ''],
  ['換班申請_YYYY', '對方回覆時間', 'counterparty_responded_at', '時間戳', '', ''],
  ['換班申請_YYYY', '簽核人', 'approver_id', '字串', '店長 employee_id', ''],
  ['換班申請_YYYY', '簽核時間', 'approved_at', '時間戳', '', ''],

  // ============= 出勤稽核_YYYY =============
  ['出勤稽核_YYYY', '稽核ID', 'audit_id', '字串', 'AUD-流水號', ''],
  ['出勤稽核_YYYY', '日期', 'date', '日期', '', ''],
  ['出勤稽核_YYYY', '員工編號', 'employee_id', '字串（FK）', '', ''],
  ['出勤稽核_YYYY', '排班ID', 'schedule_id', '字串（FK）', '', '對應的排班列'],
  ['出勤稽核_YYYY', '門市代號', 'store_id', '字串（FK）', '', ''],
  ['出勤稽核_YYYY', '排定上班', 'expected_start', '時間戳', '', ''],
  ['出勤稽核_YYYY', '排定下班', 'expected_end', '時間戳', '', ''],
  ['出勤稽核_YYYY', '實際上班', 'actual_in', '時間戳', '從打卡解析', ''],
  ['出勤稽核_YYYY', '實際下班', 'actual_out', '時間戳', '從打卡解析', ''],
  ['出勤稽核_YYYY', '遲到分鐘', 'late_minutes', '整數', '> 3 才有值', '容差 3 分鐘'],
  ['出勤稽核_YYYY', '早退分鐘', 'early_leave_minutes', '整數', '', ''],
  ['出勤稽核_YYYY', '漏打上班', 'missing_in', '布林', 'TRUE / FALSE', ''],
  ['出勤稽核_YYYY', '漏打下班', 'missing_out', '布林', 'TRUE / FALSE', ''],
  ['出勤稽核_YYYY', '實際工時分鐘', 'worked_minutes', '整數', '已扣休息時間', '早晚扣 60 分、大夜扣 0'],
  ['出勤稽核_YYYY', '加班分鐘_134倍', 'ot_minutes_134', '整數', '', '8–10 小時段'],
  ['出勤稽核_YYYY', '加班分鐘_167倍', 'ot_minutes_167', '整數', '', '> 10 小時段'],
  ['出勤稽核_YYYY', '紅字日', 'is_holiday', '布林', 'TRUE / FALSE', '依假日表判定'],
  ['出勤稽核_YYYY', '異常碼', 'anomaly_codes', '字串', '例 LATE,GPS_OUTLIER', '逗號分隔多碼'],
  ['出勤稽核_YYYY', '成本歸屬門市', 'store_id_cost', '字串（FK）', '', '通常 = store_id'],
  ['出勤稽核_YYYY', '計算時間', 'computed_at', '時間戳', '自動', '']
];
