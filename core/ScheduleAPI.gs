/**
 * 排班視圖後端 API
 *
 * 提供前端「三店月曆總覽」所需的所有資料：
 *   - 員工（含色票）
 *   - 門市
 *   - 班別
 *   - 指定月份的排班
 *   - 指定月份的紅字日
 *
 * 一次拉齊，避免前端多次來回。
 */

/** 主入口：給排班視圖用的整月資料 */
function getScheduleViewData(year, month) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  const holidays = getHolidaysForMonth_(y, m);
  return {
    year: y,
    month: m,
    today: Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd'),
    employees: getEmployeesForViewCached_(),
    stores: getStoresForViewCached_(),
    shiftTypes: getShiftTypesForViewCached_(),
    schedules: getSchedulesForMonth_(y, m),
    holidays: holidays,
    dayNotes: buildDayNotesForMonth_(holidays)
  };
}

/** 快取常變資料 — 5 分鐘 TTL。員工 / 門市 / 班別 變動不頻繁 */
function getEmployeesForViewCached_() {
  return getCachedOrCompute_('emp_view', 300, getEmployeesForView_);
}
function getStoresForViewCached_() {
  return getCachedOrCompute_('stores_view', 600, getStoresForView_);
}
function getShiftTypesForViewCached_() {
  return getCachedOrCompute_('shifttypes_view', 1800, getShiftTypesForView_);
}

function getCachedOrCompute_(key, ttl, fn) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) return JSON.parse(cached);
    const result = fn();
    const s = JSON.stringify(result);
    if (s.length < 100000) cache.put(key, s, ttl);
    return result;
  } catch (e) {
    return fn();
  }
}

/** Cache invalidation — 寫入員工/門市/班別後呼叫，立即失效 */
function invalidateViewCaches_() {
  try {
    CacheService.getScriptCache().removeAll(['emp_view', 'stores_view', 'shifttypes_view']);
  } catch (e) {}
}

function getEmployeesForView_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idCol = headers.indexOf(COL.EMPLOYEES.ID);
  const nameCol = headers.indexOf(COL.EMPLOYEES.NAME);
  const roleCol = headers.indexOf(COL.EMPLOYEES.ROLE);
  const homeCol = headers.indexOf(COL.EMPLOYEES.HOME_STORE);
  const statusCol = headers.indexOf(COL.EMPLOYEES.STATUS);

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][idCol]) continue;
    const st = String(data[i][statusCol] || '').trim().toLowerCase();
    // 排除明確離職/留職停薪；空白視為在職
    if (st === 'resigned' || st === 'leave') continue;
    out.push({
      id: data[i][idCol],
      name: data[i][nameCol],
      role: data[i][roleCol],
      home_store: data[i][homeCol]
    });
  }
  // 按員工編號排序後逐一分配色卡（保證唯一且穩定）
  out.sort(function (a, b) { return String(a.id).localeCompare(String(b.id)); });
  out.forEach(function (e, i) {
    e.color = EMPLOYEE_COLOR_PALETTE_[i % EMPLOYEE_COLOR_PALETTE_.length];
  });
  return out;
}

/** 16 色彩盤，柔和飽和度，便於辨識 */
const EMPLOYEE_COLOR_PALETTE_ = [
  '#fecaca', '#fed7aa', '#fef3c7', '#d9f99d',
  '#bbf7d0', '#a7f3d0', '#a5f3fc', '#bae6fd',
  '#bfdbfe', '#c7d2fe', '#ddd6fe', '#e9d5ff',
  '#f5d0fe', '#fbcfe8', '#fecdd3', '#d6d3d1'
];

function getStoresForView_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STORES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const idCol = headers.indexOf(COL.STORES.ID);
  const nameCol = headers.indexOf(COL.STORES.NAME);
  const statusCol = headers.indexOf(COL.STORES.STATUS);
  const mStart = headers.indexOf(COL.STORES.MORNING_START);
  const mEnd = headers.indexOf(COL.STORES.MORNING_END);
  const eStart = headers.indexOf(COL.STORES.EVENING_START);
  const eEnd = headers.indexOf(COL.STORES.EVENING_END);

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusCol] !== 'active') continue;
    out.push({
      id: data[i][idCol],
      name: data[i][nameCol],
      morning_start: String(data[i][mStart] || ''),
      morning_end: String(data[i][mEnd] || ''),
      evening_start: String(data[i][eStart] || ''),
      evening_end: String(data[i][eEnd] || '')
    });
  }
  return out;
}

function getShiftTypesForView_() {
  const ss = getAttendanceSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SHIFT_TYPES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const codeCol = headers.indexOf(COL.SHIFT_TYPES.SHIFT_CODE);
  const nameCol = headers.indexOf(COL.SHIFT_TYPES.NAME);
  const regularCol = headers.indexOf(COL.SHIFT_TYPES.IS_REGULAR);

  const out = [];
  for (let i = 1; i < data.length; i++) {
    out.push({
      code: data[i][codeCol],
      name: data[i][nameCol],
      is_regular: data[i][regularCol] === true || String(data[i][regularCol]).toUpperCase() === 'TRUE'
    });
  }
  return out;
}

function getSchedulesForMonth_(year, month) {
  const ss = getAttendanceSpreadsheet_();
  const sheetName = attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, year);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const idCol = headers.indexOf(COL.SCHEDULE.SCHEDULE_ID);
  const dateCol = headers.indexOf(COL.SCHEDULE.DATE);
  const empCol = headers.indexOf(COL.SCHEDULE.EMPLOYEE_ID);
  const storeCol = headers.indexOf(COL.SCHEDULE.STORE_ID);
  const shiftCol = headers.indexOf(COL.SCHEDULE.SHIFT_CODE);
  const startCol = headers.indexOf(COL.SCHEDULE.START_TIME);
  const endCol = headers.indexOf(COL.SCHEDULE.END_TIME);
  const payOverrideCol = headers.indexOf(COL.SCHEDULE.PAY_MODE_OVERRIDE);
  const statusCol = headers.indexOf(COL.SCHEDULE.STATUS);

  const monthStr = ('0' + month).slice(-2);
  const prefix = year + '-' + monthStr;

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][dateCol];
    const dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, TIMEZONE, 'yyyy-MM-dd')
      : String(dateVal);
    if (dateStr.indexOf(prefix) !== 0) continue;

    out.push({
      id: data[i][idCol],
      date: dateStr,
      employee_id: data[i][empCol],
      store_id: data[i][storeCol],
      shift_code: data[i][shiftCol],
      start_time: String(data[i][startCol] || ''),
      end_time: String(data[i][endCol] || ''),
      pay_override: data[i][payOverrideCol] || '',
      status: data[i][statusCol] || ''
    });
  }
  return out;
}

function getHolidaysForMonth_(year, month) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const dateCol = headers.indexOf(COL.HOLIDAYS.DATE);
  const yearCol = headers.indexOf(COL.HOLIDAYS.YEAR);
  const nameCol = headers.indexOf(COL.HOLIDAYS.NAME);
  const redCol = headers.indexOf(COL.HOLIDAYS.IS_RED);
  const lunarCol = headers.indexOf(COL.HOLIDAYS.LUNAR);
  const noteCol = headers.indexOf(COL.HOLIDAYS.NOTE);

  const monthStr = ('0' + month).slice(-2);
  const prefix = year + '-' + monthStr;

  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][yearCol]) !== String(year)) continue;
    const dateVal = data[i][dateCol];
    const dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, TIMEZONE, 'yyyy-MM-dd')
      : String(dateVal);
    if (dateStr.indexOf(prefix) !== 0) continue;
    out.push({
      date: dateStr,
      name: data[i][nameCol] || '',
      is_red: data[i][redCol] === true || String(data[i][redCol]).toUpperCase() === 'TRUE',
      lunar: lunarCol >= 0 ? String(data[i][lunarCol] || '') : '',
      note: noteCol >= 0 ? String(data[i][noteCol] || '') : ''
    });
  }
  return out;
}

/**
 * 由農曆判斷該日是否拜土地公。
 *
 * 規則：農曆「初二」或「十六」都拜，三個例外不拜：
 *   - 正月初二（過年）
 *   - 七月初二（鬼月，含閏七月）
 *   - 七月十六（中元，含閏七月）
 *
 * 農曆字串格式來自 Intl.DateTimeFormat 例如「正月1日」「四月2日」「七月16日」。
 * day 2 = 初二；day 16 = 十六。
 */
function getDayNote_(lunar) {
  if (!lunar) return '';
  const is2  = /月2日$/.test(lunar);   // 月後接 "2日" 結尾 → 初二
  const is16 = /月16日$/.test(lunar);  // 同理 → 十六
  if (!is2 && !is16) return '';
  if (lunar === '正月2日') return '';        // 元月初二
  if (/七月2日$/.test(lunar)) return '';     // 七月初二（含閏七月）
  if (/七月16日$/.test(lunar)) return '';    // 七月十六（含閏七月）
  return '拜土地公';
}

/**
 * 從 holidays 取出每日備註 map。
 * 優先級：admin 人工備註（COL.HOLIDAYS.NOTE） > 自動拜土地公 > 無。
 * 人工備註輸入空字串就會回退到自動。
 */
function buildDayNotesForMonth_(holidays) {
  const notes = {};
  holidays.forEach(function (h) {
    if (h.note) {
      notes[h.date] = h.note;
    } else {
      const auto = getDayNote_(h.lunar);
      if (auto) notes[h.date] = auto;
    }
  });
  return notes;
}

/**
 * 新增排班（前端 modal 呼叫）。
 *
 * payload: { date, employee_id, store_id, shift_code,
 *            start_time?, end_time?, status?,
 *            pay_mode_override?, hourly_rate_override? }
 *
 * 同一員工同日不同班別（如晚班+大夜連續班）允許；
 * 同店同班同員工同日才會被 schedule_id 唯一性擋下。
 */
function createSchedule(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('payload 必填');
  if (!payload.date || !/^\d{4}-\d{2}-\d{2}$/.test(payload.date)) throw new Error('日期格式錯誤');
  if (!payload.employee_id) throw new Error('員工必填');
  if (!payload.store_id) throw new Error('門市必填');
  if (!payload.shift_code) throw new Error('班別必填');

  // 速度優化：信任 client 傳來的 employee_id / store_id 已是 active 的（已從 DATA 帶下來）
  // 不再做 findEmployeeById_ / findStoreById_ 全表掃描

  const year = parseInt(payload.date.substring(0, 4), 10);
  const ass = getAttendanceSpreadsheet_();
  const sheet = ass.getSheetByName(attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, year));
  if (!sheet) throw new Error('排班_' + year + ' 分頁不存在，請執行 CreateAttendanceSheetsForNextYear');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idColIdx = headers.indexOf(COL.SCHEDULE.SCHEDULE_ID);

    const id = 'SCH-' + payload.date.replace(/-/g, '') + '-' +
               payload.employee_id + '-' + payload.shift_code;

    // 速度優化：只讀 ID 欄做唯一性檢查，不讀全表
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === id) {
          throw new Error(payload.employee_id + ' 在 ' + payload.date + ' 的此班次已存在');
        }
      }
    }

    const now = new Date();
    const me = (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
    const status = payload.status || '草稿';

    const newRow = headers.map(function () { return ''; });
    newRow[idColIdx] = id;
    newRow[headers.indexOf(COL.SCHEDULE.DATE)] = payload.date;
    newRow[headers.indexOf(COL.SCHEDULE.EMPLOYEE_ID)] = payload.employee_id;
    newRow[headers.indexOf(COL.SCHEDULE.STORE_ID)] = payload.store_id;
    newRow[headers.indexOf(COL.SCHEDULE.SHIFT_CODE)] = payload.shift_code;
    newRow[headers.indexOf(COL.SCHEDULE.START_TIME)] = payload.start_time || '';
    newRow[headers.indexOf(COL.SCHEDULE.END_TIME)] = payload.end_time || '';
    newRow[headers.indexOf(COL.SCHEDULE.PAY_MODE_OVERRIDE)] = payload.pay_mode_override || '';
    newRow[headers.indexOf(COL.SCHEDULE.HOURLY_RATE_OVERRIDE)] = payload.hourly_rate_override || '';
    newRow[headers.indexOf(COL.SCHEDULE.STATUS)] = status;
    newRow[headers.indexOf(COL.SCHEDULE.CREATED_BY)] = me;
    newRow[headers.indexOf(COL.SCHEDULE.CREATED_AT)] = now;
    newRow[headers.indexOf(COL.SCHEDULE.UPDATED_AT)] = now;

    sheet.appendRow(newRow);
    return { ok: true, schedule_id: id };
  } finally {
    lock.releaseLock();
  }
}

/** 切換單筆排班狀態（草稿 ↔ 已確認） */
function updateScheduleStatus(scheduleId, newStatus) {
  if (!scheduleId) throw new Error('schedule_id 必填');
  if (newStatus !== '草稿' && newStatus !== '已確認') throw new Error('狀態值錯誤');
  const m = /^SCH-(\d{4})/.exec(String(scheduleId));
  if (!m) throw new Error('schedule_id 格式錯誤');
  const year = m[1];

  const ass = getAttendanceSpreadsheet_();
  const sheet = ass.getSheetByName(attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, year));
  if (!sheet) throw new Error('排班_' + year + ' 分頁不存在');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idColIdx = headers.indexOf(COL.SCHEDULE.SCHEDULE_ID);
    const statusColIdx = headers.indexOf(COL.SCHEDULE.STATUS);
    const updatedColIdx = headers.indexOf(COL.SCHEDULE.UPDATED_AT);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('找不到排班：' + scheduleId);
    const ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === scheduleId) {
        sheet.getRange(i + 2, statusColIdx + 1).setValue(newStatus);
        sheet.getRange(i + 2, updatedColIdx + 1).setValue(new Date());
        return { ok: true, status: newStatus };
      }
    }
    throw new Error('找不到排班：' + scheduleId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 批次：把指定月（可選店篩選）的所有「草稿」一次轉「已確認」。
 * 用 batch read/write 加速：只動 status + updated_at 兩欄。
 */
function confirmAllDrafts(year, month, storeFilter) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  if (!y || !m) throw new Error('year/month 必填');

  const ass = getAttendanceSpreadsheet_();
  const sheet = ass.getSheetByName(attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, y));
  if (!sheet) throw new Error('排班_' + y + ' 分頁不存在');

  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { ok: true, count: 0 };
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dateColIdx = headers.indexOf(COL.SCHEDULE.DATE);
    const storeColIdx = headers.indexOf(COL.SCHEDULE.STORE_ID);
    const statusColIdx = headers.indexOf(COL.SCHEDULE.STATUS);
    const updatedColIdx = headers.indexOf(COL.SCHEDULE.UPDATED_AT);

    // 一次讀整段範圍，記憶體修改後再一次寫回
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const data = dataRange.getValues();
    const monthStr = ('0' + m).slice(-2);
    const prefix = y + '-' + monthStr;
    const now = new Date();
    let count = 0;

    for (let i = 0; i < data.length; i++) {
      const dv = data[i][dateColIdx];
      const ds = (dv instanceof Date)
        ? Utilities.formatDate(dv, TIMEZONE, 'yyyy-MM-dd')
        : String(dv);
      if (ds.indexOf(prefix) !== 0) continue;
      if (storeFilter && data[i][storeColIdx] !== storeFilter) continue;
      if (data[i][statusColIdx] !== '草稿') continue;
      data[i][statusColIdx] = '已確認';
      data[i][updatedColIdx] = now;
      count++;
    }
    if (count > 0) dataRange.setValues(data);
    return { ok: true, count: count };
  } finally {
    lock.releaseLock();
  }
}

/** 刪除排班 — 只讀 ID 欄做查找，不讀全表 */
function deleteSchedule(scheduleId) {
  if (!scheduleId) throw new Error('schedule_id 必填');
  const m = /^SCH-(\d{4})/.exec(String(scheduleId));
  if (!m) throw new Error('schedule_id 格式錯誤');
  const year = m[1];

  const ass = getAttendanceSpreadsheet_();
  const sheet = ass.getSheetByName(attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, year));
  if (!sheet) throw new Error('排班_' + year + ' 分頁不存在');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const idColIdx = headers.indexOf(COL.SCHEDULE.SCHEDULE_ID);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) throw new Error('找不到排班：' + scheduleId);
    const ids = sheet.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === scheduleId) {
        sheet.deleteRow(i + 2);
        return { ok: true };
      }
    }
    throw new Error('找不到排班：' + scheduleId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 載入當月預設排班 — 把「預設排班」分頁的每店每班員工，整月填入指定年月的排班表。
 * 已存在的排班（同 schedule_id）不會被覆寫。
 *
 * @param {number} year
 * @param {number} month
 * @param {string} storeFilter 空字串 = 全部 / 'LUCKY' / 'BINGO' / 'MONEY'
 * @param {string} status '草稿' or '已確認'
 */
function bulkLoadDefaults(year, month, storeFilter, status) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  if (!y || !m) throw new Error('year/month 必填');

  const ass = getAttendanceSpreadsheet_();
  const defSheet = ass.getSheetByName(SHEET_NAMES.DEFAULT_SCHEDULE);
  if (!defSheet) throw new Error('預設排班 分頁不存在，請先執行 MigrateAddDefaultScheduleSheet');

  const defData = defSheet.getDataRange().getValues();
  if (defData.length < 2) {
    return { ok: true, created: 0, skipped: 0, message: '預設排班 分頁為空，請先填入每店每班的固定員工' };
  }
  const defHeaders = defData[0];
  const dStoreCol = defHeaders.indexOf(COL.DEFAULT_SCHEDULE.STORE_ID);
  const dShiftCol = defHeaders.indexOf(COL.DEFAULT_SCHEDULE.SHIFT_CODE);
  const dEmpCol = defHeaders.indexOf(COL.DEFAULT_SCHEDULE.EMPLOYEE_ID);

  const defaults = [];
  for (let i = 1; i < defData.length; i++) {
    const sid = String(defData[i][dStoreCol] || '').trim();
    const sc = String(defData[i][dShiftCol] || '').trim();
    const eid = String(defData[i][dEmpCol] || '').trim();
    if (!sid || !sc || !eid) continue;
    if (storeFilter && sid !== storeFilter) continue;
    defaults.push({ store_id: sid, shift_code: sc, employee_id: eid });
  }
  if (defaults.length === 0) {
    return { ok: true, created: 0, skipped: 0, message: '無符合篩選的預設員工' };
  }

  const schedSheet = ass.getSheetByName(attendanceSheetName_(ATTENDANCE_PREFIX.SCHEDULE, y));
  if (!schedSheet) throw new Error('排班_' + y + ' 分頁不存在');

  const lock = LockService.getScriptLock();
  lock.waitLock(8000);
  try {
    const existingData = schedSheet.getDataRange().getValues();
    const headers = existingData[0];
    const idColIdx = headers.indexOf(COL.SCHEDULE.SCHEDULE_ID);
    const existingIds = {};
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][idColIdx]) existingIds[existingData[i][idColIdx]] = true;
    }

    // Stores map for default times
    const stSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STORES);
    const stData = stSheet.getDataRange().getValues();
    const stHeaders = stData[0];
    const stIdCol = stHeaders.indexOf(COL.STORES.ID);
    const storesMap = {};
    for (let i = 1; i < stData.length; i++) {
      const sid = stData[i][stIdCol];
      const o = {};
      stHeaders.forEach(function (h, j) { o[h] = stData[i][j]; });
      storesMap[sid] = o;
    }

    const days = new Date(y, m, 0).getDate();
    const monthStr = ('0' + m).slice(-2);
    const now = new Date();
    const me = (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
    const cleanStatus = status || '草稿';

    const newRows = [];
    let skipped = 0;
    for (let d = 1; d <= days; d++) {
      const dateStr = y + '-' + monthStr + '-' + ('0' + d).slice(-2);
      defaults.forEach(function (def) {
        const id = 'SCH-' + dateStr.replace(/-/g, '') + '-' + def.employee_id + '-' + def.shift_code;
        if (existingIds[id]) { skipped++; return; }
        existingIds[id] = true;

        const store = storesMap[def.store_id];
        let st = '', en = '';
        if (store) {
          if (def.shift_code === 'MORNING') {
            st = String(store[COL.STORES.MORNING_START] || '');
            en = String(store[COL.STORES.MORNING_END] || '');
          } else if (def.shift_code === 'EVENING') {
            st = String(store[COL.STORES.EVENING_START] || '');
            en = String(store[COL.STORES.EVENING_END] || '');
          }
        }

        const row = headers.map(function () { return ''; });
        row[idColIdx] = id;
        row[headers.indexOf(COL.SCHEDULE.DATE)] = dateStr;
        row[headers.indexOf(COL.SCHEDULE.EMPLOYEE_ID)] = def.employee_id;
        row[headers.indexOf(COL.SCHEDULE.STORE_ID)] = def.store_id;
        row[headers.indexOf(COL.SCHEDULE.SHIFT_CODE)] = def.shift_code;
        row[headers.indexOf(COL.SCHEDULE.START_TIME)] = st;
        row[headers.indexOf(COL.SCHEDULE.END_TIME)] = en;
        row[headers.indexOf(COL.SCHEDULE.STATUS)] = cleanStatus;
        row[headers.indexOf(COL.SCHEDULE.CREATED_BY)] = me;
        row[headers.indexOf(COL.SCHEDULE.CREATED_AT)] = now;
        row[headers.indexOf(COL.SCHEDULE.UPDATED_AT)] = now;
        newRows.push(row);
      });
    }

    if (newRows.length > 0) {
      const startRow = schedSheet.getLastRow() + 1;
      schedSheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
    }
    return { ok: true, created: newRows.length, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

function findEmployeeById_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.EMPLOYEES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const idCol = headers.indexOf(COL.EMPLOYEES.ID);
  const nameCol = headers.indexOf(COL.EMPLOYEES.NAME);
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      return { id: data[i][idCol], name: data[i][nameCol] };
    }
  }
  return null;
}

function findStoreById_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.STORES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0];
  const obj = null;
  const idCol = headers.indexOf(COL.STORES.ID);
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      const o = {};
      headers.forEach(function (h, j) { o[h] = data[i][j]; });
      return o;
    }
  }
  return null;
}

/**
 * 儲存某日備註（admin 從排班視圖編輯時呼叫）。
 * 若該日在 Holidays 分頁尚無紀錄，會新增一列。
 */
function saveDayNote(date, note) {
  if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(String(date))) {
    throw new Error('日期格式錯誤：' + date);
  }
  const cleanNote = String(note || '').trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) throw new Error('Holidays 分頁不存在');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const dateCol = headers.indexOf(COL.HOLIDAYS.DATE);
    const yearCol = headers.indexOf(COL.HOLIDAYS.YEAR);
    const noteCol = headers.indexOf(COL.HOLIDAYS.NOTE);
    if (noteCol < 0) throw new Error('Holidays 缺少「備註」欄位，請先執行 MigrateAddNoteColumn');

    for (let i = 1; i < data.length; i++) {
      const dv = data[i][dateCol];
      const ds = (dv instanceof Date)
        ? Utilities.formatDate(dv, TIMEZONE, 'yyyy-MM-dd')
        : String(dv);
      if (ds === date) {
        sheet.getRange(i + 1, noteCol + 1).setValue(cleanNote);
        return { ok: true, date: date, note: cleanNote };
      }
    }
    // 該日不在表中（極少見：紅字表為空或同步尚未完成）→ append
    const year = parseInt(date.substring(0, 4), 10);
    const newRow = headers.map(function () { return ''; });
    newRow[dateCol] = date;
    newRow[yearCol] = year;
    newRow[noteCol] = cleanNote;
    const sourceCol = headers.indexOf(COL.HOLIDAYS.SOURCE);
    if (sourceCol >= 0) newRow[sourceCol] = 'manual';
    sheet.appendRow(newRow);
    return { ok: true, date: date, note: cleanNote, appended: true };
  } finally {
    lock.releaseLock();
  }
}

