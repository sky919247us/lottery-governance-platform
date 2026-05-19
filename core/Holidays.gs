/**
 * 紅字自動同步：從 imsyuan/taiwan-holidays（jsDelivr CDN）拉 JSON 寫進 Holidays 分頁。
 *
 * 資料源：https://cdn.jsdelivr.net/gh/imsyuan/taiwan-holidays/data/{year}.json
 * Trigger：每月 1 號 04:00 自動執行 runHolidaysSyncAuto，但只在 1 月 / 7 月實際拉取
 *         （1 月：當年資料、明年資料；7 月：明年資料）
 *
 * 失敗處理：連續 3 次失敗 → 自動關閉 holidays_sync_enabled + 寄信通知 admin。
 * 手動覆寫：source='manual' 的列永不被自動同步覆蓋。
 */

const HOLIDAYS_API_BASE = 'https://cdn.jsdelivr.net/gh/imsyuan/taiwan-holidays/data/';

function runHolidaysSyncAuto() {
  const month = parseInt(Utilities.formatDate(new Date(), TIMEZONE, 'M'), 10);
  if (month !== 1 && month !== 7) return;

  const enabled = String(getSetting_('holidays_sync_enabled')).toLowerCase();
  if (enabled === 'false') {
    console.log('holidays_sync_enabled = false, skipping');
    return;
  }

  const year = new Date().getFullYear();
  // 1 月：同步當年（剛開始）+ 明年（若上游已發布）
  // 7 月：同步明年（這時上游通常剛發布下年度）
  try {
    if (month === 1) syncHolidaysForYear_(year);
    syncHolidaysForYear_(year + 1);
    setSetting_('holidays_sync_failures', '0');
  } catch (e) {
    handleHolidaysSyncFailure_(year, e);
  }
}

/** 手動同步指定年份 */
function SyncHolidays2026() { manualSync_(2026); }
function SyncHolidays2027() { manualSync_(2027); }

function manualSync_(year) {
  try {
    syncHolidaysForYear_(year);
    SpreadsheetApp.getUi().alert(year + ' 年紅字同步完成 ✓\n\n請檢查 Holidays 分頁。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('同步 ' + year + ' 失敗：\n\n' + (e.message || e));
  }
}

function syncHolidaysForYear_(year) {
  const url = HOLIDAYS_API_BASE + year + '.json';
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('HTTP ' + resp.getResponseCode() + ' 從 ' + url);
  }
  const data = JSON.parse(resp.getContentText());
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error(year + '.json 內容空或非陣列');
  }
  writeHolidaysToSheet_(year, data);
}

function writeHolidaysToSheet_(year, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) throw new Error('Holidays 分頁不存在，請先執行 MigrateAddHolidays 或重跑 Bootstrap');

  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const dateColIdx = headers.indexOf(COL.HOLIDAYS.DATE);
  const yearCol = headers.indexOf(COL.HOLIDAYS.YEAR);
  const noteCol = headers.indexOf(COL.HOLIDAYS.NOTE);
  const sourceCol = headers.indexOf(COL.HOLIDAYS.SOURCE);

  // 同步前先保存當年所有 備註 內容（key: date, value: note）
  const preservedNotes = {};
  if (noteCol >= 0) {
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][yearCol]) !== String(year)) continue;
      const dv = allData[i][dateColIdx];
      const ds = (dv instanceof Date)
        ? Utilities.formatDate(dv, TIMEZONE, 'yyyy-MM-dd')
        : String(dv);
      if (allData[i][noteCol]) preservedNotes[ds] = allData[i][noteCol];
    }
  }

  // 從底向上刪除這年的 auto 列
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][yearCol]) === String(year) && allData[i][sourceCol] === 'auto') {
      sheet.deleteRow(i + 1);
    }
  }

  const remaining = sheet.getDataRange().getValues();
  const manualDates = {};
  for (let i = 1; i < remaining.length; i++) {
    if (String(remaining[i][yearCol]) === String(year) && remaining[i][sourceCol] === 'manual') {
      manualDates[String(remaining[i][0])] = true;
    }
  }

  // 組新列：[date, year, name, is_red, lunar, note, source]
  const rows = [];
  data.forEach(function (d) {
    const dateStr = d.date.substring(0, 4) + '-' + d.date.substring(4, 6) + '-' + d.date.substring(6, 8);
    if (manualDates[dateStr]) return;
    const lunar = (d.lunar && d.lunar.date) ? d.lunar.date : '';
    rows.push([
      dateStr,
      parseInt(d.date.substring(0, 4), 10),
      d.description || '',
      d.isHoliday === true,
      lunar,
      preservedNotes[dateStr] || '',
      'auto'
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }
}

/**
 * 用 V8 Intl 把國曆字串 (yyyy-MM-dd) 轉成台灣農曆顯示，例如 "四月初二"。
 * 若 Apps Script V8 沒帶 chinese calendar locale，會 throw / 回空字串。
 */
function gregorianToLunarTW_(dateStr) {
  if (!dateStr) return '';
  try {
    const d = new Date(dateStr + 'T12:00:00+08:00');
    const fmt = new Intl.DateTimeFormat('zh-TW-u-ca-chinese', {
      month: 'long', day: 'numeric'
    });
    return fmt.format(d);
  } catch (e) {
    return '';
  }
}

/** 測試 Intl 農曆轉換是否能用 — 5 筆已知對照，跑完看結果是否正確 */
function TestLunar() {
  const cases = [
    ['2026-01-01', '預期：(去年)冬月 / 十一月 十三 之類'],
    ['2026-02-17', '預期：正月初一（2026 春節）'],
    ['2026-02-18', '預期：正月初二'],
    ['2026-05-18', '預期：四月初二（拜土地公）'],
    ['2026-05-19', '預期：四月初三'],
    ['2026-08-28', '預期：七月十六（拜土地公例外）']
  ];
  let lines = ['Intl.DateTimeFormat 農曆轉換測試\n'];
  cases.forEach(function (c) {
    const lunar = gregorianToLunarTW_(c[0]);
    lines.push(c[0] + ' → ' + (lunar || '(空字串 / 不支援)') + '\n' + c[1]);
  });
  SpreadsheetApp.getUi().alert(lines.join('\n\n'));
}

/**
 * 重新計算 Holidays 表中某年所有列的「農曆」欄（用 Intl，不打外部 API）。
 * 只動 source=auto 的列；manual 列不動。
 */
function ResyncLunar2026() { resyncLunarForYear_(2026); }
function ResyncLunar2027() { resyncLunarForYear_(2027); }

function resyncLunarForYear_(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Holidays 分頁不存在');
    return;
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateCol = headers.indexOf(COL.HOLIDAYS.DATE);
  const yearCol = headers.indexOf(COL.HOLIDAYS.YEAR);
  const lunarCol = headers.indexOf(COL.HOLIDAYS.LUNAR);
  const sourceCol = headers.indexOf(COL.HOLIDAYS.SOURCE);
  if (lunarCol < 0) {
    SpreadsheetApp.getUi().alert('Holidays 缺少 農曆 欄位');
    return;
  }

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][yearCol]) !== String(year)) continue;
    if (data[i][sourceCol] !== 'auto') continue;
    const dv = data[i][dateCol];
    const ds = (dv instanceof Date)
      ? Utilities.formatDate(dv, TIMEZONE, 'yyyy-MM-dd')
      : String(dv);
    const newLunar = gregorianToLunarTW_(ds);
    if (newLunar && newLunar !== String(data[i][lunarCol] || '')) {
      sheet.getRange(i + 1, lunarCol + 1).setValue(newLunar);
      updated++;
    }
  }
  SpreadsheetApp.getUi().alert(
    year + ' 年農曆 resync 完成 ✓\n\n' +
    '更新 ' + updated + ' 列。\n\n' +
    '檢查 Holidays 分頁是否正確（如 2026-02-17 應為「正月初一」）。'
  );
}

/** 為 Core Sheet 加 留職停薪紀錄 分頁 */
function MigrateAddLeavePeriodsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET_NAMES.LEAVE_PERIODS)) {
    SpreadsheetApp.getUi().alert('留職停薪紀錄 分頁已存在');
    return;
  }
  const sheet = ss.insertSheet(SHEET_NAMES.LEAVE_PERIODS);
  const headers = [
    COL.LEAVE_PERIODS.LEAVE_ID, COL.LEAVE_PERIODS.EMPLOYEE_ID,
    COL.LEAVE_PERIODS.TYPE,
    COL.LEAVE_PERIODS.START_DATE, COL.LEAVE_PERIODS.END_DATE,
    COL.LEAVE_PERIODS.DEDUCT_TENURE,
    COL.LEAVE_PERIODS.APPROVED_BY, COL.LEAVE_PERIODS.APPROVED_AT,
    COL.LEAVE_PERIODS.NOTE
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert(
    '✓ 留職停薪紀錄 分頁已建立\n\n' +
    '欄位說明：\n' +
    '• 類型：育嬰留停 / 一般留停 / 病假留停 / 兵役 / 其他\n' +
    '• 復職日：空白表示尚在留停中\n' +
    '• 扣除年資：TRUE/FALSE，影響特休年資計算'
  );
}

/** 為已存在的 Holidays 分頁加 備註 欄位（在 農曆 和 資料來源 之間） */
function MigrateAddNoteColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Holidays 分頁不存在');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf(COL.HOLIDAYS.NOTE) >= 0) {
    SpreadsheetApp.getUi().alert('備註 欄位已存在，無需 migration。');
    return;
  }
  const sourceColIdx = headers.indexOf(COL.HOLIDAYS.SOURCE);
  if (sourceColIdx < 0) {
    SpreadsheetApp.getUi().alert('找不到 資料來源 欄位');
    return;
  }
  sheet.insertColumnBefore(sourceColIdx + 1);
  sheet.getRange(1, sourceColIdx + 1).setValue(COL.HOLIDAYS.NOTE)
    .setFontWeight('bold').setBackground('#e8f0fe');
  SpreadsheetApp.getUi().alert('✓ 已新增 備註 欄位');
}

/** 為已存在的 Holidays 分頁加 農曆 欄位（在 紅字 和 資料來源 之間） */
function MigrateAddLunarColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.HOLIDAYS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Holidays 分頁不存在');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf(COL.HOLIDAYS.LUNAR) >= 0) {
    SpreadsheetApp.getUi().alert('農曆 欄位已存在，無需 migration。');
    return;
  }
  const sourceColIdx = headers.indexOf(COL.HOLIDAYS.SOURCE);
  if (sourceColIdx < 0) {
    SpreadsheetApp.getUi().alert('找不到 資料來源 欄位，請先確認 Holidays schema 正確');
    return;
  }
  // 在 sourceColIdx 前插入新欄
  sheet.insertColumnBefore(sourceColIdx + 1);
  sheet.getRange(1, sourceColIdx + 1).setValue(COL.HOLIDAYS.LUNAR)
    .setFontWeight('bold').setBackground('#e8f0fe');
  SpreadsheetApp.getUi().alert(
    '✓ 已新增 農曆 欄位\n\n請接著執行 SyncHolidays2026 重新拉取資料填入農曆資訊。'
  );
}

function handleHolidaysSyncFailure_(year, err) {
  const fails = parseInt(getSetting_('holidays_sync_failures') || '0', 10) + 1;
  setSetting_('holidays_sync_failures', String(fails));
  console.error('Holidays sync failed for ' + year + ' (count=' + fails + '): ' + (err.message || err));

  if (fails >= 3) {
    setSetting_('holidays_sync_enabled', 'false');
    notifyAdminEmail_(
      '【彩券行系統】假日表同步已自動停用',
      '同步失敗連續 3 次，已將 Settings.holidays_sync_enabled 設為 false。\n' +
      '請手動檢查資料源是否變動，或改為手動維護 Holidays 分頁。\n\n' +
      '最後錯誤：' + (err.message || err)
    );
  }
}
