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
  const yearCol = headers.indexOf(COL.HOLIDAYS.YEAR);
  const sourceCol = headers.indexOf(COL.HOLIDAYS.SOURCE);

  // 從底向上刪除這年的 auto 列（保留 manual 覆寫）
  for (let i = allData.length - 1; i >= 1; i--) {
    if (String(allData[i][yearCol]) === String(year) && allData[i][sourceCol] === 'auto') {
      sheet.deleteRow(i + 1);
    }
  }

  // 取出該年所有 manual 覆寫的日期，避免重複插入
  const remaining = sheet.getDataRange().getValues();
  const manualDates = {};
  for (let i = 1; i < remaining.length; i++) {
    if (String(remaining[i][yearCol]) === String(year) && remaining[i][sourceCol] === 'manual') {
      manualDates[String(remaining[i][0])] = true;
    }
  }

  // 組新列：[date, year, name, is_red, source]
  const rows = [];
  data.forEach(function (d) {
    const dateStr = d.date.substring(0, 4) + '-' + d.date.substring(4, 6) + '-' + d.date.substring(6, 8);
    if (manualDates[dateStr]) return; // 該日已有 manual 覆寫，跳過
    rows.push([dateStr, parseInt(d.date.substring(0, 4), 10), d.description || '', d.isHoliday === true, 'auto']);
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
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
