/**
 * 安裝定時觸發器：
 *   runWeeklyBackup       — 每週日 03:00 跑備份
 *   runHolidaysSyncAuto   — 每月 1 號 04:00 啟動（函式內自我過濾，僅 1/7 月實際同步）
 *
 * 手動執行 SetupAllTriggers 一次即可，會自動清掉舊的同名 trigger 再重建，避免重複。
 */

function SetupAllTriggers() {
  const ours = ['runWeeklyBackup', 'runHolidaysSyncAuto'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (ours.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runWeeklyBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  ScriptApp.newTrigger('runHolidaysSyncAuto')
    .timeBased()
    .onMonthDay(1)
    .atHour(4)
    .create();

  SpreadsheetApp.getUi().alert(
    '已安裝排程 ✓\n\n' +
    '• runWeeklyBackup（每週日 03:00 自動備份）\n' +
    '• runHolidaysSyncAuto（每月 1 號 04:00 觸發，僅 1 / 7 月實際拉資料）\n\n' +
    '可在左欄「觸發條件」面板查看 / 停用。'
  );
}

function ListTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert('目前沒有任何 trigger。\n\n執行 SetupAllTriggers 安裝排程。');
    return;
  }
  const list = triggers.map(function (t) {
    return '• ' + t.getHandlerFunction() + ' （' + t.getEventType() + '）';
  }).join('\n');
  SpreadsheetApp.getUi().alert('目前 trigger：\n\n' + list);
}
