/**
 * 每週備份：把 Core Sheet 完整複製到 Drive 備份資料夾，保留 5 年。
 *
 * Trigger 由 SetupAllTriggers 安裝（每週日 03:00）。
 * 手動測試：執行 RunBackupNow。
 */

function runWeeklyBackup() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('Backup skipped: another execution holds the lock');
    return;
  }
  try {
    backupCoreSheet_();
  } finally {
    lock.releaseLock();
  }
}

/** 手動測試備份 */
function RunBackupNow() {
  backupCoreSheet_();
  SpreadsheetApp.getUi().alert(
    '備份完成 ✓\n\n' +
    '請看 Backup_Log 分頁與 Google Drive 中「[彩券行] 備份」資料夾。'
  );
}

function backupCoreSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd_HHmm');
  const fileName = ss.getName() + ' - backup ' + timestamp;

  try {
    const folder = getOrCreateBackupFolder_();
    const file = DriveApp.getFileById(ss.getId()).makeCopy(fileName, folder);
    logBackup_(new Date(), file.getId(), 'success', '');
    pruneOldBackups_(folder);
  } catch (e) {
    logBackup_(new Date(), '', 'failed', String(e.message || e));
    notifyAdminEmail_(
      '【彩券行系統】備份失敗',
      '備份時間：' + timestamp + '\n錯誤：' + (e.message || e)
    );
    throw e;
  }
}

function getOrCreateBackupFolder_() {
  const folderId = getSetting_('backup_folder_id');
  if (folderId) {
    try {
      return DriveApp.getFolderById(String(folderId));
    } catch (e) {
      // folder was deleted, fall through to recreate
      console.warn('Stored backup folder not found, creating new one');
    }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetFile = DriveApp.getFileById(ss.getId());
  const parents = sheetFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder('[彩券行] 備份');
  setSetting_('backup_folder_id', folder.getId());
  return folder;
}

function logBackup_(when, fileId, status, note) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BACKUP_LOG);
  if (!sheet) return;
  sheet.appendRow([when, fileId, status, note]);
}

/** 刪除超過 5 年的備份檔（會計憑證年限） */
function pruneOldBackups_(folder) {
  const cutoff = new Date();
  cutoff.setFullYear(cutoff.getFullYear() - 5);
  const files = folder.getFiles();
  let trashed = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getDateCreated() < cutoff) {
      f.setTrashed(true);
      trashed++;
    }
  }
  if (trashed > 0) {
    console.log('Pruned ' + trashed + ' old backup files');
  }
}
