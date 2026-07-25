/**
 * Peter’s Lab Poodle Campaign – Mibang Dashboard
 * Google Sheet에서 [확장 프로그램] → [Apps Script]를 연 뒤 이 코드를 전체 교체합니다.
 * 배포: [배포] → [배포 관리] → 연필 아이콘 → 새 버전 → 배포
 */

const SPREADSHEET_ID = '1kuzNyMSObduViqgl0uVYcKg5cJs7in1RR1budvcXj84';

const SHEETS = {
  DM: 'DM_LIST',
  TASKS: 'TASK_STATUS',
  LOG: 'DAILY_LOG'
};

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return json_({
      ok: true,
      service: 'mibang-campaign-sync',
      spreadsheet: ss.getName(),
      dmRows: Math.max(0, ss.getSheetByName(SHEETS.DM).getLastRow() - 2),
      taskRows: Math.max(0, ss.getSheetByName(SHEETS.TASKS).getLastRow() - 2),
      logRows: Math.max(0, ss.getSheetByName(SHEETS.LOG).getLastRow() - 2)
    });
  } catch (error) {
    return json_({ ok: false, error: errorMessage_(error) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    validatePayload_(payload);

    // 웹앱 실행 시 getActiveSpreadsheet()가 비어 있을 수 있으므로 ID로 직접 엽니다.
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const syncedAt = new Date();

    replaceDmList_(ss.getSheetByName(SHEETS.DM), payload.dmList || [], syncedAt);
    upsertTaskStatus_(ss.getSheetByName(SHEETS.TASKS), payload.tasks || [], syncedAt);

    // 자동 저장 때는 일일 로그를 쌓지 않고, 일일 마감 버튼을 누른 경우에만 기록합니다.
    if (payload.mode === 'daily') {
      appendDailyLog_(ss.getSheetByName(SHEETS.LOG), payload, syncedAt);
    }

    SpreadsheetApp.flush();

    return json_({
      ok: true,
      syncId: payload.syncId,
      mode: payload.mode || 'autosave',
      dmRows: (payload.dmList || []).length,
      taskRows: (payload.tasks || []).length,
      syncedAt: syncedAt.toISOString()
    });
  } catch (error) {
    console.error(error);
    return json_({ ok: false, error: errorMessage_(error) });
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(payload) {
  if (!payload || !payload.syncId) throw new Error('syncId is required');
  if (!payload.date) throw new Error('date is required');
  if (!Array.isArray(payload.dmList)) throw new Error('dmList must be an array');
  if (!Array.isArray(payload.tasks)) throw new Error('tasks must be an array');
}

function ensureSheet_(sheet, name) {
  if (!sheet) throw new Error(name + ' sheet not found');
}

function replaceDmList_(sheet, list, syncedAt) {
  ensureSheet_(sheet, SHEETS.DM);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, 11).clearContent();
  }

  if (!list.length) return;

  const rows = list.map(item => [
    String(item.id || item.account || Utilities.getUuid()),
    item.account || '',
    item.dog || '',
    item.url || '',
    item.residence || '',
    item.found || '',
    item.sent || '',
    item.status || 'not_sent',
    item.follow || '',
    item.note || '',
    syncedAt
  ]);

  sheet.getRange(3, 1, rows.length, 11).setValues(rows);
}

function upsertTaskStatus_(sheet, tasks, syncedAt) {
  ensureSheet_(sheet, SHEETS.TASKS);

  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 3
    ? sheet.getRange(3, 1, lastRow - 2, 2).getValues()
    : [];

  const rowByKey = {};
  existing.forEach((row, index) => {
    const key = String(row[0] || '') + '|' + String(row[1] || '');
    if (key !== '|') rowByKey[key] = index + 3;
  });

  tasks.forEach(task => {
    const key = String(task.date || '') + '|' + String(task.key || '');
    const row = [
      task.date || '',
      task.key || '',
      task.ko || '',
      task.ja || '',
      task.completed === true,
      task.completed ? syncedAt : '',
      syncedAt
    ];

    const targetRow = rowByKey[key] || sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    rowByKey[key] = targetRow;
  });
}

function appendDailyLog_(sheet, payload, syncedAt) {
  ensureSheet_(sheet, SHEETS.LOG);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    const syncIds = sheet.getRange(3, 13, lastRow - 2, 1).getValues().flat().map(String);
    if (syncIds.includes(String(payload.syncId))) return;
  }

  const summary = payload.summary || {};
  const tasks = payload.tasks || [];
  const completedCount = tasks.filter(task => task.completed === true).length;

  sheet.appendRow([
    payload.date || '',
    syncedAt,
    payload.owner || 'Mibang',
    completedCount,
    tasks.length,
    Number(summary.dmCandidates || 0),
    Number(summary.dmToday || 0),
    Number(summary.dmTotal || 0),
    Number(summary.replies || 0),
    Number(summary.applications || 0),
    summary.contentPosted === true,
    payload.note || '',
    payload.syncId
  ]);
}

function errorMessage_(error) {
  return String(error && error.message ? error.message : error);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
