/**
 * Peter’s Lab Poodle Campaign – Mibang Dashboard
 * Google Sheet에서 [확장 프로그램] → [Apps Script]를 연 뒤 이 코드를 붙여넣습니다.
 * 배포: [배포] → [새 배포] → [웹 앱] → 실행 사용자: 나 / 액세스: 모든 사용자
 */

const SHEETS = {
  DM: 'DM_LIST',
  TASKS: 'TASK_STATUS',
  LOG: 'DAILY_LOG'
};

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'mibang-campaign-sync' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    validatePayload_(payload);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const syncedAt = new Date();

    upsertDmList_(ss.getSheetByName(SHEETS.DM), payload.dmList || [], syncedAt);
    upsertTaskStatus_(ss.getSheetByName(SHEETS.TASKS), payload.tasks || [], syncedAt);
    appendDailyLog_(ss.getSheetByName(SHEETS.LOG), payload, syncedAt);

    SpreadsheetApp.flush();
    return json_({ ok: true, syncId: payload.syncId, syncedAt: syncedAt.toISOString() });
  } catch (error) {
    return json_({ ok: false, error: String(error && error.message ? error.message : error) });
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(p) {
  if (!p || !p.syncId) throw new Error('syncId is required');
  if (!p.date) throw new Error('date is required');
}

function ensureSheet_(sheet, name) {
  if (!sheet) throw new Error(name + ' sheet not found');
}

function upsertDmList_(sheet, list, syncedAt) {
  ensureSheet_(sheet, SHEETS.DM);
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 3 ? sheet.getRange(3, 1, lastRow - 2, 1).getValues().flat() : [];
  const rowById = {};
  existing.forEach((id, i) => { if (id) rowById[String(id)] = i + 3; });

  list.forEach(item => {
    const id = String(item.id || item.account || Utilities.getUuid());
    const row = [
      id,
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
    ];
    const targetRow = rowById[id] || sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    rowById[id] = targetRow;
  });
}

function upsertTaskStatus_(sheet, tasks, syncedAt) {
  ensureSheet_(sheet, SHEETS.TASKS);
  const lastRow = sheet.getLastRow();
  const values = lastRow >= 3 ? sheet.getRange(3, 1, lastRow - 2, 2).getValues() : [];
  const rowByKey = {};
  values.forEach((r, i) => {
    const key = String(r[0] || '') + '|' + String(r[1] || '');
    if (key !== '|') rowByKey[key] = i + 3;
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
  const completedCount = tasks.filter(t => t.completed === true).length;
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
