/**
 * さわでん 現場管理アプリ  Google Apps Script
 *
 * このファイルは Apps Script エディタに貼り付けるためのものです。
 * 貼り付けたあと、必ず「デプロイ → デプロイを管理 → 編集 → バージョン: 新バージョン」
 * で再デプロイしてください（再デプロイしないと反映されません）。
 * ウェブアプリの URL は変わりません。
 *
 * シートの列構成
 *   sites     … A:現場名
 *   workers   … A:作業者名
 *   rates     … A:作業者名 B:単価
 *   reports   … A:日付 B:現場名 C:作業者 D:作業時間 E:作業内容
 *                F:駐車場代 G:高速代 H:備考 I:ID J:カレンダーID
 *   schedules … A:ID B:日付 C:現場名 D:作業者 E:メモ
 *
 * I列（ID）と J列（カレンダーID）は後から追加した列です。
 * 既存の日報には ID が無いので、reports を読み込んだときに自動で採番します（初回だけ）。
 */

const SHEET_ID = '1rH_3zr9JV6kin6sC71X21OObh--ZLr8qVRpeTDCZYWE';

const REPORT_HEADER = ['日付', '現場名', '作業者', '作業時間', '作業内容', '駐車場代', '高速代', '備考', 'ID', 'カレンダーID'];
const REPORT_ID_COL = 9;     // I列
const REPORT_EVENT_COL = 10; // J列
const SCHEDULE_HEADER = ['ID', '日付', '現場名', '作業者', 'メモ'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    if (action === 'addSite') return addItem('sites', data.value);
    else if (action === 'deleteSite') return deleteItem('sites', data.value);
    else if (action === 'addWorker') return addItem('workers', data.value);
    else if (action === 'deleteWorker') return deleteItem('workers', data.value);
    else if (action === 'saveReport') return saveReport(data);
    // 以前のバージョンは POST でも日報を削除できたので、そのまま受けられるように残す
    // （古いページが端末にキャッシュされていても動くように）
    else if (action === 'deleteReport') {
      return deleteReport(SpreadsheetApp.openById(SHEET_ID), data.id, data.date, data.site);
    }
    else if (action === 'addSchedule') return addSchedule(data);
    else if (action === 'addSchedules') return addSchedules(data.list);
    else if (action === 'deleteSchedule') return deleteSchedule(data.id);
    else return respond({ status: 'error', message: 'unknown action' });
  } catch(err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    const type = e.parameter.type;
    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (type === 'sites') {
      const sheet = getOrCreateSheet(ss, 'sites');
      const values = sheet.getDataRange().getValues().flat().filter(v => v !== '');
      return respond({ status: 'ok', data: values });
    } else if (type === 'workers') {
      const sheet = getOrCreateSheet(ss, 'workers');
      const values = sheet.getDataRange().getValues().flat().filter(v => v !== '');
      return respond({ status: 'ok', data: values });
    } else if (type === 'reports') {
      const sheet = getOrCreateSheet(ss, 'reports');
      ensureReportSheet(sheet);
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return respond({ status: 'ok', data: [] });
      ensureReportIds(sheet, rows);
      const data = rows.slice(1).map(row => ({
        id: String(row[REPORT_ID_COL - 1] || ''),
        date: toDateStr(row[0]),
        site: row[1],
        workers: row[2],
        timeStr: row[3],
        tasks: row[4],
        parking: row[5],
        highway: row[6],
        notes: row[7]
      }));
      return respond({ status: 'ok', data: data });
    } else if (type === 'schedules') {
      const sheet = getOrCreateSheet(ss, 'schedules');
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return respond({ status: 'ok', data: [] });
      const data = rows.slice(1).filter(row => row[0]).map(row => ({
        id: String(row[0]),
        date: toDateStr(row[1]),
        site: row[2],
        workers: row[3] ? String(row[3]).split('、').filter(v => v !== '') : [],
        memo: row[4] || ''
      }));
      return respond({ status: 'ok', data: data });
    } else if (type === 'rates') {
      const sheet = getOrCreateSheet(ss, 'rates');
      const rows = sheet.getDataRange().getValues();
      const data = {};
      rows.forEach(row => { if (row[0]) data[String(row[0])] = row[1] || 0; });
      return respond({ status: 'ok', data: data });
    } else if (type === 'setRate') {
      const sheet = getOrCreateSheet(ss, 'rates');
      const worker = e.parameter.worker;
      const rate = e.parameter.rate;
      const rows = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]) === worker) {
          sheet.getRange(i + 1, 2).setValue(rate);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([worker, rate]);
      return respond({ status: 'ok' });
    } else if (type === 'deleteReport') {
      return deleteReport(ss, e.parameter.id, e.parameter.date, e.parameter.site);
    } else {
      return respond({ status: 'error', message: 'unknown type' });
    }
  } catch(err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

/* ---------- 日報 ---------- */

function saveReport(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'reports');
  ensureReportSheet(sheet);

  const id = Utilities.getUuid();
  let eventId = '';
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const start = new Date(data.date + 'T00:00:00');
    const event = calendar.createAllDayEvent(data.title, start, { description: data.detail });
    eventId = event.getId();
  } catch(err) {
    // カレンダー登録に失敗しても、日報そのものは必ず保存する
  }

  sheet.appendRow([
    data.date, data.site, data.workers, data.timeStr, data.tasks,
    data.parking, data.highway, data.notes, id, eventId
  ]);
  return respond({ status: 'ok', id: id, calendar: eventId !== '' });
}

function deleteReport(ss, id, date, site) {
  const sheet = getOrCreateSheet(ss, 'reports');
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const rowId = String(rows[i][REPORT_ID_COL - 1] || '');
    const rowDate = toDateStr(rows[i][0]);
    const rowSite = String(rows[i][1] || '');
    // ID が渡されていれば ID だけで特定する。
    // 同じ日・同じ現場の日報が2件あっても取り違えない。
    const matched = id ? (rowId === String(id)) : (rowDate === date && rowSite === site);
    if (!matched) continue;
    deleteCalendarEvent(rowDate, rowSite, String(rows[i][REPORT_EVENT_COL - 1] || ''));
    sheet.deleteRow(i + 1);
    return respond({ status: 'ok' });
  }
  return respond({ status: 'error', message: 'not found' });
}

function deleteCalendarEvent(date, site, eventId) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    // カレンダーIDが記録されていれば、その予定だけを確実に消す
    if (eventId) {
      const event = calendar.getEventById(eventId);
      if (event) { event.deleteEvent(); return; }
    }
    // 記録が無い古い日報は、これまでどおりタイトル一致で探す。
    // ただし日報1件の削除なので、消すのも1件だけにする。
    const targetDate = new Date(date + 'T00:00:00');
    const events = calendar.getEventsForDay(targetDate);
    const targetTitle = '日報 ' + site;
    for (let i = 0; i < events.length; i++) {
      if (events[i].getTitle() === targetTitle) {
        events[i].deleteEvent();
        return;
      }
    }
  } catch(err) {}
}

function ensureReportSheet(sheet) {
  if (sheet.getMaxColumns() < REPORT_HEADER.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), REPORT_HEADER.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REPORT_HEADER);
    return;
  }
  // 既存シートに ID / カレンダーID の見出しが無ければ追加する
  const header = sheet.getRange(1, 1, 1, REPORT_HEADER.length).getValues()[0];
  if (!header[REPORT_ID_COL - 1]) sheet.getRange(1, REPORT_ID_COL).setValue('ID');
  if (!header[REPORT_EVENT_COL - 1]) sheet.getRange(1, REPORT_EVENT_COL).setValue('カレンダーID');
}

// ID が無い既存の日報に採番する（初回の読み込みのときだけ書き込みが走る）
function ensureReportIds(sheet, rows) {
  const ids = [];
  let needsWrite = false;
  for (let i = 1; i < rows.length; i++) {
    let id = rows[i][REPORT_ID_COL - 1];
    if (!id) {
      id = Utilities.getUuid();
      rows[i][REPORT_ID_COL - 1] = id;
      needsWrite = true;
    }
    ids.push([id]);
  }
  if (needsWrite && ids.length > 0) {
    sheet.getRange(2, REPORT_ID_COL, ids.length, 1).setValues(ids);
  }
}

/* ---------- 予定表 ---------- */

function addSchedule(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'schedules');
  ensureScheduleSheet(sheet);
  const id = Utilities.getUuid();
  sheet.appendRow([id, data.date, data.site, joinWorkers(data.workers), data.memo || '']);
  return respond({ status: 'ok', id: id });
}

// 端末に溜まっていた予定をまとめてスプレッドシートへ移すため
function addSchedules(list) {
  if (!list || list.length === 0) return respond({ status: 'ok', count: 0 });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'schedules');
  ensureScheduleSheet(sheet);
  const rows = list.map(s => [Utilities.getUuid(), s.date, s.site, joinWorkers(s.workers), s.memo || '']);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SCHEDULE_HEADER.length).setValues(rows);
  return respond({ status: 'ok', count: rows.length });
}

function deleteSchedule(id) {
  if (!id) return respond({ status: 'error', message: 'id required' });
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, 'schedules');
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return respond({ status: 'ok' });
    }
  }
  return respond({ status: 'error', message: 'not found' });
}

function ensureScheduleSheet(sheet) {
  if (sheet.getLastRow() === 0) sheet.appendRow(SCHEDULE_HEADER);
}

function joinWorkers(workers) {
  if (Array.isArray(workers)) return workers.join('、');
  return String(workers || '');
}

/* ---------- 共通 ---------- */

function addItem(sheetName, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, sheetName);
  sheet.appendRow([value]);
  return respond({ status: 'ok' });
}

function deleteItem(sheetName, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === value) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return respond({ status: 'ok' });
}

function toDateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v || '');
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function respond(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function setup() {
  SpreadsheetApp.openById(SHEET_ID);
  CalendarApp.getDefaultCalendar();
}
