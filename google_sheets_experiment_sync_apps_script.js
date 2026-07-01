const SHEET_NAME = 'experimentDashboard';

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || '';
  try {
    const action = (e && e.parameter && e.parameter.action) || 'read';
    if (action === 'ping') {
      return jsonp_(callback, {ok: true, message: 'pong'});
    }
    if (action === 'write') {
      const payload = JSON.parse(e.parameter.payload || '{}');
      writeState_(payload);
      return jsonp_(callback, {ok: true, updatedAt: payload.updatedAt || new Date().toISOString()});
    }
    const state = readState_();
    const response = {
      ok: true,
      updatedAt: state.updatedAt || '',
      data: state.data || {experiments: []}
    };
    return jsonp_(callback, response);
  } catch (err) {
    return jsonp_(callback, {ok: false, error: String(err)});
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    if (payload.action !== 'write') {
      return json_({ok: false, error: 'unknown action'});
    }
    writeState_(payload);
    return json_({ok: true, updatedAt: payload.updatedAt || new Date().toISOString()});
  } catch (err) {
    return json_({ok: false, error: String(err)});
  }
}

function readState_() {
  const sheet = getSheet_();
  const raw = sheet.getRange(1, 1).getValue();
  if (!raw) return {updatedAt: '', data: {experiments: []}};
  try {
    return JSON.parse(raw);
  } catch (err) {
    return {updatedAt: '', data: {experiments: []}};
  }
}

function writeState_(payload) {
  const sheet = getSheet_();
  const state = {
    updatedAt: payload.updatedAt || new Date().toISOString(),
    data: payload.data || {experiments: []}
  };
  sheet.getRange(1, 1).setValue(JSON.stringify(state));
  sheet.getRange(1, 2).setValue(state.updatedAt);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트에서 확장 프로그램 > Apps Script로 만든 스크립트가 아닙니다.');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, obj) {
  const body = callback
    ? `${callback}(${JSON.stringify(obj)});`
    : JSON.stringify(obj);
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
