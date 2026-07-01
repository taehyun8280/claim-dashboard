const SHEET_NAME = 'experimentDashboard';

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';

  try {
    const action = params.action || 'read';

    if (action === 'ping') {
      return jsonp_(callback, {ok: true, message: 'pong'});
    }

    if (action === 'write') {
      const payload = JSON.parse(params.payload || '{}');
      writeState_(payload);
      return jsonp_(callback, {ok: true, updatedAt: payload.updatedAt || new Date().toISOString()});
    }

    const state = readState_();
    return jsonp_(callback, {
      ok: true,
      updatedAt: state.updatedAt || '',
      data: state.data || {experiments: []}
    });
  } catch (err) {
    return jsonp_(callback, {ok: false, error: String(err)});
  }
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';

  try {
    const body = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(body);

    if (payload.action !== 'write') {
      return jsonp_(callback, {ok: false, error: 'unknown action'});
    }

    writeState_(payload);
    return jsonp_(callback, {ok: true, updatedAt: payload.updatedAt || new Date().toISOString()});
  } catch (err) {
    return jsonp_(callback, {ok: false, error: String(err)});
  }
}

function readState_() {
  const sheet = getSheet_();
  const raw = sheet.getRange(1, 1).getValue();
  if (!raw) return {updatedAt: '', data: {experiments: []}};

  try {
    const state = JSON.parse(raw);
    if (!state || typeof state !== 'object') throw new Error('invalid state');
    return {
      updatedAt: state.updatedAt || '',
      data: state.data || {experiments: []}
    };
  } catch (err) {
    return {updatedAt: '', data: {experiments: []}};
  }
}

function writeState_(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sheet = getSheet_();
    const current = readState_();
    const currentTime = Date.parse(current.updatedAt || '') || 0;
    const nextTime = Date.parse(payload.updatedAt || '') || Date.now();

    if (currentTime > nextTime) return;

    const state = {
      updatedAt: payload.updatedAt || new Date().toISOString(),
      data: payload.data || {experiments: []}
    };

    sheet.getRange(1, 1).setValue(JSON.stringify(state));
    sheet.getRange(1, 2).setValue(state.updatedAt);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('구글 시트에서 확장 프로그램 > Apps Script로 만든 스크립트가 아닙니다.');
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function jsonp_(callback, obj) {
  const body = callback
    ? `${callback}(${JSON.stringify(obj)});`
    : JSON.stringify(obj);

  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
