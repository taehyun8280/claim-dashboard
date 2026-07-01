var SHEET_NAME = 'experimentDashboard';

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var callback = params.callback || '';

  try {
    if (params.action === 'ping') {
      return jsonp_(callback, {ok: true, message: 'pong'});
    }

    if (params.action === 'write') {
      var getPayload = JSON.parse(params.payload || '{}');
      writeState_(getPayload);
      return jsonp_(callback, {ok: true, updatedAt: getPayload.updatedAt || new Date().toISOString()});
    }

    var state = readState_();
    return jsonp_(callback, state);
  } catch (err) {
    return jsonp_(callback, {updatedAt: '', data: {experiments: []}, error: String(err)});
  }
}

function doPost(e) {
  var params = e && e.parameter ? e.parameter : {};

  try {
    var payloadText = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '{}');
    var payload = JSON.parse(payloadText);
    if (payload.action !== 'write') {
      return jsonp_('', {ok: false, error: 'unknown action'});
    }
    writeState_(payload);
    return jsonp_('', {ok: true, updatedAt: payload.updatedAt || new Date().toISOString()});
  } catch (err) {
    return jsonp_('', {ok: false, error: String(err)});
  }
}

function readState_() {
  var sheet = getSheet_();
  var raw = sheet.getRange(1, 1).getValue();
  if (!raw) return {updatedAt: '', experiments: []};

  try {
    var state = JSON.parse(raw);
    return {
      updatedAt: state.updatedAt || '',
      experiments: state.experiments || (state.data && state.data.experiments) || []
    };
  } catch (err) {
    return {updatedAt: '', experiments: []};
  }
}

function writeState_(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    var current = readState_();
    var currentTime = Date.parse(current.updatedAt || '') || 0;
    var nextTime = Date.parse(payload.updatedAt || '') || Date.now();
    if (currentTime > nextTime) return;

    var state = {
      updatedAt: payload.updatedAt || new Date().toISOString(),
      experiments: payload.experiments || (payload.data && payload.data.experiments) || []
    };
    var sheet = getSheet_();
    sheet.getRange(1, 1).setValue(JSON.stringify(state));
    sheet.getRange(1, 2).setValue(state.updatedAt);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from Google Sheets: Extensions > Apps Script.');
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function jsonp_(callback, obj) {
  var json = JSON.stringify(obj);
  var body = callback ? callback + '(' + json + ');' : json;
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
