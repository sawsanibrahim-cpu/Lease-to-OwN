/**
 * Kavak quote data — Google Sheet <-> Supabase sync
 * Attach to: "lease to own Warranty & Services", tab "Extract 1"
 *
 * Run setup() ONCE from the editor to install triggers.
 * Set SYNC_SECRET in File > Project Settings > Script Properties first.
 */

var CFG = {
  ENDPOINT:   'https://wpigboshlcehnlyipado.supabase.co/functions/v1/sheet-sync',
  SHEET_NAME: 'Extract 1',
  HEADER_ROW: 1,
  CHUNK:      300,
  KEY_HEADER: 'car_id'   // natural key. po_id is NOT unique (797 of 857).
};

function props_()  { return PropertiesService.getScriptProperties(); }
function secret_() { return props_().getProperty('SYNC_SECRET') || ''; }
function sheet_()  {
  var sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + CFG.SHEET_NAME);
  return sh;
}

/** Echo guard: our own write-back fires onEdit/onChange. Ignore those. */
function suppressed_() {
  var until = Number(props_().getProperty('SUPPRESS_UNTIL') || 0);
  return Date.now() < until;
}
function suppress_(ms) {
  props_().setProperty('SUPPRESS_UNTIL', String(Date.now() + (ms || 30000)));
}

function post_(body) {
  var res = UrlFetchApp.fetch(CFG.ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-sync-secret': secret_() },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('sheet-sync ' + code + ': ' + text);
  return JSON.parse(text);
}

function headers_() {
  var sh = sheet_();
  return sh.getRange(CFG.HEADER_ROW, 1, 1, sh.getLastColumn())
           .getValues()[0]
           .map(function (h) { return h === null || h === undefined ? '' : String(h); });
}

/* ------------------------------------------------------------------ */
/* Full reconcile — the safety net.                                    */
/* onEdit never fires for formula recalculation, IMPORTRANGE refresh,  */
/* or API writes, and this sheet is formula-driven. Do not remove.     */
/* ------------------------------------------------------------------ */
function fullSync() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var sh   = sheet_();
    var hdrs = headers_();
    var last = sh.getLastRow();
    var runId = Utilities.getUuid();
    var nRows = Math.max(0, last - CFG.HEADER_ROW);

    if (nRows === 0) { post_({ source: 'cron', runId: runId, headers: hdrs, full: true, finalize: true }); return; }

    var all = sh.getRange(CFG.HEADER_ROW + 1, 1, nRows, hdrs.length).getValues();
    var res = null;

    for (var i = 0; i < all.length; i += CFG.CHUNK) {
      var chunk = all.slice(i, i + CFG.CHUNK);
      var isLast = (i + CFG.CHUNK) >= all.length;
      res = post_({
        source:   'cron',
        runId:    runId,
        headers:  i === 0 ? hdrs : [],   // reconcile columns once per run
        rows:     chunk,
        full:     true,
        finalize: isLast
      });
    }
    if (res) applyPendingHeaders_(res.pendingHeaders);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Installable onEdit — queue dirty rows, flushed by flushQueue().     */
/* Debounced so a burst of edits is one request, not one per keystroke.*/
/* ------------------------------------------------------------------ */
function onEditInstallable(e) {
  if (suppressed_() || !e || !e.range) return;
  if (e.range.getSheet().getName() !== CFG.SHEET_NAME) return;

  var first = e.range.getRow();
  var last  = first + e.range.getNumRows() - 1;

  // A header edit needs the column reconcile path, not the row path.
  if (first <= CFG.HEADER_ROW) { syncHeadersOnly(); return; }

  var q = JSON.parse(props_().getProperty('DIRTY_ROWS') || '[]');
  for (var r = Math.max(first, CFG.HEADER_ROW + 1); r <= last; r++) {
    if (q.indexOf(r) === -1) q.push(r);
  }
  props_().setProperty('DIRTY_ROWS', JSON.stringify(q.slice(0, 5000)));
}

function flushQueue() {
  var q = JSON.parse(props_().getProperty('DIRTY_ROWS') || '[]');
  if (!q.length) return;
  props_().setProperty('DIRTY_ROWS', '[]');

  var sh   = sheet_();
  var hdrs = headers_();
  var rows = [];
  for (var i = 0; i < q.length; i++) {
    var r = q[i];
    if (r > sh.getLastRow()) continue;                       // row was deleted
    rows.push(sh.getRange(r, 1, 1, hdrs.length).getValues()[0]);
  }
  if (!rows.length) return;

  var res = post_({ source: 'onEdit', runId: Utilities.getUuid(), headers: hdrs, rows: rows });
  applyPendingHeaders_(res.pendingHeaders);
}

/* ------------------------------------------------------------------ */
/* Installable onChange — structural edits (insert/remove column).     */
/* ------------------------------------------------------------------ */
function onChangeInstallable(e) {
  if (suppressed_()) return;
  var t = e && e.changeType ? e.changeType : 'OTHER';
  if (t === 'INSERT_COLUMN' || t === 'REMOVE_COLUMN' || t === 'OTHER') {
    syncHeadersOnly();
  }
  if (t === 'INSERT_ROW' || t === 'REMOVE_ROW') {
    fullSync();
  }
}

/** Push headers only. Cheap; no row payload. */
function syncHeadersOnly() {
  var res = post_({ source: 'onChange', runId: Utilities.getUuid(), headers: headers_() });
  applyPendingHeaders_(res.pendingHeaders);
}

/* ------------------------------------------------------------------ */
/* DB -> Sheet: write admin header edits back, exact text.             */
/* ------------------------------------------------------------------ */
function applyPendingHeaders_(pending) {
  if (!pending || !pending.length) return;
  var sh = sheet_();
  var acked = [];

  suppress_(60000);   // our writes must not bounce back as sheet edits
  try {
    for (var i = 0; i < pending.length; i++) {
      var p = pending[i];
      if (!p.colIndex || p.header === null || p.header === undefined) continue;
      sh.getRange(CFG.HEADER_ROW, p.colIndex).setValue(p.header);
      acked.push(p.colIndex);
    }
    SpreadsheetApp.flush();
  } finally {
    if (acked.length) post_({ source: 'ack', runId: Utilities.getUuid(), ack: acked });
    suppress_(5000);
  }
}

/** Poll for admin edits even when nobody touches the sheet. */
function pollPendingHeaders() {
  var res = post_({ source: 'poll', runId: Utilities.getUuid() });
  applyPendingHeaders_(res.pendingHeaders);
}

/* ------------------------------------------------------------------ */
/* One-time setup                                                      */
/* ------------------------------------------------------------------ */
function setup() {
  if (!secret_()) {
    throw new Error('Set SYNC_SECRET in Project Settings > Script Properties first.');
  }
  var ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('flushQueue').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('pollPendingHeaders').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('fullSync').timeBased().everyMinutes(15).create();

  props_().setProperty('DIRTY_ROWS', '[]');
  fullSync();
}

/** Manual button for testing. */
function runFullSyncNow() { fullSync(); }

/**
 * Diagnose the shared secret without running a sync.
 * Compares what this script holds against what the server expects.
 */
function checkSecret() {
  var mine = props_().getProperty('SYNC_SECRET');

  if (mine === null) {
    Logger.log('SYNC_SECRET is NOT set in Script Properties.');
    Logger.log('Add it: File > Project Settings > Script Properties.');
    return;
  }

  var health = JSON.parse(
    UrlFetchApp.fetch(CFG.ENDPOINT, { method: 'get', muteHttpExceptions: true }).getContentText()
  );

  Logger.log('Apps Script secret : %s chars', mine.length);
  Logger.log('Supabase secret    : %s chars', health.secretLength);

  // HTTP headers are ASCII-only. Anything else gets percent-encoded in transit,
  // which is why a 66-char secret can arrive as 396 chars.
  var nonAscii = mine.replace(/[\x20-\x7E]/g, '');
  if (nonAscii.length) {
    Logger.log('>> %s of %s characters are NOT ASCII: %s',
               nonAscii.length, mine.length, nonAscii.slice(0, 20));
    Logger.log('>> Headers cannot carry these literally, so they are sent percent-encoded.');
    Logger.log('>> Replace the secret with A-Z a-z 0-9 - _ only, in BOTH places.');
  }

  if (mine !== mine.trim()) {
    Logger.log('>> Your value has leading/trailing whitespace. Retype it without.');
  }
  if (mine.length !== health.secretLength) {
    Logger.log('>> Lengths differ. The two values are not the same string.');
    Logger.log('>> Easiest fix: pick one value, retype it in BOTH places by hand.');
  } else {
    Logger.log('>> Lengths match. Running a live check...');
    var res = UrlFetchApp.fetch(CFG.ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-sync-secret': mine },
      payload: JSON.stringify({ source: 'check' }),
      muteHttpExceptions: true
    });
    Logger.log('Server said %s: %s', res.getResponseCode(), res.getContentText());
  }
}
