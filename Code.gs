/**
 * AgriLedger → Google Sheet sync
 * Paste this whole file into Extensions › Apps Script of the Google Sheet,
 * then follow SETUP-GUIDE.html. The app sends records; this script writes
 * them into one tab per section (Farmers, Purchases, Payments …).
 * The app is the place to make corrections; this Sheet is a live copy.
 */
var ALLOWED_EMAIL = 'anujagrawalsha7@gmail.com';   // only this Google account may own the Sheet
var INR = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('AgriLedger')
    .addItem('1. Set up (run once)', 'setup')
    .addItem('2. Show sync key', 'showKey')
    .addItem('Make a new sync key', 'resetKey')
    .addToUi();
}

function checkOwner_() {
  var me = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (me && me !== ALLOWED_EMAIL.toLowerCase()) {
    throw new Error('Please open this Sheet signed in as ' + ALLOWED_EMAIL + ' (now signed in as ' + me + ').');
  }
  return me;
}

function setup() {
  checkOwner_();
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SYNC_KEY')) props.setProperty('SYNC_KEY', newKey_());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var info = ss.getSheetByName('Read me') || ss.insertSheet('Read me', 0);
  info.clear();
  info.getRange(1, 1, 6, 1).setValues([
    ['AgriLedger — live copy of the app'],
    ['Each tab is filled by the AgriLedger app on the phone. Make corrections in the app, not here.'],
    ['Money is in rupees. Quantities are in quintals. Dates are DD-MM-YYYY.'],
    ['Government procurement values are agency records, not your turnover.'],
    ['Menu AgriLedger › Show sync key gives the key to paste into the app (Settings › Google Sheet).'],
    ['Owner: ' + ALLOWED_EMAIL]
  ]);
  info.getRange(1, 1).setFontWeight('bold').setFontSize(14).setFontColor('#1d5a26');
  info.setColumnWidth(1, 720);
  showKey();
}

function showKey() {
  checkOwner_();
  var key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  var msg = key ? 'Your sync key (paste it into the app: Settings › Google Sheet):\n\n' + key
                : 'Run "AgriLedger › 1. Set up" first.';
  try { SpreadsheetApp.getUi().alert('AgriLedger', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(msg); }
}

function resetKey() {
  checkOwner_();
  PropertiesService.getScriptProperties().setProperty('SYNC_KEY', newKey_());
  showKey();
}

function newKey_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">AgriLedger sync is running. Paste this page\'s address into the app.</p>');
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad request' }); }
  var key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  if (!key || body.key !== key) return out_({ ok: false, error: 'wrong key' });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (body.ping) return out_({ ok: true, sheet: ss.getName(), owner: Session.getEffectiveUser().getEmail() });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(28000)) return out_({ ok: false, error: 'The Sheet is busy. Will try again.' });
  try {
    var n = 0;
    for (var name in body.tabs) n += upsert_(ss, name, body.tabs[name]);
    return out_({ ok: true, rows: n });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* Write rows into a tab: update the row with the same ID, or add a new row. */
function upsert_(ss, name, t) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var heads = t.cols.map(function (c) { return c.h; }), w = heads.length;
  var last = sh.getLastRow();
  var cur = last ? sh.getRange(1, 1, 1, Math.max(w, sh.getLastColumn())).getValues()[0] : [];
  if (heads.join('|') !== cur.slice(0, w).join('|')) {
    sh.getRange(1, 1, 1, w).setValues([heads]).setFontWeight('bold').setBackground('#1d5a26').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    t.cols.forEach(function (c, i) {
      var col = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1);
      if (c.t === 'd') col.setNumberFormat('dd-mm-yyyy');
      else if (c.t === 'm') col.setNumberFormat(INR);
      else if (c.t === 'q') col.setNumberFormat('0.00##');
      else if (c.t === 't') col.setNumberFormat('@');
    });
    if (last === 0) last = 1;
  }
  // newest version of each ID wins
  var byId = {}, order = [];
  t.rows.forEach(function (r) {
    var id = String(r[0]);
    if (!(id in byId)) order.push(id);
    byId[id] = r.map(function (v, i) { return t.cols[i].t === 'd' ? toDate_(v) : v; });
  });
  last = Math.max(sh.getLastRow(), 1);
  var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var pos = {};
  for (var i = 0; i < ids.length; i++) pos[String(ids[i][0])] = i + 2;
  var updates = [], adds = [];
  order.forEach(function (id) { if (pos[id]) updates.push([pos[id], byId[id]]); else adds.push(byId[id]); });
  if (updates.length > 25) {
    var range = sh.getRange(2, 1, last - 1, w), data = range.getValues();
    updates.forEach(function (u) { data[u[0] - 2] = u[1]; });
    range.setValues(data);
  } else {
    updates.forEach(function (u) { sh.getRange(u[0], 1, 1, w).setValues([u[1]]); });
  }
  if (adds.length) sh.getRange(last + 1, 1, adds.length, w).setValues(adds);
  return order.length;
}

function toDate_(v) {
  if (!v) return '';
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v));
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : v;
}
