/**
 * Magdum Family Budget – Google Apps Script backend (installable app version)
 * Paste into Extensions > Apps Script of the "Magdum Family Budget" Google Sheet.
 * 1. Run setupKey once and copy the app key from the Execution log.
 * 2. Run setupTriggers once (daily missed-entry email).
 * 3. Deploy > New deployment > Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 *    The app only works with the secret key, which is never stored in the app's code.
 */

const USERS = {
  'rohitmagdum91@gmail.com': 'Rohit',
  'anjalimagdum23@gmail.com': 'Anjali'
};
const TZ = 'Asia/Kolkata';
const SHEET_ID = ''; // leave blank when this script is opened from the sheet itself
const FIRST = 5;     // first data row in Categories and Entries

/* ---------- helpers ---------- */
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
var CURRENT_USER = '';
function who_() {
  for (const e in USERS) { if (USERS[e] === CURRENT_USER) return { email: e, name: USERS[e] }; }
  throw new Error('Unknown user');
}
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function ym_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  return String(v || '').trim();
}
function ymd_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : ''; }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Tab missing: ' + name);
  return s;
}
function lock_(fn) {
  const l = LockService.getDocumentLock() || LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}

/* ---------- secure API for the installed app ---------- */
function setupKey() {
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  Logger.log('APP KEY (copy this into the app on both phones): ' + key);
  return key;
}
const API = {
  getData: getData, addEntries: addEntries, deleteEntry: deleteEntry, saveCategory: saveCategory,
  deleteCategory: deleteCategory, saveTargets: saveTargets, saveFundStart: saveFundStart,
  saveLoanOutstanding: saveLoanOutstanding, checkPin: checkPin, setPin: setPin, makePdf: makePdf, makeExcel: makeExcel
};
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
    if (!key || !req.key || req.key !== key) throw new Error('Wrong app key');
    const names = Object.keys(USERS).map(function (k) { return USERS[k]; });
    if (names.indexOf(req.user) < 0) throw new Error('Unknown user');
    CURRENT_USER = req.user;
    const fn = API[req.fn];
    if (!fn) throw new Error('Unknown action');
    out = { ok: true, result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  return ContentService.createTextOutput('Family Budget data service is running. Open the installed app to use it.');
}

/* ---------- read everything ---------- */
function getData() {
  const me = who_();
  const ss = ss_();

  const cs = sh_('Categories');
  const cRows = Math.max(cs.getLastRow() - FIRST + 1, 1);
  const cats = [];
  cs.getRange(FIRST, 1, cRows, 10).getValues().forEach(function (r, i) {
    if (!r[1]) return;
    cats.push({
      row: i + FIRST, section: String(r[0]), name: String(r[1]), kind: String(r[2]),
      monthly: num_(r[3]), yearly: num_(r[4]), due: String(r[5] || ''),
      start: ym_(r[6]), end: ym_(r[7]), group: String(r[8] || ''), notes: String(r[9] || '')
    });
  });

  const es = sh_('Entries');
  const eRows = Math.max(es.getLastRow() - FIRST + 1, 1);
  const entries = [];
  es.getRange(FIRST, 1, eRows, 10).getValues().forEach(function (r, i) {
    if (!(r[0] instanceof Date) || !r[3]) return;
    const d = ymd_(r[0]);
    entries.push({ row: i + FIRST, date: d, month: d.slice(0, 7), year: d.slice(0, 4),
      cat: String(r[3]), note: String(r[7] || ''), amt: num_(r[8]), by: String(r[9] || '') });
  });

  const targets = sh_('Savings Targets').getRange(6, 2, 5, 7).getValues().map(function (r) { return r.map(num_); });
  const pv = sh_('Planner').getRange(13, 1, 4, 4).getValues();
  const planner = pv.map(function (r) { return { name: String(r[0]), start: num_(r[1]), target: num_(r[3]) }; });
  const loans = sh_('Loans').getRange(5, 1, 6, 7).getValues().map(function (r, i) {
    return { row: i + 5, name: String(r[0]), outstanding: num_(r[1]), rate: num_(r[2]), emi: num_(r[3]),
      renewal: String(r[4] || ''), notes: String(r[6] || '') };
  });

  return {
    me: me, cats: cats, entries: entries, targets: targets, planner: planner, loans: loans,
    pinSet: !!PropertiesService.getScriptProperties().getProperty('PIN'),
    sheetUrl: ss.getUrl(), sheetId: ss.getId(),
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
  };
}

/* ---------- entries ---------- */
function writeEntry_(es, e, me) {
  const last = es.getLastRow();
  let row = -1;
  if (last >= FIRST) {
    const colA = es.getRange(FIRST, 1, last - FIRST + 1, 1).getValues();
    for (let i = 0; i < colA.length; i++) { if (colA[i][0] === '' || colA[i][0] === null) { row = i + FIRST; break; } }
  }
  if (row === -1) {
    row = Math.max(last + 1, FIRST);
    if (row > FIRST) es.getRange(row - 1, 1, 1, 10).copyTo(es.getRange(row, 1, 1, 10), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  }
  const p = String(e.date).split('-');
  es.getRange(row, 1).setValue(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
  es.getRange(row, 4).setValue(e.cat);
  es.getRange(row, 8).setValue(e.note || '');
  es.getRange(row, 9).setValue(Number(e.amt));
  es.getRange(row, 10).setValue(me.name);
  return row;
}
function addEntries(list) {
  const me = who_();
  if (!list || !list.length) throw new Error('Nothing to save');
  list.forEach(function (e) {
    if (!e.cat || !e.date || isNaN(Number(e.amt)) || Number(e.amt) === 0) throw new Error('Enter date, category and amount');
  });
  return lock_(function () {
    const es = sh_('Entries');
    list.forEach(function (e) { writeEntry_(es, e, me); });
    return getData();
  });
}
function deleteEntry(row) {
  who_();
  row = Number(row);
  if (row < FIRST) throw new Error('Bad row');
  return lock_(function () {
    const es = sh_('Entries');
    [1, 4, 8, 9, 10].forEach(function (c) { es.getRange(row, c).setValue(''); });
    return getData();
  });
}

/* ---------- categories ---------- */
function saveCategory(c) {
  who_();
  if (!c || !c.name || !c.section || !c.kind) throw new Error('Name, section and kind are needed');
  return lock_(function () {
    const cs = sh_('Categories');
    let row = Number(c.row || 0);
    if (!row) {
      const last = Math.max(cs.getLastRow(), FIRST);
      const colB = cs.getRange(FIRST, 2, last - FIRST + 2, 1).getValues();
      for (let i = 0; i < colB.length; i++) { if (!colB[i][0]) { row = i + FIRST; break; } }
    }
    const vals = [c.section, c.name, c.kind, null, Number(c.yearly) || '', c.due || '', c.start || '', c.end || '', c.group || '', c.notes || ''];
    vals.forEach(function (v, i) {
      if (i === 3) {
        const cell = cs.getRange(row, 4);
        if (!cell.getFormula()) cell.setValue(Number(c.monthly) || 0);
        return;
      }
      const cell = cs.getRange(row, i + 1);
      if (i === 6 || i === 7) cell.setNumberFormat('@');
      cell.setValue(v);
    });
    return getData();
  });
}
function deleteCategory(row) {
  who_();
  row = Number(row);
  if (row < FIRST) throw new Error('Bad row');
  return lock_(function () { sh_('Categories').getRange(row, 1, 1, 10).clearContent(); return getData(); });
}

/* ---------- targets, funds, loans ---------- */
function saveTargets(grid) {
  who_();
  if (!grid || grid.length !== 5) throw new Error('Bad targets');
  return lock_(function () {
    sh_('Savings Targets').getRange(6, 2, 5, 7).setValues(grid.map(function (r) { return r.map(function (v) { return Number(v) || 0; }); }));
    return getData();
  });
}
function saveFundStart(index, start, target) {
  who_();
  return lock_(function () {
    const pl = sh_('Planner');
    pl.getRange(13 + Number(index), 2).setValue(Number(start) || 0);
    if (target !== null && target !== undefined && target !== '') pl.getRange(13 + Number(index), 4).setValue(Number(target) || 0);
    return getData();
  });
}
function saveLoanOutstanding(row, value) {
  who_();
  return lock_(function () { sh_('Loans').getRange(Number(row), 2).setValue(Number(value) || 0); return getData(); });
}

/* ---------- PIN ---------- */
function hash_(p) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'fb-salt-' + String(p)));
}
function checkPin(p) {
  who_();
  const h = PropertiesService.getScriptProperties().getProperty('PIN');
  return !h || h === hash_(p);
}
function setPin(oldPin, newPin) {
  who_();
  if (!/^\d{4}$/.test(String(newPin))) throw new Error('PIN must be 4 digits');
  const props = PropertiesService.getScriptProperties();
  const h = props.getProperty('PIN');
  if (h && h !== hash_(oldPin)) throw new Error('Current PIN is wrong');
  props.setProperty('PIN', hash_(newPin));
  return true;
}

/* ---------- PDF report ---------- */
function makePdf(month) {
  const me = who_();
  const d = getData();
  const inMonth = d.entries.filter(function (e) { return e.month === month; });
  const fmt = function (n) { return '₹' + Math.round(n).toLocaleString('en-IN'); };
  const active = function (c) { return (!c.start || month >= c.start) && (!c.end || month <= c.end); };
  const kinds = ['Monthly', 'EMI', 'Set-aside', 'Saving'];
  let rows = '', tp = 0, ts = 0;
  d.cats.filter(function (c) { return kinds.indexOf(c.kind) >= 0; }).forEach(function (c) {
    const plan = active(c) ? c.monthly : 0;
    const spent = inMonth.filter(function (e) { return e.cat === c.name; }).reduce(function (s, e) { return s + e.amt; }, 0);
    if (!plan && !spent) return;
    tp += plan; ts += spent;
    rows += '<tr><td>' + c.section + '</td><td>' + c.name + '</td><td class="r">' + fmt(plan) + '</td><td class="r">' + fmt(spent) + '</td><td class="r">' + fmt(plan - spent) + '</td></tr>';
  });
  let log = '';
  inMonth.sort(function (a, b) { return a.date < b.date ? -1 : 1; }).forEach(function (e) {
    log += '<tr><td>' + e.date + '</td><td>' + e.cat + '</td><td>' + (e.note || '') + '</td><td class="r">' + fmt(e.amt) + '</td><td>' + e.by + '</td></tr>';
  });
  const html = '<html><head><style>body{font-family:Arial,sans-serif;font-size:11px;color:#15201C}h1{color:#0F766E;font-size:18px}' +
    'table{border-collapse:collapse;width:100%;margin-bottom:18px}th{background:#0F766E;color:#fff;text-align:left;padding:5px}' +
    'td{border-bottom:1px solid #ddd;padding:4px}.r{text-align:right}</style></head><body>' +
    '<h1>Magdum Family Budget – ' + month + '</h1><p>Generated by ' + me.name + ' on ' + d.today + '</p>' +
    '<h2>Planned vs spent</h2><table><tr><th>Section</th><th>Category</th><th class="r">Planned</th><th class="r">Spent</th><th class="r">Remaining</th></tr>' +
    rows + '<tr><td></td><td><b>Total</b></td><td class="r"><b>' + fmt(tp) + '</b></td><td class="r"><b>' + fmt(ts) + '</b></td><td class="r"><b>' + fmt(tp - ts) + '</b></td></tr></table>' +
    '<h2>All entries</h2><table><tr><th>Date</th><th>Category</th><th>Note</th><th class="r">Amount</th><th>By</th></tr>' + log + '</table></body></html>';
  const blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf');
  return { name: 'Budget ' + month + '.pdf', type: 'application/pdf', b64: Utilities.base64Encode(blob.getBytes()) };
}
function makeExcel() {
  who_();
  const id = ss_().getId();
  DriveApp.getFileById(id); // makes sure Drive permission is granted
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  return { name: 'Magdum Family Budget.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', b64: Utilities.base64Encode(blob.getBytes()) };
}

/* ---------- daily missed-entry email (run setupTriggers once) ---------- */
function dailyCheck() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const es = sh_('Entries');
  const n = Math.max(es.getLastRow() - FIRST + 1, 1);
  const any = es.getRange(FIRST, 1, n, 1).getValues().some(function (r) { return r[0] instanceof Date && ymd_(r[0]) === today; });
  if (any) return;
  MailApp.sendEmail({
    to: Object.keys(USERS).join(','),
    subject: 'Family Budget: no entries today (' + today + ')',
    body: 'Nobody added any expense today. Please open the Family Budget app and add today\'s entries.'
  });
}
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyCheck') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(21).nearMinute(30).inTimezone(TZ).create();
  return 'Daily check set for about 9:30 PM';
}
