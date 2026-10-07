/**
 * OnCloud33 Dusserah 2026 – Sponsorship Tracker backend (Google Apps Script + Google Sheets)
 *
 * SETUP: Sheet > Extensions > Apps Script > paste this as Code.gs, add an HTML file named Index (paste Index.html) > run setup() once
 *        > Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > share the /exec URL.
 * Sheets used: "Items" (sponsorship items) and "Submissions" (donor entries) - created automatically.
 * NOTE: Do not delete or sort rows in "Items" (row position = item id). To remove an item, put Y in the "Deleted" column.
 */
const ADMIN_PASSWORD = 'Dusserah2026';
const SHEET_ID = '1EGMumiBxnLya47f4eEn-RaWbxNpsW-Fs4RZHKKFRQoQ';   // your Google Sheet (works standalone or from Extensions > Apps Script)

const SHEETS = {
  Items:       { hdr: ['Category', 'Item', 'Date', 'Slot', 'Unit', 'Requirement', 'Est. amount', 'Deleted'], text: [1, 2, 3, 4, 5] },
  External:    { hdr: ['ID', 'Name', 'Amount', 'Date', 'Added'], text: [1, 2, 4, 5] },
  Submissions: { hdr: ['SID', 'Timestamp', 'Donor', 'Phone', 'Tower', 'Flat', 'Item Details', 'Qty', 'Amount', 'Mode', 'Payment Screenshot', 'Item Ref (do not edit)'],  text: [1, 2, 3, 4, 5, 6, 11] }
};
/* "Actual cash donations through other source": sum of column L of this sheet (no sign-in needed by donors; the script reads it as you) */
const OTHER_SHEET_ID = '1N28jKtSsuePNc5PMhAwyuCFXutCJhtWwxljgqoaeLrA';
const OTHER_TAB = '';        // tab name; '' = first tab
const OTHER_COL = 12;        // column L
const OTHER_HEADER_ROWS = 1; // rows to skip at the top
const FIELD_COL = { cat: 1, item: 2, date: 3, slot: 4, unit: 5, req: 6, est: 7 };
const EPS = 1e-9;
// Physical-booking limits: Rice 120%, other Maha Prasad 110%, Daily Prasad 120%, rest 100%. Cash has no limit.
function capOf_(it) {
  const c = String(it.cat).trim().toLowerCase(), n = String(it.item).trim().toLowerCase();
  return it.req * (c === 'maha prasad' ? (n === 'rice' ? 1.2 : 1.1) : c === 'daily prasad' ? 1.2 : 1);
}

/* ---------- entry points ---------- */
// GitHub Pages hosts the page; this script is only the data API.
function doGet() { return json_(state_()); }

function doPost(e) {
  try { return json_(handle_(JSON.parse(e.postData.contents))); }
  catch (err) { return json_({ error: err.message }); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(b) {
  const p = String(b.path);
  if (p === '/api/state') return state_();
  if (p === '/api/submit') return locked_(() => submit_(b));
  if (p.indexOf('/api/admin/') === 0) {
    if (b.pw !== ADMIN_PASSWORD) throw new Error('Wrong password');
    const a = p.slice(11);
    if (a === 'subs') return subs_();
    if (a === 'shot') return shotGet_(b);
    if (a === 'ext') return ext_();
    const fn = { 'item': update_, 'item/add': add_, 'item/delete': del_, 'sub/delete': subDel_, 'ext/add': extAdd_, 'ext/update': extUpd_, 'ext/delete': extDel_ }[a];
    if (!fn) throw new Error('Not found');
    return locked_(() => fn(b));
  }
  throw new Error('Not found');
}

/** Run once: creates both sheets and loads the starting items. */
function setup() {
  const sh = sheet_('Items'), sb = sheet_('Submissions'); sheet_('External'); otherTab_();
  if (!sb.getRange(1, 10).getValue()) sb.getRange(1, 10).setValue('Mode').setFontWeight('bold').setBackground('#fde7d0');
  sb.getRange(1, 11).setValue('Payment Screenshot').setFontWeight('bold').setBackground('#fde7d0');
  shotFolder_();   // creates the private Drive folder (also asks for Drive permission)
  migrateSubs_();  // old rows: replace numeric Item ID with readable item details
  if (sh.getLastRow() < 2) append_(sh, seed_());
  CacheService.getScriptCache().remove('state');
}

/* ---------- helpers ---------- */
function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  const s = ss_(); let sh = s.getSheetByName(name);
  if (!sh) {
    const c = SHEETS[name];
    sh = s.insertSheet(name);
    sh.getRange(1, 1, 1, c.hdr.length).setValues([c.hdr]).setFontWeight('bold').setBackground('#fde7d0');
    sh.setFrozenRows(1);
    c.text.forEach(col => sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('@')); // keep "11/10/26", phones etc. as text
  }
  return sh;
}

function append_(sh, rows) {
  const r = sh.getLastRow() + 1, w = rows[0].length;
  if (r + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length);
  SHEETS[sh.getName()].text.forEach(c => sh.getRange(r, c, rows.length, 1).setNumberFormat('@'));
  sh.getRange(r, 1, rows.length, w).setValues(rows);
}

function locked_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(25000);
  try { const r = fn(); SpreadsheetApp.flush(); CacheService.getScriptCache().remove('state'); return r; }
  finally { l.releaseLock(); }
}

function dateStr_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yy') : String(v).trim(); }
function dayOf_(d) { const m = /^(\d+)\/(\d+)\/(\d+)$/.exec(d); return m && +m[2] === 10 && +m[3] === 26 ? +m[1] : 0; }

function items_() {
  const sh = sheet_('Items'), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 8).getValues().map((r, i) => {
    const d = dateStr_(r[2]), req = Number(r[5]) || 0;
    return { id: i, cat: String(r[0]), item: String(r[1]), date: d, slot: String(r[3]), unit: String(r[4]),
             req: req, est: Number(r[6]) || 0, day: dayOf_(d),
             del: String(r[7]).toUpperCase() === 'Y' || !r[0] || !r[1] || !(req > 0) };
  });
}

function subs_() {
  const sh = sheet_('Submissions'), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 12).getValues().filter(r => r[0] !== '').map(r => ({
    sid: String(r[0]), ts: String(r[1]), donor: String(r[2]), phone: String(r[3]), tower: String(r[4]), flat: String(r[5]),
    rid: (r[11] !== '' && r[11] != null) ? Number(r[11]) : Number(r[6]), qty: Number(r[7]), amt: Number(r[8]), mode: String(r[9] || 'Physical'), ref: String(r[10] || '') }));
}

function state_() {
  const c = CacheService.getScriptCache(), hit = c.get('state');
  if (hit) return JSON.parse(hit);
  const s = { items: items_(), subs: subs_().map(x => ({ sid: x.sid, rid: x.rid, qty: x.qty, amt: x.amt })), other: other_(), ext: ext_().reduce((a, x) => a + x.amt, 0) };   // public gets only the external TOTAL, never names // public: no donor name/phone/flat
  try { c.put('state', JSON.stringify(s), 30); } catch (e) {}
  return s;
}

function sponsored_(subs, rid) { return subs.reduce((a, s) => a + (s.rid === rid ? s.qty : 0), 0); }

/* ---------- donor submit ---------- */
function submit_(b) {
  const donor = String(b.donor || '').trim().slice(0, 80), phone = String(b.phone || '').trim();
  const tower = String(b.tower || ''), flat = String(b.flat || '').trim().slice(0, 10);
  if (!donor) throw new Error('Enter donor name');
  if (!/^\d{10}$/.test(phone)) throw new Error('Enter a valid 10-digit phone number');
  if (['1', '2', '3', '4', '5'].indexOf(tower) < 0) throw new Error('Select tower (1-5)');
  if (!flat) throw new Error('Enter flat number');
  const mode = b.mode === 'Cash' ? 'Cash' : 'Physical';
  if (mode === 'Cash' && !b.shot) throw new Error('Please upload your payment screenshot');
  const need = {}, cash = {};
  (b.lines || []).slice(0, 200).forEach(l => {
    const q = Number(l.qty), id = Number(l.rid);
    if (mode === 'Cash') {                       // cash: donor chooses the rupee amount per line
      const a = Math.round(Number(l.amt));
      if (!isFinite(a) || a < 1 || a > 10000000) throw new Error('Enter a valid amount (minimum \u20b91) for each item');
      cash[id] = (cash[id] || 0) + a; need[id] = 0;
    } else if (q > 0 && isFinite(q)) need[id] = (need[id] || 0) + q;
  });
  const ids = Object.keys(need).map(Number);
  if (!ids.length) throw new Error('Cart is empty');

  const items = items_(), subs = subs_();
  ids.forEach(id => {
    const it = items[id];
    if (!it || it.del) throw new Error('An item in your cart was removed. Please review your cart.');
    if (mode === 'Cash') {                       // cash has no limit; store the share of the item that the amount covers
      if (!(it.est > 0)) throw new Error('"' + it.item + '" has no estimated cost, so cash cannot be accepted for it.');
      need[id] = Math.round(cash[id] * it.req / it.est * 1e6) / 1e6;
      return;
    }
    const left = capOf_(it) - sponsored_(subs, id);
    if (mode === 'Physical' && need[id] > left + EPS)
      throw new Error('Only ' + Math.max(0, Math.round(left * 1000) / 1000) + ' ' + it.unit + ' left for "' + it.item + '" (' + it.date + '). Someone just sponsored it - reduce the quantity or choose Cash.');
  });
  const ref = mode === 'Cash' ? saveShot_(b.shot, donor) : '';   // Drive link of the payment screenshot
  const ts = new Date().toISOString();
  append_(sheet_('Submissions'), ids.map(id => [Utilities.getUuid(), ts, donor, phone, tower, flat, detail_(items[id]),
    need[id], mode === 'Cash' ? cash[id] : Math.round(need[id] * items[id].est / items[id].req), mode, ref, id]));
  return { ok: true };
}

/* ---------- admin ---------- */
function checkField_(f, v, id) {
  if (f === 'req' || f === 'est') {
    v = Number(v);
    if (!isFinite(v) || v < 0 || (f === 'req' && v <= 0)) throw new Error('Invalid number');
    if (f === 'req') {
      const sp = sponsored_(subs_(), id), it = items_()[id], m = it ? capOf_({ req: 1, cat: it.cat, item: it.item }) : 1;
      if (v * m < sp - EPS) throw new Error("Quantity can't be lower than already sponsored (" + sp + ')');
    }
    return v;
  }
  if (!(f in FIELD_COL)) throw new Error('Bad field');
  v = String(v).trim().slice(0, 120);
  if (!v) throw new Error('Value required');
  if (f === 'date' && v !== 'All Days' && !/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(v)) throw new Error('Use dd/mm/yy or All Days');
  return v;
}

function update_(b) {
  const id = Number(b.id), sh = sheet_('Items');
  if (!(id >= 0) || id + 2 > sh.getLastRow()) throw new Error('Bad request');
  sh.getRange(id + 2, FIELD_COL[b.field]).setValue(checkField_(b.field, b.value, id));
  return { ok: true };
}

function add_(b) {
  const rows = b.rows || [];
  if (!rows.length || rows.length > 100) throw new Error('Bad request');
  append_(sheet_('Items'), rows.map(x => [checkField_('cat', x.cat), checkField_('item', x.item), checkField_('date', x.date),
    checkField_('slot', x.slot), checkField_('unit', x.unit), checkField_('req', x.req), checkField_('est', x.est), '']));
  return { ok: true };
}

function del_(b) {
  const id = Number(b.id), sh = sheet_('Items');
  if (!(id >= 0) || id + 2 > sh.getLastRow()) throw new Error('Bad request');
  sh.getRange(id + 2, 8).setValue('Y');            // soft delete keeps ids stable
  return { ok: true };
}

function subDel_(b) {
  const sh = sheet_('Submissions'), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(b.sid)) { sh.deleteRow(i + 2); break; }
  return { ok: true };
}

/* ---------- other-source cash (column L) ----------
 * Preferred: tab "Other Source" in THIS sheet, filled by IMPORTRANGE (created by setup()). One-time: open that tab and click "Allow access".
 * Fallback: read the other spreadsheet directly (works if the script owner can open it). */
function sumCol_(sh, headerRows) {
  const n = sh.getLastRow() - headerRows;
  if (n < 1) return 0;
  let t = 0;
  sh.getRange(headerRows + 1, 1, n, OTHER_COL).getValues().forEach(r => {
    if (r.some(c => /total/i.test(String(c)))) return;          // skip any "Total" row so nothing is counted twice
    const v = r[OTHER_COL - 1];
    if (String(v).trim() === '') return;
    const x = typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
    if (isFinite(x)) t += x;
  });
  return Math.round(t * 100) / 100;
}
function other_() {
  try {
    const loc = ss_().getSheetByName('Other Source');
    if (loc && loc.getLastRow() > 1 && !/^#|Loading/i.test(String(loc.getRange(1, 1).getValue()))) return sumCol_(loc, OTHER_HEADER_ROWS);
  } catch (e) {}
  try {
    const ss = SpreadsheetApp.openById(OTHER_SHEET_ID);
    return sumCol_(OTHER_TAB ? ss.getSheetByName(OTHER_TAB) : ss.getSheets()[0], OTHER_HEADER_ROWS);
  } catch (e) { return null; }    // null = could not read
}
/** Creates the "Other Source" tab with an IMPORTRANGE of columns A:L (called from setup). */
function otherTab_() {
  const s = ss_();
  if (s.getSheetByName('Other Source')) return;
  const sh = s.insertSheet('Other Source');
  sh.getRange(1, 1).setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + OTHER_SHEET_ID + '","' + (OTHER_TAB ? OTHER_TAB + '!' : '') + 'A:L")');
}

/* ---------- readable item details in Submissions ---------- */
function detail_(it) { return [it.cat, it.item, it.date, it.slot].join(' | '); }
// Column G = readable details; column L (hidden) = numeric item reference used by the app. Safe to run repeatedly.
function migrateSubs_() {
  const sh = sheet_('Submissions'), H = SHEETS.Submissions.hdr;
  const hr = sh.getRange(1, 1, 1, H.length); hr.setValues([H]); hr.setFontWeight('bold').setBackground('#fde7d0');
  const n = sh.getLastRow() - 1;
  if (n >= 1) {
    const items = items_(), g = sh.getRange(2, 7, n, 1).getValues(), l = sh.getRange(2, 12, n, 1).getValues();
    let ch = false;
    for (let i = 0; i < n; i++) {
      const empty = l[i][0] === '' || l[i][0] == null, v = g[i][0];
      if (empty && v !== '' && v != null && isFinite(Number(v))) {
        const id = Number(v), it = items[id];
        l[i][0] = id; g[i][0] = it ? detail_(it) : 'Item #' + id; ch = true;
      }
    }
    if (ch) { sh.getRange(2, 7, n, 1).setValues(g); sh.getRange(2, 12, n, 1).setValues(l); }
  }
  sh.hideColumns(12);
}

/* ---------- payment screenshots (private Drive folder; viewable only through the admin password) ---------- */
function shotFolder_() {
  const name = 'Dusserah 2026 Payment Screenshots', it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function saveShot_(d, donor) {
  const m = /^data:image\/(jpeg|png|webp);base64,/.exec(String(d).slice(0, 40));
  if (!m) throw new Error('Invalid screenshot. Please upload an image.');
  const body = String(d).slice(m[0].length);
  if (body.length < 200 || body.length > 4000000) throw new Error('Screenshot is empty or too large');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const nm = 'pay_' + stamp + '_' + String(donor).replace(/[^A-Za-z0-9]/g, '').slice(0, 20) + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1]);
  return shotFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(body), 'image/' + m[1], nm)).getUrl();
}
function shotGet_(b) {      // only files linked from a Submissions row can be read
  const sh = sheet_('Submissions'), n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('Not found');
  const rows = sh.getRange(2, 1, n, 11).getValues();
  for (let i = 0; i < n; i++) if (String(rows[i][0]) === String(b.sid)) {
    const m = /\/d\/([^\/?]+)/.exec(String(rows[i][10]));
    if (!m) throw new Error('No screenshot for this entry');
    const blob = DriveApp.getFileById(m[1]).getBlob();
    return { data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  }
  throw new Error('Not found');
}

/* ---------- external sponsorship (admin only) ---------- */
function ext_() {
  const sh = sheet_('External'), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 5).getValues().filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), amt: Number(r[2]),
    date: r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[3]) }));
}
function extAdd_(b) {
  const name = String(b.name || '').trim().slice(0, 80), amt = Number(b.amt), date = String(b.date || '').trim();
  if (!name) throw new Error('Enter donor name');
  if (!isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Select a date');
  append_(sheet_('External'), [[Utilities.getUuid(), name, Math.round(amt * 100) / 100, date, new Date().toISOString()]]);
  return { ok: true };
}
function extUpd_(b) {
  const sh = sheet_('External'), n = sh.getLastRow() - 1;
  const name = String(b.name || '').trim().slice(0, 80), amt = Number(b.amt), date = String(b.date || '').trim();
  if (!name) throw new Error('Enter donor name');
  if (!isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Select a date');
  if (n < 1) throw new Error('Entry not found');
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(b.id)) {
    sh.getRange(i + 2, 2, 1, 3).setValues([[name, Math.round(amt * 100) / 100, date]]);
    return { ok: true };
  }
  throw new Error('Entry not found');
}
function extDel_(b) {
  const sh = sheet_('External'), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(b.id)) { sh.deleteRow(i + 2); break; }
  return { ok: true };
}

/* ---------- starting data ---------- */
function seed_() {
  const rows = [], ds = d => d + '/10/26';
  const add = (cat, item, date, slot, unit, req, est) => rows.push([cat, item, date, slot, unit, req, est, '']);
  for (let d = 11; d <= 20; d++) add('Puja', 'Puja Items', ds(d), 'Day ' + (d - 10), 'Qty', 1, 500);
  add('Puja', 'Flowers', ds(10), 'Day 0', 'kg', 5, 1500);
  for (let d = 11; d <= 20; d++) add('Puja', 'Flowers', ds(d), 'Day ' + (d - 10), 'kg', d === 18 ? 5 : 3, d === 18 ? 1500 : 1000);
  for (let d = 11; d <= 20; d++) add('Puja', 'Ammavaru Saree', ds(d), 'Day ' + (d - 10), 'Qty', 1, 500);
  for (let d = 11; d <= 20; d++) if (d !== 17) add('Daily Prasad', 'Prasad item to be shared', ds(d), 'Morning', 'Kg', 5, 1500);
  for (let d = 11; d <= 20; d++) add('Daily Prasad', 'Prasad item to be shared', ds(d), 'Evening', 'Kg', 10, 3000);
  add('Daily Prasad', 'Cylinders', 'All Days', 'Morning', 'Qty', 1, 1000);
  add('Daily Prasad', 'Cylinders', 'All Days', 'Evening', 'Qty', 3, 3000);
  [['Rice', 'Kg', 100, 8000], ['Cooking Oil', 'Kg', 45, 8400], ['Curd', 'Kg', 60, 4800],
   ['Disposables - Paper plates', 'Qty', 1000, 2200], ['Disposables - Water glasses', 'Kg', 2000, 4400],
   ['Groceries', 'Kg', 100, 8000], ['Vegetables', 'Kg', 100, 5000], ['Cylinders', 'Qty', 4, 4000]]
    .forEach(x => add('Maha Prasad', x[0], ds(17), 'Morning', x[1], x[2], x[3]));
  return rows;
}
