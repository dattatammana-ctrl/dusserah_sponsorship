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
  External:    { hdr: ['ID', 'Name', 'Amount', 'Date', 'Added', 'Mode'], text: [1, 2, 4, 5, 6] },
  Archive:     { hdr: ['SID', 'Timestamp', 'Donor', 'Phone', 'Tower', 'Flat', 'Item Details', 'Qty', 'Amount', 'Mode', 'Payment Screenshot', 'Item Ref'], text: [1, 2, 3, 4, 5, 6, 7, 11] },
  Submissions: { hdr: ['SID', 'Timestamp', 'Donor', 'Phone', 'Tower', 'Flat', 'Item Details', 'Qty', 'Amount', 'Mode', 'Payment Screenshot', 'Item Ref (do not edit)'],  text: [1, 2, 3, 4, 5, 6, 11] }
};
const STATEMENT_TAB = 'Statement';   // bank statement: Type CR = money in, DR = money out
const r2 = n => Math.round(n * 100) / 100;
/* Maha Prasad donations now come from the separate donations form sheet (read live, column H = Amount, column B = Name) */
const FORM_SHEET_ID = '1HWj72ifsPxbtMbznQ8Lj6ZfSKZFZ3nX11en6CVoH5ZM';
const FORM_TAB = '';             // '' = first tab
const FORM_NAME_COL = 2, FORM_AMT_COL = 8, FORM_MODE_COL = 9, FORM_HEADER_ROWS = 1;
const MAHA_PRASAD_BUDGET = 0;    // Maha Prasad items are live in the Items tab and carry their own budget, so no extra budget is added here.
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
    if (a === 'mix') return mix_();
    const fn = { 'item': update_, 'item/add': add_, 'item/delete': del_, 'sub/delete': subDel_, 'ext/add': extAdd_, 'ext/update': extUpd_, 'ext/delete': extDel_ }[a];
    if (!fn) throw new Error('Not found');
    return locked_(() => fn(b));
  }
  throw new Error('Not found');
}

/** Run once: creates both sheets and loads the starting items. */
function setup() {
  const sh = sheet_('Items'), sb = sheet_('Submissions'); const ex = sheet_('External'); formTab_();
  if (!ex.getRange(1, 6).getValue()) ex.getRange(1, 6).setValue('Mode').setFontWeight('bold').setBackground('#fde7d0');
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
  return sh.getRange(2, 1, n, 12).getValues().map((r, i) => ({ r: r, row: i + 2 })).filter(x => x.r[0] !== '' || x.r[2] !== '' || x.r[8] !== '').map(x => {
    const r = x.r, a = Number(r[11]), b = Number(r[6]);
    return { sid: r[0] !== '' ? String(r[0]) : 'row' + x.row, ts: String(r[1]), donor: String(r[2]), phone: String(r[3]), tower: String(r[4]), flat: String(r[5]),
      rid: (r[11] !== '' && isFinite(a)) ? a : (r[6] !== '' && isFinite(b)) ? b : -1,   // -1 = not linked to an item (still counted in totals)
      qty: Number(r[7]) || 0, amt: Number(r[8]) || 0, mode: String(r[9] || 'Physical'), ref: String(r[10] || '') };
  });
}

function state_() {
  const c = CacheService.getScriptCache(), hit = c.get('state');
  if (hit) return JSON.parse(hit);
  const au = auction_(), subs = subs_(), items = items_(), m = money_(subs, ext_(), au);
  const s = { items: items, subs: subs.map(x => ({ sid: x.sid, rid: x.rid, qty: x.qty, amt: x.amt, pq: x.mode === 'Cash' ? 0 : x.qty })), other: m.other, hundi: m.hundi, hh: m.hh, hand: m.hand, spend: m.spend, phys: m.phys, bank: m.bank, net: m.net, sc: m.sc, sp: m.sp, ec: m.ec, ep: m.ep, mp: m.mp, mpBudget: MAHA_PRASAD_BUDGET, ext: m.ext, auction: au, donors: donors_(subs, items), cs: catSpends_() };   // public: totals only, no names/phones
  try { c.put('state', JSON.stringify(s), 30); } catch (e) {}
  return s;
}

function sponsored_(subs, rid) { return subs.reduce((a, s) => a + (s.rid === rid ? s.qty : 0), 0); }
// Only PHYSICAL bookings count toward the physical-quantity limit; cash has no limit and never uses it up.
function physical_(subs, rid) { return subs.reduce((a, s) => a + (s.rid === rid && s.mode !== 'Cash' ? s.qty : 0), 0); }

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
    const left = capOf_(it) - physical_(subs, id);
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
      const sp = physical_(subs_(), id), it = items_()[id], m = it ? capOf_({ req: 1, cat: it.cat, item: it.item }) : 1;
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
  const m = /^row(\d+)$/.exec(String(b.sid));                    // rows typed in by hand without an ID
  if (m && Number(m[1]) >= 2 && Number(m[1]) <= n + 1 && ids[Number(m[1]) - 2][0] === '') { sh.deleteRow(Number(m[1])); return { ok: true }; }
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(b.sid)) { sh.deleteRow(i + 2); break; }
  return { ok: true };
}

/* ---------- Maha Prasad donations form (separate sheet, read live) ---------- */
function formSum_(sh) {
  const n = sh.getLastRow() - FORM_HEADER_ROWS, o = { cash: 0, phys: 0, total: 0 };
  if (n < 1) return o;
  sh.getRange(FORM_HEADER_ROWS + 1, 1, n, FORM_MODE_COL).getValues().forEach(r => {
    if (String(r[FORM_NAME_COL - 1]).trim() === '') return;              // count only rows that have a donor name
    const v = r[FORM_AMT_COL - 1], x = typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
    if (String(v).trim() === '' || !isFinite(x)) return;
    /phys/i.test(String(r[FORM_MODE_COL - 1])) ? o.phys += x : o.cash += x;
  });
  o.cash = r2(o.cash); o.phys = r2(o.phys); o.total = r2(o.cash + o.phys);
  return o;
}
// Preferred: the live "Maha Prasad Form" tab in THIS sheet (IMPORTRANGE, created by setup). Fallback: read the form sheet directly.
function formSheet_() {
  try {
    const loc = ss_().getSheetByName('Maha Prasad Form');
    if (loc && loc.getLastRow() > 1 && !/^#|Loading/i.test(String(loc.getRange(1, 1).getValue()))) return loc;
  } catch (e) {}
  try {
    const ss = SpreadsheetApp.openById(FORM_SHEET_ID);
    return FORM_TAB ? ss.getSheetByName(FORM_TAB) : ss.getSheets()[0];
  } catch (e) { return null; }
}
function form_() { try { const sh = formSheet_(); return sh ? formSum_(sh) : null; } catch (e) { return null; } }
// one entry per named row of the donations form: { cat, mode, amt }
function formRows_() {
  const sh = formSheet_(), out = [];
  if (!sh) return out;
  const n = sh.getLastRow() - FORM_HEADER_ROWS;
  if (n < 1) return out;
  sh.getRange(FORM_HEADER_ROWS + 1, 1, n, FORM_MODE_COL).getValues().forEach(r => {
    if (String(r[FORM_NAME_COL - 1]).trim() === '') return;
    const v = r[FORM_AMT_COL - 1], x = typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
    if (String(v).trim() === '' || !isFinite(x)) return;
    out.push({ cat: String(r[4]).trim() || '(no category)', mode: /phys/i.test(String(r[FORM_MODE_COL - 1])) ? 'Physical' : 'Cash', amt: x });
  });
  return out;
}
/** Creates the live "Maha Prasad Form" tab (IMPORTRANGE of columns A:J). Safe to re-run. */
function formTab_() {
  const s = ss_();
  if (s.getSheetByName('Maha Prasad Form')) return;
  s.insertSheet('Maha Prasad Form').getRange(1, 1).setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/' + FORM_SHEET_ID + '","' + (FORM_TAB ? FORM_TAB + '!' : '') + 'A:J")');
}
/* ---------- bank statement (live hundi) ---------- */
function statement_() {
  const sh = ss_().getSheetByName(STATEMENT_TAB);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getDataRange().getValues(), h = v[0].map(x => String(x).trim().toLowerCase());
  const col = (re, d) => { const i = h.findIndex(x => re.test(x)); return i < 0 ? d : i; };
  const A = col(/amount/, 2), S = col(/status/, 4), T = col(/^type$/, 5);
  let cr = 0, dr = 0;
  v.slice(1).forEach(r => {
    const t = String(r[T]).trim().toUpperCase(), x = Number(String(r[A]).replace(/[^0-9.\-]/g, ''));
    if (!isFinite(x) || /err|fail|declin|revers|cancel|reject/i.test(String(r[S]))) return;   // failed transactions are ignored
    if (t === 'CR') cr += x; else if (t === 'DR') dr += x;
  });
  return { cr: r2(cr), dr: r2(dr) };
}
const split_ = rows => rows.reduce((o, x) => { x.mode === 'Cash' ? o.cash += x.amt : o.phys += x.amt; return o; }, { cash: 0, phys: 0 });

/* Hand Hold Cash tab: sums the "Amount" column (defaults to column C), skipping "Total" rows and rows whose "Deposited" column is Yes. */
function handHold_() {
  const sh = ss_().getSheetByName('Hand Hold Cash');
  if (!sh || sh.getLastRow() < 2) return 0;
  const v = sh.getDataRange().getValues(), h = v[0].map(x => String(x).trim().toLowerCase());
  let A = h.findIndex(x => /amount/.test(x) && !/deposit/.test(x)); if (A < 0) A = 2;
  const D = h.findIndex(x => /deposit/.test(x));      // "Deposited to bank account" column: Yes = already in the bank, so it is no longer hand-hold
  let t = 0;
  v.slice(1).forEach(r => {
    if (/^total/i.test(String(r[0]).trim())) return;
    if (D >= 0 && /^\s*(y|yes)\b/i.test(String(r[D]))) return;
    const x = num_(r[A]); if (isFinite(x)) t += x;
  });
  return r2(t);
}

/**
 * other  = Residents Donation form (Other Source tab) total
 * hundi  = bank credits - form cash (Other Source + Maha Prasad form + Submissions) - External cash - Auction credited to A/C
 * hh     = Hand Hold Cash tab; hand = hh + auction hand-hold
 * Spends = bank debits + all physical items.
 */
function money_(subs, ext, au) {
  const st = statement_(), f = form_(), res = otherForm_(), sb = split_(subs), ex = split_(ext);
  const ab = au ? au.bank : 0, ah = au ? au.hand : 0, hh = handHold_();
  const ok = !!(st && f);
  return { other: res,
           hundi: ok && res != null ? r2(st.cr - res - sb.cash - f.cash - ex.cash - ab) : null,
           hh: hh, hand: hh,   // the Hand Hold Cash sheet already includes the auction handhold rows, so they are not added again
           spend: ok ? r2(st.dr + sb.phys + f.phys + ex.phys) : null,
           phys: ok ? r2(sb.phys + f.phys + ex.phys) : null, bank: ok ? st.dr : null, net: st ? r2(st.cr - st.dr) : null,
           sc: f ? r2(sb.cash + f.cash) : null, sp: f ? r2(sb.phys + f.phys) : null, ec: r2(ex.cash), ep: r2(ex.phys),   // cash vs physical split of Sponsored / External
           mp: f ? f.total : null, ext: r2(ex.cash + ex.phys) };
}

/* Resident donation form = the "Other Source" tab (donations form): column L amount, only rows that have a name in column D */
const OTHER_SHEET_ID = '1N28jKtSsuePNc5PMhAwyuCFXutCJhtWwxljgqoaeLrA';
function sumOther_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  let t = 0;
  sh.getRange(2, 1, n, 12).getValues().forEach(r => {
    if (String(r[3]).trim() === '') return;
    const v = r[11], x = typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, ''));
    if (String(v).trim() !== '' && isFinite(x)) t += x;
  });
  return r2(t);
}
function otherSheet_() {
  try {
    const loc = ss_().getSheetByName('Other Source');
    if (loc && loc.getLastRow() > 1 && !/^#|Loading/i.test(String(loc.getRange(1, 1).getValue()))) return loc;
  } catch (e) {}
  try { return SpreadsheetApp.openById(OTHER_SHEET_ID).getSheets()[0]; } catch (e) { return null; }
}
function otherForm_() { try { const sh = otherSheet_(); return sh ? sumOther_(sh) : null; } catch (e) { return null; } }
/* ---------- auction (tab "Auction"): public gets totals + winner name/flat/day/alankarana/amount; mobile, payment status & comments stay private ---------- */
function auctionRows_() {
  const sh = ss_().getSheetByName('Auction');
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getDataRange().getValues(), h = v[0].map(x => String(x).trim().toLowerCase());
  const col = (re, d) => { const i = h.findIndex(x => re.test(x)); return i < 0 ? d : i; };
  const NO = col(/item|saree/, 0), DAY = col(/^day/, 1), AL = col(/alank|ammavaru/, 2), TW = col(/tower/, 3), FL = col(/^flat/, 4), AM = col(/amount/, 5), WN = col(/winner/, 6), MB = col(/mobile|phone/, 7), PY = col(/payment/, 8), CM = col(/comment|remark|note/, 9);
  const whole = x => (typeof x === 'number' && isFinite(x)) ? String(Math.round(x)) : String(x).trim();
  const out = [];
  v.slice(1).forEach(r => {
    const name = String(r[WN]).trim(), amt = Number(String(r[AM]).replace(/[^0-9.\-]/g, ''));
    if (name === '' && !(amt > 0)) return;
    const d = r[DAY];
    const pd = /^\s*paid\b/i.test(String(r[PY])), cm = String(r[CM] || '');
    out.push({ no: whole(r[NO]), day: typeof d === 'number' ? 'Day ' + Math.round(d) : String(d).trim(), dn: typeof d === 'number' ? d : 999,
               alank: String(r[AL]).trim(), name: name, tower: whole(r[TW]), flat: whole(r[FL]), mob: r[MB],
               amt: isFinite(amt) ? amt : 0, paid: pd,
               hand: pd && !/paid\s*to\s*committee/i.test(cm) && /hand\s*-?\s*hold/i.test(cm) });
  });
  return out;
}
function auction_() {
  const a = auctionRows_();
  if (!a) return null;
  a.sort((x, y) => x.dn - y.dn || Number(x.no) - Number(y.no));
  const sum = rows => r2(rows.reduce((t, x) => t + x.amt, 0));
  const paid = sum(a.filter(x => x.paid)), hand = sum(a.filter(x => x.paid && x.hand));
  return { total: sum(a), paid: paid, hand: hand, bank: r2(paid - hand), count: a.length,
           winners: a.map(x => ({ no: x.no, day: x.day, alank: x.alank, name: x.name, tower: x.tower, flat: x.flat, amt: x.amt })) };
}

/* ---------- donor honour board: unique donors across Submissions, Maha Prasad form and Residents form (External and Auction excluded; auction is shown separately) ---------- */
const BOARD_EXCLUDE = /commit/i;     // the committee's own entries are not donors
function digits_(v) { return typeof v === 'number' ? String(Math.round(v)) : String(v == null ? '' : v).replace(/\D/g, ''); }
function mobile_(v) { const d = digits_(v); return d.length >= 10 ? d.slice(-10) : ''; }
function flatNo_(v) { const t = typeof v === 'number' ? String(Math.round(v)) : String(v == null ? '' : v).trim().toUpperCase().replace(/\s+/g, ''); return /^\d+$/.test(t) ? String(parseInt(t, 10)) : t; }
function towerNo_(v) { const d = digits_(v); return d === '' ? '' : String(parseInt(d, 10)); }
function tf_(t, f) { const tt = towerNo_(t), ff = flatNo_(f); return tt !== '' && ff !== '' ? 'T' + tt + '-' + ff : ''; }
function tfText_(s) { const m = /(\d+)\s*[-\u2013]\s*([A-Za-z0-9]+)/.exec(String(s)); return m ? tf_(m[1], m[2]) : ''; }
function tidy_(n) { n = String(n).replace(/\s+/g, ' ').trim(); return (n === n.toUpperCase() && /[A-Z]/.test(n)) ? n.toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()) : n; }
function num_(v) { return typeof v === 'number' ? v : Number(String(v).replace(/[^0-9.\-]/g, '')); }
function donors_(subs, items) {
  const recs = [];
  subs.forEach(x => { const it = items[x.rid]; recs.push({ name: x.donor, tf: tf_(x.tower, x.flat), mob: mobile_(x.phone), cat: it ? String(it.cat).trim() : 'Others', amt: x.amt }); });
  try {
    const sh = formSheet_(), n = sh ? sh.getLastRow() - FORM_HEADER_ROWS : 0;
    if (n >= 1) sh.getRange(FORM_HEADER_ROWS + 1, 1, n, FORM_MODE_COL).getValues().forEach(r => {
      if (String(r[FORM_NAME_COL - 1]).trim() === '') return;
      const v = r[FORM_AMT_COL - 1], a = num_(v);
      if (String(v).trim() === '' || !isFinite(a)) return;
      recs.push({ name: r[FORM_NAME_COL - 1], tf: tfText_(r[2]), mob: mobile_(r[3]), cat: 'Maha Prasad', amt: a });
    });
  } catch (e) {}
  try {
    const sh = otherSheet_(), n = sh ? sh.getLastRow() - 1 : 0;
    if (n >= 1) sh.getRange(2, 1, n, 12).getValues().forEach(r => {
      if (String(r[3]).trim() === '') return;
      const v = r[11], a = num_(v);
      if (String(v).trim() === '' || !isFinite(a)) return;
      recs.push({ name: r[3], tf: tf_(r[1], r[2]), mob: mobile_(r[8]), cat: 'Residents Donation', amt: a });
    });
  } catch (e) {}

  const ok = recs.filter(r => String(r.name).trim() !== '' && isFinite(r.amt) && r.amt > 0 && !BOARD_EXCLUDE.test(String(r.name)));
  // A household = tower-flat.  A mobile number only stands in for a missing flat, and only if that mobile belongs to exactly one flat
  // (a number shared by several flats, e.g. someone submitting on others' behalf, is never used to merge families).
  const mobFlats = {};
  ok.forEach(r => { if (r.tf && r.mob) (mobFlats[r.mob] = mobFlats[r.mob] || {})[r.tf] = 1; });
  const groups = {};
  ok.forEach(r => {
    let k = r.tf;
    if (!k && r.mob) { const fl = Object.keys(mobFlats[r.mob] || {}); k = fl.length === 1 ? fl[0] : 'm:' + r.mob; }
    if (!k) k = 'n:' + String(r.name).trim().toLowerCase();
    const g = groups[k] || (groups[k] = { tf: /^T\d/.test(k) ? k : '', total: 0, cats: {}, best: null });
    g.total += r.amt; g.cats[r.cat] = (g.cats[r.cat] || 0) + r.amt;
    if (!g.best || r.amt > g.best.amt || (r.amt === g.best.amt && String(r.name).length > String(g.best.name).length)) g.best = r;
  });
  return Object.keys(groups).map(k => {
    const g = groups[k];
    return { name: tidy_(g.best.name), flat: g.tf ? g.tf.replace('-', ' - ') : '', total: r2(g.total),
             cats: Object.keys(g.cats).map(c => ({ c: c, a: r2(g.cats[c]) })).sort((x, y) => y.a - x.a) };
  }).sort((x, y) => y.total - x.total || x.name.localeCompare(y.name));
}

/* ---------- category spends (tab "Category Spends": A=Category, C=Physical, D=Cash, E=Total) ---------- */
function catSpends_() {
  try {
    const sh = ss_().getSheetByName('Category Spends');
    if (!sh || sh.getLastRow() < 2) return null;
    const n = x => { const a = num_(x); return isFinite(a) ? a : 0; };
    const m = {}, order = [];
    sh.getDataRange().getValues().slice(1).forEach(r => {
      const c = String(r[0]).trim();
      if (c === '' || /^total$/i.test(c)) return;
      const p = n(r[2]), k = n(r[3]), t = String(r[4]).trim() === '' ? p + k : n(r[4]);
      if (!m[c]) { m[c] = { cat: c, p: 0, c: 0, t: 0 }; order.push(c); }
      m[c].p += p; m[c].c += k; m[c].t += t;
    });
    return order.map(c => ({ cat: c, p: r2(m[c].p), c: r2(m[c].c), t: r2(m[c].t) }));
  } catch (e) { return null; }
}

/** Admin: where the bank credits come from, what is physical, and what is spent. */
function cashbook_(subs, ext) {
  const st = statement_();
  if (!st) return null;
  const f = form_(), res = otherForm_(), sb = split_(subs), ex = split_(ext), au = auction_(), ap = au ? au.bank : 0;   // only auction money credited to A/C
  const ok = res != null && !!f;                                   // the balancing figure needs every part
  const mpc = f ? f.cash : null, mpp = f ? f.phys : 0;
  const phys = r2(sb.phys + mpp + ex.phys);
  return {
    credit: { total: st.cr, resident: res, subs: r2(sb.cash), mp: mpc == null ? null : r2(mpc), ext: r2(ex.cash),
              auction: r2(ap),
              hundi: ok ? r2(st.cr - res - sb.cash - mpc - ex.cash - ap) : null },
    physical: { subs: r2(sb.phys), mp: f ? r2(mpp) : null, ext: r2(ex.phys), total: f ? phys : null },
    spend: { bank: st.dr, physical: f ? phys : null, total: f ? r2(st.dr + phys) : null },
    pendingBank: r2(st.cr - st.dr)
  };
}

/** Admin: cash vs physical contributions for every source and category. */
function mix_() {
  const S = subs_(), E = ext_(), items = items_(), acc = {}, order = ['Submissions', 'Maha Prasad Form', 'External'];
  const add = (src, cat, mode, amt) => {
    const k = src + '|' + cat, o = acc[k] || (acc[k] = { src: src, cat: cat, cash: 0, phys: 0 });
    mode === 'Cash' ? o.cash += amt : o.phys += amt;
  };
  S.forEach(x => { const it = items[x.rid]; add('Submissions', it ? it.cat : '(not linked)', x.mode === 'Cash' ? 'Cash' : 'Physical', x.amt); });
  formRows_().forEach(x => add('Maha Prasad Form', x.cat, x.mode, x.amt));
  E.forEach(x => add('External', 'External Contributions', x.mode === 'Physical' ? 'Physical' : 'Cash', x.amt));
  const rows = Object.keys(acc).map(k => acc[k]).map(o => ({ src: o.src, cat: o.cat, cash: r2(o.cash), phys: r2(o.phys) }))
    .sort((a, b) => order.indexOf(a.src) - order.indexOf(b.src) || a.cat.localeCompare(b.cat));
  return { rows: rows, book: cashbook_(S, E) };
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
  return sh.getRange(2, 1, n, 6).getValues().filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), name: String(r[1]), amt: Number(r[2]), mode: /phys/i.test(String(r[5])) ? 'Physical' : 'Cash',
    date: r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[3]) }));
}
function extMode_(b) {
  if (b.mode !== 'Cash' && b.mode !== 'Physical') throw new Error('Select payment mode (Cash or Physical item)');
  return b.mode;
}
function extAdd_(b) {
  const mode = extMode_(b);
  const name = String(b.name || '').trim().slice(0, 80), amt = Number(b.amt), date = String(b.date || '').trim();
  if (!name) throw new Error('Enter donor name');
  if (!isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Select a date');
  append_(sheet_('External'), [[Utilities.getUuid(), name, Math.round(amt * 100) / 100, date, new Date().toISOString(), mode]]);
  return { ok: true };
}
function extUpd_(b) {
  const mode = extMode_(b);
  const sh = sheet_('External'), n = sh.getLastRow() - 1;
  const name = String(b.name || '').trim().slice(0, 80), amt = Number(b.amt), date = String(b.date || '').trim();
  if (!name) throw new Error('Enter donor name');
  if (!isFinite(amt) || amt <= 0) throw new Error('Enter a valid amount');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Select a date');
  if (n < 1) throw new Error('Entry not found');
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === String(b.id)) {
    sh.getRange(i + 2, 2, 1, 3).setValues([[name, Math.round(amt * 100) / 100, date]]);
    sh.getRange(i + 2, 6).setValue(mode);
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
