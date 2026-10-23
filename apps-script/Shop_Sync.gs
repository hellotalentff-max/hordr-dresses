// ============================================================
//  RENTAL HOUSE — sheet sync with the Click Lounge Studio dress system
//
//  One copy goes in EACH rental house's OWN Google Sheet (Extensions → Apps Script):
//  House of Roan's sheet gets one, Serra's sheet gets one. Each copy signs in with
//  that house's partner PIN, so it only ever sees that house's own dresses.
//
//  It keeps one tab ("Tracker") in two-way sync with the dress system:
//
//    • Rentals booked anywhere (the dress website, CLS bookings) appear here.
//      CLS bookings show as "CLS client" with dates only, read-only.
//    • A new row typed here (Name of Renter + Dress Name + Event Date) is sent to the
//      dress system and blocks the dress for everyone. It gets a Rental ID.
//    • Changing Booking Status, dates, payment Status, Price, Contact or Notes on your
//      own rows is sent too (e.g. set Booking Status to "Returned").
//
//  One-time setup: fill in the three values under SETTINGS, save, then choose
//  setupSync in the function dropdown and click Run.
// ============================================================

// ── SETTINGS ────────────────────────────────────────────────
var SHOP_NAME_HERE     = 'House of Roan';   // exactly as in the Brand column at Click Lounge Studio: 'House of Roan' or 'Serra'
var DRESS_API_URL_HERE = 'https://script.google.com/macros/s/AKfycbzxb4jzWZ8IWijNr-q63t-n4pZUu7QOlsPDFCHHk_CXjIjD8EIu2eGWfoGPQIo4z6eT/exec';   // the dress web app URL from Click Lounge Studio (ends in /exec)
var PARTNER_PIN_HERE   = '';                // this house's partner PIN

var SYNC = {
  TAB: 'Tracker',                  // the tab to keep in sync (created if missing)
  PRICE_TAB: 'Price List',         // your price per dress; fills the Tracker's Price column
  TZ: 'Asia/Manila',
  DEFAULT_CO: '',                  // used when C/O is left blank; blank = your first channel (e.g. "HORDR Direct")
  CLS_PARTNERSHIP_RATE: 500,       // what CLS pays you per dress, shown on CLS rows
  LOOK_BACK_DAYS: 120,             // how far back / ahead rentals are pulled in
  LOOK_AHEAD_DAYS: 400
};

// Column headers, matched by name (order doesn't matter, other columns of yours are left alone)
var COL = {
  renter: 'Name of Renter', dress: 'Dress Name', event: 'Event Date', pickup: 'Date of Pick-up',
  ret: 'Date of Return', status: 'Booking Status', co: 'C/O', rate: 'Partnership Rate', pay: 'Status',
  price: 'Price', contact: 'Contact No.', notes: 'Notes', id: 'Rental ID', note: 'Sync', snap: '_sync'
};
var COL_ORDER = ['renter', 'dress', 'event', 'pickup', 'ret', 'status', 'co', 'rate', 'pay',
                 'price', 'contact', 'notes', 'id', 'note', 'snap'];
var FIELDS = ['renter', 'dress', 'event', 'pickup', 'ret', 'status', 'co', 'rate', 'pay', 'price', 'contact', 'notes'];
var DATE_FIELDS = { event: 1, pickup: 1, ret: 1 };

// ── MENU ────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('👗 Dress sync')
    .addItem('Sync now', 'syncNow')
    .addItem('Set up (first time)', 'setupSync')
    .addToUi();
}

// ── SETUP ───────────────────────────────────────────────────
function setupSync() {
  var cfg = sync_settings_();
  if (!cfg.shop || !cfg.url || !cfg.pin) throw new Error('Fill in SHOP_NAME_HERE, DRESS_API_URL_HERE and PARTNER_PIN_HERE at the top of this file first.');
  var who = api_('post', { action: 'whoami' });
  if (who.role !== 'partner') throw new Error('That PIN signs in as "' + who.role + '". Use the ' + cfg.shop + ' partner PIN.');
  if (who.brand !== cfg.shop) throw new Error('That PIN belongs to "' + who.brand + '", not "' + cfg.shop + '". Check SHOP_NAME_HERE and PARTNER_PIN_HERE.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(SYNC.TZ);
  var sh = ss.getSheetByName(SYNC.TAB) || ss.insertSheet(SYNC.TAB);
  var head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  COL_ORDER.forEach(function (k) {
    if (head.indexOf(COL[k]) >= 0) return;
    head.push(COL[k]);
    sh.getRange(1, head.length).setValue(COL[k]);
  });
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  sh.hideColumns(head.indexOf(COL.snap) + 1);
  sync_formatColumns_(sh, head);
  sync_dropdowns_(sh, head);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['syncNow', 'onEditSync'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditSync').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('syncNow').timeBased().everyMinutes(15).create();

  var n = syncNow();
  console.log('✅ Connected as ' + cfg.shop + '. "' + SYNC.TAB + '" now has ' + n + ' rental(s) and syncs on every edit and every 15 minutes.');
}

// ── TRIGGERS ────────────────────────────────────────────────
/** Runs every 15 minutes and from the menu. Also refreshes the dress dropdown. */
function syncNow() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) return 0;
  try {
    var sh = sync_sheet_();
    try { sync_dropdowns_(sh, sync_head_(sh)); } catch (e) { console.warn('Dropdown refresh skipped: ' + e.message); }
    return sync_run_(sh);
  } finally { lock.releaseLock(); }
}

/** Installed by setupSync: sends your edits right away. */
function onEditSync(e) {
  var sh = e && e.range ? e.range.getSheet() : null;
  if (!sh || sh.getName() !== SYNC.TAB || e.range.getRow() < 2) return;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(60000)) return;           // the 15-minute run will pick it up
  try { sync_run_(sh); } finally { lock.releaseLock(); }
}

// ── THE SYNC ────────────────────────────────────────────────
function sync_run_(sh) {
  var head = sync_head_(sh), c = sync_cols_(head);
  var before = sh.getDataRange().getValues();
  var today = sync_iso_(new Date());
  var notes = {};          // by rental ID or by 'row:N'
  var newIds = {};         // rental ID → row index in `before`
  var prices = sync_prices_();
  var defaultCo = sync_defaultCo_();

  // 0. Blank Price on your own rows → fill it from the Price List tab
  for (var p = 1; p < before.length; p++) {
    var dn = String(before[p][c.dress] || '').trim().toLowerCase(), pr = before[p][c.price];
    if (!dn || (pr !== '' && pr !== null) || String(before[p][c.co] || '') === 'CLS booking' || prices[dn] === undefined) continue;
    before[p][c.price] = prices[dn];
    sh.getRange(p + 1, c.price + 1).setValue(prices[dn]);
  }

  // 1. Push: new rows and edits to your own rows
  for (var i = 1; i < before.length; i++) {
    var row = sync_read_(before[i], c), id = String(before[i][c.id] || '').trim();
    var snap = sync_parse_(before[i][c.snap]);
    if (!id) {
      if (!row.renter && !row.dress && !row.event) continue;
      if (!row.renter || !row.dress || !row.event) { notes['row:' + i] = 'Fill in Name of Renter, Dress Name and Event Date to send'; continue; }
      var co = row.co || defaultCo;
      if (/^CLS/.test(co)) { notes['row:' + i] = '⚠ CLS packages are booked by Click Lounge Studio'; continue; }
      if (row.event < today) { notes['row:' + i] = 'Past date — kept here only'; continue; }
      if (row.status === 'Cancelled' || row.status === 'Returned') { notes['row:' + i] = 'Not sent (' + row.status + ')'; continue; }
      try {
        var res = api_('post', sync_clean_({ action: 'reserve', dress: row.dress, eventDate: row.event, co: co,
          renter: row.renter, contact: row.contact, notes: row.notes,
          pickup: row.pickup || undefined, returnDate: row.ret || undefined,
          clientPrice: row.price === '' ? undefined : row.price }));
      } catch (e) { notes['row:' + i] = '⚠ ' + e.message; continue; }
      sh.getRange(i + 1, c.id + 1).setValue(res.rentalId);            // save the ID straight away
      newIds[res.rentalId] = i;
      try {
        var st = row.status || 'Confirmed';                            // blank = a real booking, not a 48-hour hold
        if (st !== 'Reserved') api_('post', { action: 'setStatus', rentalId: res.rentalId, status: st });
        if (row.pay && row.pay !== 'Unpaid') api_('post', { action: 'updateRental', rentalId: res.rentalId, payStatus: row.pay });
      } catch (e) { notes[res.rentalId] = '⚠ Added, but: ' + e.message; }
      continue;
    }
    if (!snap) continue;                                             // never synced yet: just pull
    if (!snap.mine) { if (FIELDS.some(function (f) { return row[f] !== snap[f]; })) notes[id] = 'CLS booking — read only (change undone)'; continue; }
    try {
      if (row.dress !== snap.dress) notes[id] = '⚠ To change the dress, set this row to Cancelled and add a new row';
      var up = {}, map = { renter: 'renter', contact: 'contact', notes: 'notes', event: 'eventDate',
                           pickup: 'pickup', ret: 'returnDate', co: 'co', price: 'clientPrice', pay: 'payStatus' };
      Object.keys(map).forEach(function (f) { if (row[f] !== snap[f]) up[map[f]] = row[f]; });
      if (up.co === '') up.co = defaultCo;
      if (up.co && /^CLS/.test(up.co)) throw new Error('CLS packages are booked by Click Lounge Studio');
      if (Object.keys(up).length) { up.action = 'updateRental'; up.rentalId = id; api_('post', up); }
      if (row.status !== snap.status && row.status) api_('post', { action: 'setStatus', rentalId: id, status: row.status });
    } catch (e) { notes[id] = '⚠ ' + e.message + ' (change undone)'; }
  }

  // 2. Pull everything the dress system has for this house
  var from = sync_iso_(new Date(Date.now() - SYNC.LOOK_BACK_DAYS * 864e5));
  var to = sync_iso_(new Date(Date.now() + SYNC.LOOK_AHEAD_DAYS * 864e5));
  var list = api_('post', { action: 'listRentals', from: from, to: to });

  // 3. Write back, without overwriting anything someone typed while this ran
  var now = sh.getDataRange().getValues();
  var width = Math.max(head.length, now[0] ? now[0].length : 0);
  var rowOf = {}, beforeRowOf = {};
  now.forEach(function (r, k) { var x = String(r[c.id] || '').trim(); if (k && x) rowOf[x] = k; });
  before.forEach(function (r, k) { var x = String(r[c.id] || '').trim(); if (k && x) beforeRowOf[x] = k; });
  Object.keys(newIds).forEach(function (x) { beforeRowOf[x] = newIds[x]; });
  var stamp = Utilities.formatDate(new Date(), SYNC.TZ, 'MMM d, h:mm a');

  list.forEach(function (r) {
    var want = sync_fromServer_(r), k = rowOf[r.rentalId];
    if (k === undefined) { var blank = []; for (var j = 0; j < width; j++) blank.push(''); now.push(blank); k = now.length - 1; rowOf[r.rentalId] = k; }
    var b = beforeRowOf[r.rentalId] !== undefined ? sync_read_(before[beforeRowOf[r.rentalId]], c) : null;
    var cur = sync_read_(now[k], c);
    FIELDS.forEach(function (f) {
      if (b && cur[f] !== b[f]) return;                  // typed during this sync → keep it, it goes next time
      now[k][c[f]] = sync_cell_(f, want[f]);
    });
    now[k][c.id] = r.rentalId;
    now[k][c.snap] = JSON.stringify(want);
    now[k][c.note] = notes[r.rentalId] || (r.mine ? '✓ ' + stamp : 'CLS booking — read only');
  });
  Object.keys(notes).forEach(function (key) {
    if (key.indexOf('row:') === 0) { var n = +key.slice(4); if (now[n] && !String(now[n][c.id] || '')) now[n][c.note] = notes[key]; }
  });
  // rows that were cleared since the last run lose their old note
  for (var q = 1; q < now.length; q++) {
    if (!notes['row:' + q] && !String(now[q][c.id] || '') && !now[q][c.renter] && !now[q][c.dress]) now[q][c.note] = '';
  }

  if (now.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), now.length - sh.getMaxRows());
  // write only the columns this script manages, so your own formulas elsewhere are untouched
  COL_ORDER.forEach(function (key) {
    if (key === 'note' || key === 'snap' || key === 'id' || FIELDS.indexOf(key) >= 0) {
      var colVals = now.slice(1).map(function (r) { return [r[c[key]] === undefined || r[c[key]] === null ? '' : r[c[key]]]; });
      if (colVals.length) sh.getRange(2, c[key] + 1, colVals.length, 1).setValues(colVals);
    }
  });
  sync_shade_(sh, now, c);
  return list.length;
}

// ── HELPERS ─────────────────────────────────────────────────
/**
 * Settings from the top of this file. Script properties win if set
 * (SHOP_NAME, DRESS_API_URL, PARTNER_PIN), and the older HORDR_API_URL /
 * HORDR_PARTNER_PIN names still work, so House of Roan's copy keeps running.
 */
function sync_settings_() {
  var p = {}, shop = '', url = '', pin = '';
  try { p = PropertiesService.getScriptProperties().getProperties() || {}; } catch (e) {}
  Object.keys(p).forEach(function (k) {
    var n = String(k).replace(/\s+/g, '').toUpperCase();
    if (n === 'SHOP_NAME') shop = String(p[k] || '').trim();
    if (n === 'DRESS_API_URL' || n === 'HORDR_API_URL') url = String(p[k] || '').trim();
    if (n === 'PARTNER_PIN' || n === 'HORDR_PARTNER_PIN') pin = String(p[k] || '').trim();
  });
  return { shop: shop || String(SHOP_NAME_HERE || '').trim(),
           url: url || String(DRESS_API_URL_HERE || '').trim(),
           pin: pin || String(PARTNER_PIN_HERE || '').trim() };
}

/** Call the dress web app. Throws with the dress system's own message on failure. */
function api_(method, body) {
  var cfg = sync_settings_(), res;
  if (method === 'get') {
    var qs = Object.keys(body).map(function (k) { return k + '=' + encodeURIComponent(body[k]); }).join('&');
    res = UrlFetchApp.fetch(cfg.url + '?' + qs, { muteHttpExceptions: true, followRedirects: true });
  } else {
    body.pin = cfg.pin;
    res = UrlFetchApp.fetch(cfg.url, { method: 'post', contentType: 'text/plain;charset=utf-8',
      payload: JSON.stringify(body), muteHttpExceptions: true, followRedirects: true });
  }
  var j;
  try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('The dress system did not answer. Check DRESS_API_URL_HERE.'); }
  if (!j.ok) throw new Error(j.error);
  return j.data;
}

/** C/O for a row left blank: SYNC.DEFAULT_CO if it is one of yours, else your first channel (saved by the dropdown refresh). */
function sync_defaultCo_() {
  var mine = sync_parse_(PropertiesService.getDocumentProperties().getProperty('MY_COS')) || [];
  if (SYNC.DEFAULT_CO && (!mine.length || mine.indexOf(SYNC.DEFAULT_CO) >= 0)) return SYNC.DEFAULT_CO;
  return mine[0] || SYNC.DEFAULT_CO || 'Walk-in';
}

function sync_sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SYNC.TAB);
  if (!sh) throw new Error('No "' + SYNC.TAB + '" tab. Run setupSync first.');
  return sh;
}
function sync_head_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String); }
function sync_cols_(head) {
  var c = {};
  COL_ORDER.forEach(function (k) {
    c[k] = head.indexOf(COL[k]);
    if (c[k] < 0) throw new Error('The "' + SYNC.TAB + '" tab is missing the "' + COL[k] + '" column. Run setupSync.');
  });
  return c;
}

/** A sheet row → plain comparable values (dates as yyyy-MM-dd, numbers as numbers, text trimmed). */
function sync_read_(a, c) {
  var o = {};
  FIELDS.forEach(function (f) {
    var v = a[c[f]];
    if (DATE_FIELDS[f]) o[f] = sync_dateIso_(v);
    else if (f === 'price' || f === 'rate') o[f] = v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v);
    else o[f] = v === null || v === undefined ? '' : String(v).trim();
  });
  return o;
}

/** A rental from the dress system → the same plain values. */
function sync_fromServer_(r) {
  var num = function (v) { return v === '' || v === null || v === undefined || isNaN(Number(v)) ? '' : Number(v); };
  if (!r.mine) return { mine: false, renter: 'CLS client', dress: r.dress, event: r.pickup || '', pickup: r.pickup || '',
    ret: r.returnBy || '', status: r.status || '', co: 'CLS booking', rate: SYNC.CLS_PARTNERSHIP_RATE,
    pay: '', price: '', contact: '', notes: '' };
  return { mine: true, renter: r.renter || '', dress: r.dress || '', event: r.eventDate || '', pickup: r.pickup || '',
    ret: r.returnBy || '', status: r.status || '', co: r.co || '', rate: num(r.rate), pay: r.payStatus || '',
    price: num(r.clientPrice), contact: String(r.contact || ''), notes: String(r.notes || '') };
}

function sync_cell_(f, v) {
  if (DATE_FIELDS[f]) return v ? Utilities.parseDate(v, SYNC.TZ, 'yyyy-MM-dd') : '';
  return v === null || v === undefined ? '' : v;
}
function sync_dateIso_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, SYNC.TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim(), m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  if (s) { var d = new Date(s); if (!isNaN(d)) return Utilities.formatDate(d, SYNC.TZ, 'yyyy-MM-dd'); }
  return s;
}
function sync_iso_(d) { return Utilities.formatDate(d, SYNC.TZ, 'yyyy-MM-dd'); }
function sync_parse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function sync_clean_(o) { Object.keys(o).forEach(function (k) { if (o[k] === undefined || o[k] === '') delete o[k]; }); return o; }

function sync_formatColumns_(sh, head) {
  var n = Math.max(sh.getMaxRows() - 1, 1);
  ['event', 'pickup', 'ret'].forEach(function (k) { sh.getRange(2, head.indexOf(COL[k]) + 1, n, 1).setNumberFormat('mmm d, yyyy'); });
  ['rate', 'price'].forEach(function (k) { sh.getRange(2, head.indexOf(COL[k]) + 1, n, 1).setNumberFormat('₱#,##0'); });
}

/**
 * Dropdowns for Dress Name, Booking Status, C/O and payment Status, straight from the dress system.
 * Uses the signed-in lists (myDresses / myConfig), so each house only gets its OWN dresses and channels.
 */
function sync_dropdowns_(sh, head) {
  var n = Math.max(sh.getMaxRows() - 1, 1);
  var all = api_('post', { action: 'myDresses' });
  sync_priceList_(all);
  var dresses = all.map(function (d) { return d.name; });
  var cfg = api_('post', { action: 'myConfig' });
  var cos = Object.keys(cfg.coRules || {}).filter(function (k) { return !/^CLS/.test(k); });
  PropertiesService.getDocumentProperties().setProperty('MY_COS', JSON.stringify(cos));
  var lists = { dress: dresses, status: cfg.statuses, co: cos, pay: cfg.payStatuses };
  Object.keys(lists).forEach(function (k) {
    if (!lists[k] || !lists[k].length) return;
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(lists[k], true).setAllowInvalid(true).build();
    sh.getRange(2, head.indexOf(COL[k]) + 1, n, 1).setDataValidation(rule);
  });
}

/** Grey out CLS rows so it's clear they're read-only (only in the columns this script manages). */
function sync_shade_(sh, rows, c) {
  if (rows.length < 2) return;
  COL_ORDER.forEach(function (k) {
    var bg = rows.slice(1).map(function (r) { return [String(r[c.co] || '') === 'CLS booking' ? '#f1eeee' : null]; });
    sh.getRange(2, c[k] + 1, bg.length, 1).setBackgrounds(bg);
  });
}

/**
 * The "Price List" tab belongs to the house: type your price for each dress in the Price column.
 * The sync never overwrites a price you typed. It only adds dresses that are new in your collection
 * (and fills a blank price from the dress system when it has one).
 */
function sync_priceList_(dresses) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SYNC.PRICE_TAB) || ss.insertSheet(SYNC.PRICE_TAB);
  var v = sh.getLastColumn() ? sh.getDataRange().getValues() : [];
  if (!v.length || String(v[0][0]).trim() !== 'Dress Name') {
    sh.getRange(1, 1, 1, 2).setValues([['Dress Name', 'Price']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    v = [['Dress Name', 'Price']];
  }
  var have = {};
  v.slice(1).forEach(function (r, i) { have[String(r[0]).trim().toLowerCase()] = { row: i + 2, price: r[1] }; });
  var add = [];
  dresses.forEach(function (d) {
    var p = d.regularRate === null || d.regularRate === undefined ? '' : d.regularRate, h = have[d.name.toLowerCase()];
    if (!h) add.push([d.name, p]);
    else if ((h.price === '' || h.price === null) && p !== '') sh.getRange(h.row, 2).setValue(p);
  });
  if (add.length) sh.getRange(v.length + 1, 1, add.length, 2).setValues(add);
  sh.getRange(2, 2, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('₱#,##0');
}

/** Dress name (lowercase) → your price, read from the Price List tab. */
function sync_prices_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SYNC.PRICE_TAB), out = {};
  if (!sh || !sh.getLastColumn()) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var name = String(r[0] || '').trim().toLowerCase(), p = String(r[1] === null ? '' : r[1]).replace(/[₱Pp,\s]/g, '');
    if (name && p !== '' && !isNaN(Number(p))) out[name] = Number(p);
  });
  return out;
}
