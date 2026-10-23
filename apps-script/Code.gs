/**
 * HORDR × Click Lounge Studio — Dress Rental back end (Google Apps Script)
 *
 * Reads and writes the "HORDR Dress Rental Tracker" Google Sheet and keeps one
 * Google Calendar per rental house in sync (House of Roan, Serra, …), so each
 * house's calendar only shows its own gowns.
 *
 * Each rental house connects its OWN Google Sheet to this one with
 * Shop_Sync.gs (same folder), signed in with that house's partner PIN.
 *
 * PUBLIC (no PIN) — GET  ?action=...
 *   config                                   rental rules + status lists + the houses
 *   dresses       [&brand=Serra]             active dresses (no renter data)
 *   availability  &date=YYYY-MM-DD&co=CLS%20Prestige[&brand=Serra]   every dress: Available / Booked + price
 *   dressCalendar &dress=NAME&month=YYYY-MM&co=CLS%20Prestige   day-by-day for one dress
 *
 * TWO PINs — POST body (Content-Type: text/plain) JSON {action, pin, ...}
 *   STAFF_PIN   (CLS)     — everything below
 *   a partner PIN (one per rental house, in PARTNER_PINS) — whoami, myDresses,
 *                           listRentals and reserve / setStatus / updateRental,
 *                           all limited to that house's OWN dresses. A CLS booking
 *                           of one of their dresses shows as dates only; another
 *                           house's dresses and bookings are not visible at all.
 *   reserve       {dress, eventDate, co, renter, contact?, clsBookingId?, notes?, clientPrice?, rate?, holdHours?}
 *   setStatus     {rentalId, status, date?}      Confirmed / Picked Up / Returned / Cancelled
 *   updateRental  {rentalId, eventDate?, pickup?, returnDate?, renter?, contact?, notes?, rate?, payStatus?, co?}
 *   listRentals   {from?, to?}
 *
 * Run setup() once from the editor. It creates the calendars, adds helper
 * columns, generates the staff PIN and installs the 15-minute maintenance trigger.
 * listCalendars() prints each house's calendar, to share with that house.
 */

// ───────────────────────────── CONFIG ─────────────────────────────
var CONFIG = {
  SHEET_ID: '',                 // leave blank if this script is bound to the sheet (Extensions → Apps Script)
  TZ: 'Asia/Manila',
  CALENDAR_NAME: 'HORDR × CLS Dress Rentals',   // House of Roan's calendar; other houses get "<House> × CLS Dress Rentals"
  HOLD_HOURS: 48,               // how long a "Reserved" booking made through the app is held before auto-release
  PUBLIC_HOLDS: false,          // true = clients may place Reserved holds without a PIN
  SYNC_LOOKBACK_DAYS: 60,       // calendar sync ignores bookings that ended longer ago than this
  CALENDAR_PRIVACY: true,       // true = the shared calendars show no client names or money for CLS bookings
  TABS: { tracker: 'Tracker', dresses: 'Dresses', lists: 'Lists' },
  BLOCKING: ['Reserved', 'Confirmed', 'Picked Up'],
  STATUSES: ['Inquiry', 'Reserved', 'Confirmed', 'Picked Up', 'Returned', 'Cancelled'],
  PAY_STATUSES: ['Unpaid', 'Deposit Paid', 'Fully Paid', 'Refunded'],
  DEFAULT_CO: 'HORDR Direct',
  // The house a dress belongs to when its Brand cell is blank. Every dress on
  // the tab predates the column, so without this they would all become
  // ownerless the moment it is added.
  DEFAULT_BRAND: 'House of Roan',
  ADDON_LESS: 99,               // Package add-on price = Regular Rate − 99 …
  ADDON_MIN: 500                // … but never below ₱500
};

// Tracker column headers (matched by name, so column order can change)
var H = {
  renter: 'Name of Renter', dress: 'Dress Name', event: 'Event Date', pickup: 'Date of Pick-up',
  ret: 'Date of Return', status: 'Booking Status', co: 'C/O', rate: 'Partnership Rate',
  pay: 'Status', contact: 'Contact No.', clsId: 'CLS Booking ID', notes: 'Notes', clientPrice: 'Client Price',
  blockStart: 'Block Start', calId: 'Calendar Event ID',
  // added by setup()
  id: 'Rental ID', holdUntil: 'Hold Until', createdBy: 'Created By', syncKey: 'Calendar Sync Key'
};
var EXTRA_HEADERS = [H.clientPrice, H.contact, H.clsId, H.notes, H.id, H.holdUntil, H.createdBy, H.syncKey];
// columns that are nice to have: if one is missing, writes to it are skipped instead of failing
var OPTIONAL_HEADERS = [H.clientPrice, H.contact, H.clsId, H.notes, H.holdUntil, H.createdBy, H.rate, H.pay];

// ───────────────────────────── WEB ENTRY POINTS ─────────────────────────────
function doGet(e) {
  return respond_(function () {
    var p = (e && e.parameter) || {};
    switch (p.action) {
      case 'config':        return apiConfig_();
      case 'dresses':       return apiDresses_(p.brand);
      case 'availability':  return apiAvailability_(p.date, p.co, p.brand);
      case 'dressCalendar': return apiDressCalendar_(p.dress, p.month, p.co);
      default: throw new Error('Unknown action: ' + p.action);
    }
  });
}

function doPost(e) {
  return respond_(function () {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var isPublicHold = b.action === 'reserve' && CONFIG.PUBLIC_HOLDS;
    var ctx = isPublicHold && !roleFor_(b.pin) ? { role: 'public', brand: '' } : checkPin_(b.pin);
    switch (b.action) {
      case 'whoami':       return { role: ctx.role, brand: ctx.brand };
      case 'myDresses':    return apiMyDresses_(ctx);
      case 'myConfig':     return apiMyConfig_(ctx);
      case 'reserve':      return apiReserve_(b, ctx);
      case 'setStatus':    return apiSetStatus_(b, ctx);
      case 'updateRental': return apiUpdateRental_(b, ctx);
      case 'listRentals':  return apiListRentals_(b.from, b.to, ctx);
      default: throw new Error('Unknown action: ' + b.action);
    }
  });
}

function respond_(fn) {
  var out;
  try { out = { ok: true, data: fn() }; }
  catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────── PUBLIC API ─────────────────────────────
function apiConfig_() {
  var rules = loadRules_();
  return { cleaningDays: rules.cleaning, coRules: rules.co, statuses: CONFIG.STATUSES,
           payStatuses: CONFIG.PAY_STATUSES, holdHours: CONFIG.HOLD_HOURS, publicHolds: CONFIG.PUBLIC_HOLDS,
           houses: houses_() };
}

/** The houses that have at least one active dress, House of Roan first. The website's shop switch. */
function houses_() {
  var seen = {}, out = [];
  loadDresses_().forEach(function (d) { if (d.active && !seen[d.brand]) { seen[d.brand] = 1; out.push(d.brand); } });
  return out.sort(function (a, b) { return (b === CONFIG.DEFAULT_BRAND) - (a === CONFIG.DEFAULT_BRAND); });
}

/** Optional ?brand= filter on the public lists. Blank = every house. */
function brandMatch_(d, brand) { return !brand || d.brand === brand; }

function apiDresses_(brand) {
  return loadDresses_().filter(function (d) { return d.active && brandMatch_(d, brand); }).map(publicDress_);
}

function apiAvailability_(dateIso, co, brand) {
  var ev = requireDay_(dateIso, 'date');
  var rules = loadRules_(), coKey = coOrDefault_(co, rules);
  var win = windowFor_(ev, coKey, rules);
  var rentals = loadRentals_().rows, now = new Date();
  var list = loadDresses_().filter(function (d) { return d.active && brandMatch_(d, brand); }).map(function (d) {
    var chk = checkFree_(rentals, d.name, win.start, win.blockEnd, null, now);
    var o = publicDress_(d);
    o.rate = priceFor_(d, coKey, rules);
    o.status = chk.free ? 'Available' : 'Booked';
    o.nextFreeFrom = chk.free ? null : dayToIso_(chk.nextFree);
    return o;
  });
  return { eventDate: dayToIso_(ev), co: coKey, pickup: dayToIso_(win.start), returnBy: dayToIso_(win.ret),
           available: list.filter(function (x) { return x.status === 'Available'; }).length, dresses: list };
}

function apiDressCalendar_(dressName, month, co) {
  var dress = findDress_(dressName);
  var m = /^(\d{4})-(\d{2})$/.exec(month || '');
  if (!m) throw new Error('month must be YYYY-MM');
  var rules = loadRules_(), coKey = coOrDefault_(co, rules);
  var first = Date.UTC(+m[1], +m[2] - 1, 1) / 864e5, last = Date.UTC(+m[1], +m[2], 0) / 864e5;
  var today = todayDay_(), now = new Date();
  var mine = loadRentals_().rows.filter(function (r) { return r.dress === dress.name && isActive_(r, now); });
  var days = [];
  for (var d = first; d <= last; d++) {
    var booked = mine.some(function (r) { return d >= r.start && d <= r.blockEnd; });
    var w = windowFor_(d, coKey, rules);
    var canBook = dress.active && d >= today && w.start >= today &&
                  !mine.some(function (r) { return w.start <= r.blockEnd && w.blockEnd >= r.start; });
    days.push({ date: dayToIso_(d), booked: booked, canBook: canBook });
  }
  return { dress: publicDress_(dress), month: month, co: coKey, days: days };
}

// ───────────────────────────── STAFF API ─────────────────────────────
function apiReserve_(b, ctx) {
  var role = ctx.role;
  var rules = loadRules_();
  var dress = findDress_(b.dress);
  if (!dress.active) throw new Error('This dress is not active.');
  // A house books its own gowns and no one else's.
  if (role === 'partner' && dress.brand !== ctx.brand) throw new Error('Dress not found: ' + b.dress);
  if (role === 'partner' && isCls_(b.co)) throw new Error('Only Click Lounge Studio can book the CLS packages.');
  if (role === 'partner' && b.co && !coAllowed_(ctx, b.co, rules)) throw new Error('co must be one of: ' + myCos_(ctx, rules).join(', '));
  var coKey = coOrDefault_(b.co, rules);
  if (!b.co || !rules.co[b.co]) throw new Error('co must be one of: ' + Object.keys(rules.co).join(', '));
  if (!b.renter || !String(b.renter).trim()) throw new Error('renter is required.');
  if (role === 'public' && !b.contact) throw new Error('contact is required.');
  var ev = requireDay_(b.eventDate, 'eventDate');
  var win = computeWindow_(ev, b.pickup ? requireDay_(b.pickup, 'pickup') : null,
                           b.returnDate ? requireDay_(b.returnDate, 'returnDate') : null, coKey, rules);
  if (win.start < todayDay_()) throw new Error('Pick-up date is in the past.');
  if (win.ret < win.start) throw new Error('Return date is before pick-up date.');

  return withLock_(function () {
    var t = loadRentals_(), now = new Date();
    var chk = checkFree_(t.rows, dress.name, win.start, win.blockEnd, null, now);
    if (!chk.free) throw new Error(clashMsg_(dress.name, chk, coKey, rules, role));
    var rowNum = nextEmptyRow_(t);
    var id = newId_();
    // staff calls (e.g. the POS booking form) may set their own hold to match their payment window
    var hrs = role === 'staff' && Number(b.holdHours) > 0 ? Math.min(Number(b.holdHours), 168) : CONFIG.HOLD_HOURS;
    var hold = new Date(now.getTime() + hrs * 3600e3);
    var vals = {};
    vals[H.id] = id; vals[H.renter] = String(b.renter).trim(); vals[H.dress] = dress.name;
    vals[H.event] = dayToDate_(ev);
    vals[H.pickup] = b.pickup ? dayToDate_(win.start) : '';
    vals[H.ret] = b.returnDate ? dayToDate_(win.ret) : '';
    vals[H.status] = 'Reserved'; vals[H.co] = coKey;
    var partner = partnerCostFor_(coKey, rules, dress.brand);
    vals[H.rate] = b.rate !== undefined && b.rate !== '' ? Number(b.rate) : (partner === null ? '' : partner);
    vals[H.clientPrice] = b.clientPrice !== undefined && b.clientPrice !== '' ? Number(b.clientPrice) : priceFor_(dress, coKey, rules);
    vals[H.pay] = 'Unpaid'; vals[H.contact] = b.contact || ''; vals[H.clsId] = b.clsBookingId || '';
    vals[H.notes] = b.notes || ''; vals[H.holdUntil] = hold;
    vals[H.createdBy] = role === 'partner' ? ctx.brand : role === 'public' ? 'Public' : 'CLS staff';
    writeRow_(t, rowNum, vals);
    var r = rowFromSheet_(t, rowNum);
    syncRowToCalendar_(t, r);
    return { rentalId: id, row: rowNum, status: 'Reserved', holdUntil: hold.toISOString(),
             pickup: dayToIso_(r.start), returnBy: dayToIso_(r.ret) };
  });
}

function apiSetStatus_(b, ctx) {
  var role = ctx.role;
  if (CONFIG.STATUSES.indexOf(b.status) < 0) throw new Error('status must be one of: ' + CONFIG.STATUSES.join(', '));
  return withLock_(function () {
    var t = loadRentals_(), r = findRental_(t, b.rentalId), now = new Date(), today = todayDay_();
    requireOwn_(r, ctx, t.ix);
    var when = b.date ? requireDay_(b.date, 'date') : today;
    var vals = {};
    vals[H.status] = b.status;
    if (b.status !== 'Reserved') vals[H.holdUntil] = '';
    if (CONFIG.BLOCKING.indexOf(b.status) >= 0 && CONFIG.BLOCKING.indexOf(r.status) < 0) {
      var chk = checkFree_(t.rows, r.dress, r.start, r.blockEnd, r.id, now);
      if (!chk.free) throw new Error('Cannot re-activate: ' + clashMsg_(r.dress, chk, r.co, loadRules_(), role).replace('Sorry, ', ''));
    }
    if (b.status === 'Picked Up' && r.pickupRaw === null) vals[H.pickup] = dayToDate_(when);
    if (b.status === 'Returned' && (r.retRaw === null || r.retRaw > when)) vals[H.ret] = dayToDate_(when);
    writeRow_(t, r.row, vals);
    var fresh = rowFromSheet_(t, r.row);
    syncRowToCalendar_(t, fresh);
    return rentalOut_(fresh, ctx);
  });
}

function apiUpdateRental_(b, ctx) {
  var role = ctx.role;
  var rules = loadRules_();
  return withLock_(function () {
    var t = loadRentals_(), r = findRental_(t, b.rentalId), vals = {};
    requireOwn_(r, ctx, t.ix);
    if (role === 'partner' && b.co !== undefined && isCls_(b.co)) throw new Error('Only Click Lounge Studio can book the CLS packages.');
    if (b.renter !== undefined) vals[H.renter] = b.renter;
    if (b.contact !== undefined) vals[H.contact] = b.contact;
    if (b.notes !== undefined) vals[H.notes] = b.notes;
    if (b.rate !== undefined) vals[H.rate] = b.rate === '' ? '' : Number(b.rate);
    if (b.clientPrice !== undefined) vals[H.clientPrice] = b.clientPrice === '' ? '' : Number(b.clientPrice);
    if (b.clsBookingId !== undefined) vals[H.clsId] = b.clsBookingId;
    if (b.payStatus !== undefined) {
      if (CONFIG.PAY_STATUSES.indexOf(b.payStatus) < 0) throw new Error('payStatus must be one of: ' + CONFIG.PAY_STATUSES.join(', '));
      vals[H.pay] = b.payStatus;
    }
    if (b.co !== undefined) {
      if (!rules.co[b.co]) throw new Error('Unknown co.');
      if (role === 'partner' && !coAllowed_(ctx, b.co, rules)) throw new Error('co must be one of: ' + myCos_(ctx, rules).join(', '));
      vals[H.co] = b.co;
    }
    if (b.eventDate !== undefined) vals[H.event] = dayToDate_(requireDay_(b.eventDate, 'eventDate'));
    if (b.pickup !== undefined) vals[H.pickup] = b.pickup === '' ? '' : dayToDate_(requireDay_(b.pickup, 'pickup'));
    if (b.returnDate !== undefined) vals[H.ret] = b.returnDate === '' ? '' : dayToDate_(requireDay_(b.returnDate, 'returnDate'));

    // Re-check the dress is still free with the new dates
    var ev = vals[H.event] !== undefined ? dayNum_(vals[H.event]) : r.event;
    var pu = vals[H.pickup] !== undefined ? dayNum_(vals[H.pickup]) : r.pickupRaw;
    var rt = vals[H.ret] !== undefined ? dayNum_(vals[H.ret]) : r.retRaw;
    var co = vals[H.co] !== undefined ? vals[H.co] : r.co;
    var w = computeWindow_(ev, pu, rt, co, rules);
    if (!w) throw new Error('Event Date or Pick-up date is required.');
    if (w.ret < w.start) throw new Error('Return date is before pick-up date.');
    if (CONFIG.BLOCKING.indexOf(r.status) >= 0) {
      var chk = checkFree_(t.rows, r.dress, w.start, w.blockEnd, r.id, new Date());
      if (!chk.free) throw new Error(clashMsg_(r.dress, chk, co, rules, role).replace('Sorry, ', 'Those dates clash: '));
    }
    writeRow_(t, r.row, vals);
    var fresh = rowFromSheet_(t, r.row);
    syncRowToCalendar_(t, fresh);
    return rentalOut_(fresh, ctx);
  });
}

function apiListRentals_(fromIso, toIso, ctx) {
  var from = fromIso ? requireDay_(fromIso, 'from') : todayDay_() - 7;
  var to = toIso ? requireDay_(toIso, 'to') : todayDay_() + 90;
  var now = new Date(), t = loadRentals_(), ix = t.ix;
  return t.rows
    .filter(function (r) { return r.start !== null && r.start <= to && r.blockEnd >= from; })
    .filter(function (r) { return ownsRental_(ctx, r, ix); })
    .sort(function (a, b) { return a.start - b.start; })
    .map(function (r) { var o = rentalOut_(r, ctx); o.active = isActive_(r, now); o.brand = brandOfRental_(r, ix); return o; });
}

/** The channels this caller may book under. */
function myCos_(ctx, rules) {
  return Object.keys(rules.co).filter(function (k) { return coAllowed_(ctx, k, rules); });
}

/**
 * `config` for a signed-in house: the same rules, with the C/O list cut down
 * to their own channels. The public `config` is unchanged — it has to list
 * the CLS packages, because the wardrobe page prices against them.
 */
function apiMyConfig_(ctx) {
  var rules = loadRules_(), out = apiConfig_();
  if (ctx.role === 'partner') {
    var mine = {};
    myCos_(ctx, rules).forEach(function (k) { mine[k] = out.coRules[k]; });
    out.coRules = mine;
    out.houses = [ctx.brand];
  }
  out.brand = ctx.brand || '';
  return out;
}

/** The calling house's own dresses. Partners use this instead of `dresses`,
 *  which is the public wardrobe and quite rightly shows every house's gowns. */
function apiMyDresses_(ctx) {
  var all = loadDresses_().filter(function (d) { return d.active; });
  if (ctx.role === 'partner') all = all.filter(function (d) { return d.brand === ctx.brand; });
  return all.map(publicDress_);
}

// ───────────────────────────── MAINTENANCE (trigger) ─────────────────────────────
/** Runs every 15 min: gives IDs to hand-typed rows, releases expired holds, syncs the calendars. */
function maintenance() {
  withLock_(function () {
    var t = loadRentals_(), now = new Date(), cutoff = todayDay_() - CONFIG.SYNC_LOOKBACK_DAYS;
    var released = 0, synced = 0;
    t.rows.forEach(function (r) {
      var vals = {};
      if (!r.id) { vals[H.id] = newId_(); r.id = vals[H.id]; }
      if (r.status === 'Reserved' && r.holdUntil && r.holdUntil < now) {
        vals[H.status] = 'Cancelled'; vals[H.holdUntil] = '';
        vals[H.notes] = (r.notes ? r.notes + ' · ' : '') + 'Auto-released: hold expired ' + Utilities.formatDate(now, CONFIG.TZ, 'MMM d');
        released++;
      }
      if (Object.keys(vals).length) writeRow_(t, r.row, vals);
      var fresh = Object.keys(vals).length ? rowFromSheet_(t, r.row) : r;
      if (fresh.start !== null && fresh.blockEnd >= cutoff) { if (syncRowToCalendar_(t, fresh)) synced++; }
      else if (fresh.calId && fresh.start === null) { if (syncRowToCalendar_(t, fresh)) synced++; }
    });
    console.log('maintenance: released ' + released + ' hold(s), updated ' + synced + ' calendar event(s)');
  });
}

// ───────────────────────────── SETUP (run once) ─────────────────────────────
function setup() {
  var props = PropertiesService.getScriptProperties();
  // 1. Calendars — one per house (House of Roan keeps the one it already has)
  calendar_(CONFIG.DEFAULT_BRAND);
  // 2. Helper columns
  var sh = ss_().getSheetByName(CONFIG.TABS.tracker);
  if (!sh) throw new Error('Tab "' + CONFIG.TABS.tracker + '" not found.');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
  var fmtFrom = headers.indexOf(H.calId) + 1;
  EXTRA_HEADERS.forEach(function (h) {
    if (headers.indexOf(h) >= 0) return;
    var col = headers.length + 1;
    if (col > sh.getMaxColumns()) sh.insertColumnAfter(sh.getMaxColumns());
    sh.getRange(1, col).setValue(h);
    if (fmtFrom > 0) sh.getRange(1, fmtFrom).copyFormatToRange(sh, col, col, 1, 1);
    headers.push(h);
  });
  var holdCol = headers.indexOf(H.holdUntil) + 1;
  sh.getRange(2, holdCol, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.hideColumns(headers.indexOf(H.syncKey) + 1);
  // 2b. Brand column on the Dresses tab (blank = CONFIG.DEFAULT_BRAND)
  var dsh = ss_().getSheetByName(CONFIG.TABS.dresses);
  if (dsh) {
    var dh = dsh.getRange(1, 1, 1, dsh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
    if (dh.indexOf('Brand') < 0) {
      var dcol = dh.length + 1;
      if (dcol > dsh.getMaxColumns()) dsh.insertColumnAfter(dsh.getMaxColumns());
      dsh.getRange(1, dcol).setValue('Brand');
    }
  }
  // 3. PINs — one for CLS, one per rental house
  if (!props.getProperty('STAFF_PIN')) props.setProperty('STAFF_PIN', String(Math.floor(100000 + Math.random() * 900000)));
  var pins = partnerPins_();
  if (!Object.keys(pins).length) {
    pins[CONFIG.DEFAULT_BRAND] = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('PARTNER_PINS', JSON.stringify(pins));
  }
  // 1b. now that the houses are known, make sure each has its calendar
  allBrands_().forEach(function (b) { calendar_(b); });
  // 4. Trigger
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'maintenance') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('maintenance').timeBased().everyMinutes(15).create();
  // 5. First sync
  maintenance();
  var names = Object.keys(partnerPins_());
  console.log('✅ Setup done.' +
              '\nCLS staff PIN: ' + props.getProperty('STAFF_PIN') +
              '\nRental houses with a PIN: ' + (names.join(', ') || 'none') +
              '\n(Read or change them with listPartners() and setMyPartnerPin().)\n');
  listCalendars();
}

/** Quick check from the editor: select and Run, then read the Execution log. */
function testReadOnly() {
  var d = apiDresses_();
  console.log(d.length + ' active dress(es): ' + d.map(function (x) { return x.name + ' (' + x.brand + ')'; }).join(', '));
  var date = Utilities.formatDate(new Date(Date.now() + 7 * 864e5), CONFIG.TZ, 'yyyy-MM-dd');
  ['CLS Prestige', 'CLS Photographer', CONFIG.DEFAULT_CO].forEach(function (co) {
    var a = apiAvailability_(date, co);
    console.log(co + ' on ' + date + ' (pick-up ' + a.pickup + ', return ' + a.returnBy + '): ' + a.available + ' available');
  });
}

// ───────────────────────────── CALENDARS (one per house) ─────────────────────────────
var STATUS_COLOR = { 'Reserved': 'YELLOW', 'Confirmed': 'GREEN', 'Picked Up': 'BLUE', 'Returned': 'GRAY' };

/** Creates/replaces/deletes the calendar event for one row. Returns true if the calendar changed. */
function syncRowToCalendar_(t, r) {
  var brand = brandOfRental_(r, t.ix);
  var want = r.start !== null && (STATUS_COLOR[r.status] !== undefined) &&
             !(r.status === 'Reserved' && r.holdUntil && r.holdUntil < new Date());
  // the house is part of the key, so a row whose dress moves to another house moves calendar too
  var key = want ? [r.status, r.dress, r.renter, r.co, dayToIso_(r.start), dayToIso_(r.ret), CONFIG.CALENDAR_PRIVACY, brand].join('|') : '';
  if (key === (r.syncKey || '') && (!want || r.calId)) return false;
  if (r.calId) deleteCalendarEvent_(r.calId);
  var vals = {}; vals[H.calId] = ''; vals[H.syncKey] = '';
  if (want) {
    var prefix = r.status === 'Reserved' ? '[HOLD] ' : r.status === 'Returned' ? '✔ ' : '';
    var hide = CONFIG.CALENDAR_PRIVACY && isCls_(r.co);   // each calendar is shared with its house
    var title = prefix + '👗 ' + r.dress + ' · ' + (hide ? 'CLS booking' : (r.renter || '—') + ' (' + (r.co || '—') + ')');
    var ev = calendar_(brand).createAllDayEvent(title, dayToDate_(r.start), dayToDate_(r.ret + 1), {
      description: [
        'Status: ' + r.status + (hide || !r.pay ? '' : ' · Payment: ' + r.pay),
        'Pick-up: ' + dayToIso_(r.start) + '  ·  Return: ' + dayToIso_(r.ret),
        'Available again: ' + dayToIso_(r.blockEnd + 1),
        hide ? '' : 'Event date: ' + (r.event !== null ? dayToIso_(r.event) : '—'),
        hide || !r.contact ? '' : 'Contact: ' + r.contact,
        hide || !r.clsId ? '' : 'CLS Booking: ' + r.clsId,
        hide || !r.notes ? '' : 'Notes: ' + r.notes,
        'Rental ID: ' + (r.id || '')
      ].filter(String).join('\n')
    });
    ev.setColor(CalendarApp.EventColor[STATUS_COLOR[r.status]]);
    vals[H.calId] = ev.getId(); vals[H.syncKey] = key;
  }
  writeRow_(t, r.row, vals);
  return true;
}

/** House of Roan keeps the original CALENDAR_ID property; every other house gets its own. */
function calendarProp_(brand) {
  return !brand || brand === CONFIG.DEFAULT_BRAND ? 'CALENDAR_ID' : 'CALENDAR_ID:' + brand;
}

/** The house's calendar, created the first time a booking needs it. */
function calendar_(brand) {
  var props = PropertiesService.getScriptProperties(), key = calendarProp_(brand);
  var id = props.getProperty(key), cal = id && CalendarApp.getCalendarById(id);
  if (cal) return cal;
  var own = key === 'CALENDAR_ID';
  cal = CalendarApp.createCalendar(own ? CONFIG.CALENDAR_NAME : brand + ' × CLS Dress Rentals',
                                   { timeZone: CONFIG.TZ, color: own ? CalendarApp.Color.PINK : CalendarApp.Color.PLUM });
  props.setProperty(key, cal.getId());
  return cal;
}

/** An event may sit in any house's calendar (rows made before the split all live in House of Roan's). */
function deleteCalendarEvent_(eventId) {
  var props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).forEach(function (k) {
    if (k !== 'CALENDAR_ID' && k.indexOf('CALENDAR_ID:') !== 0) return;
    try {
      var cal = CalendarApp.getCalendarById(props[k]), ev = cal && cal.getEventById(eventId);
      if (ev) ev.deleteEvent();
    } catch (e) { /* already gone */ }
  });
}

/** Every house that has a PIN or a dress. */
function allBrands_() {
  var seen = {};
  seen[CONFIG.DEFAULT_BRAND] = 1;
  Object.keys(partnerPins_()).forEach(function (b) { seen[b] = 1; });
  loadDresses_().forEach(function (d) { seen[d.brand] = 1; });
  return Object.keys(seen);
}

/**
 * ▶ Run to see each house's calendar. Share each one with that house only:
 * Google Calendar → the calendar's ⋮ → Settings and sharing → Share with specific people.
 */
function listCalendars() {
  var lines = ['Calendars (share each one with its own house only):'];
  allBrands_().forEach(function (b) {
    var cal = calendar_(b);
    lines.push('  • ' + b + ' — "' + cal.getName() + '"  ID: ' + cal.getId());
  });
  var out = lines.join('\n');
  console.log(out);
  return out;
}

// ───────────────────────────── DATA LOADING ─────────────────────────────
function ss_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function loadRules_() {
  var v = ss_().getSheetByName(CONFIG.TABS.lists).getDataRange().getValues();
  var head = v[0].map(function (x) { return String(x).trim(); }), rules = { cleaning: 0, partnerRate: null, co: {} };
  var daysCol = head.indexOf('Rental Days'), leadCol = head.indexOf('Pick-up Days Before Event'), rateCol = head.indexOf('CLS Flat Rate');
  var houseCol = head.indexOf('House');
  if (daysCol < 1 || leadCol < 0) throw new Error('Lists tab is missing the C/O rules table (Rental Days / Pick-up Days Before Event).');
  rules.partnerRateBy = {};
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][6]).indexOf('Cleaning days') === 0) rules.cleaning = Number(v[i][7]) || 0;
    // "Partnership rate" is the one every house gets. "Partnership rate: Serra"
    // (or "Partnership rate — Serra") sets one house's own deal.
    if (String(v[i][6]).indexOf('Partnership rate') === 0) {
      var who = String(v[i][6]).slice('Partnership rate'.length).replace(/^[\s:—–-]+/, '').trim();
      if (who) rules.partnerRateBy[who] = numOrNull_(v[i][7]);
      else rules.partnerRate = numOrNull_(v[i][7]);
    }
    var co = String(v[i][daysCol - 1] || '').trim();
    if (co) rules.co[co] = { days: Math.max(1, Number(v[i][daysCol]) || 1), lead: Number(v[i][leadCol]) || 0,
                             rate: rateCol >= 0 ? numOrNull_(v[i][rateCol]) : null, // null = use the dress's Regular Rate
                             // Optional "House" column: which rental house this
                             // channel belongs to. Blank = everyone's (Referral,
                             // Walk-in), so leaving it empty keeps today's behaviour.
                             house: houseCol >= 0 ? String(v[i][houseCol] || '').trim() : '' };
  }
  if (!Object.keys(rules.co).length) throw new Error('No C/O rules found on the Lists tab.');
  return rules;
}

function loadDresses_() {
  var v = ss_().getSheetByName(CONFIG.TABS.dresses).getDataRange().getValues();
  var h = v[0].map(function (x) { return String(x).trim(); }), c = function (n) { return h.indexOf(n); };
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][c('Dress Name')] || '').trim();
    if (!name) continue;
    out.push({
      name: name, category: v[i][c('Category')] || '', color: v[i][c('Color')] || '', size: v[i][c('Size')] || '',
      regularRate: numOrNull_(v[i][c('Regular Rate')]), photographerRate: addOnRate_(v[i][c('Photographer Rate')], v[i][c('Regular Rate')]),
      photo: v[i][c('Photo URL')] || '', active: String(v[i][c('Active')]).trim().toLowerCase() === 'yes',
      notes: v[i][c('Notes')] || '',
      brand: String((c('Brand') >= 0 ? v[i][c('Brand')] : '') || '').trim() || CONFIG.DEFAULT_BRAND
    });
  }
  return out;
}

function loadRentals_() {
  var sh = ss_().getSheetByName(CONFIG.TABS.tracker);
  var v = sh.getDataRange().getValues();
  var head = v[0].map(function (x) { return String(x).trim(); }), col = {};
  Object.keys(H).forEach(function (k) { col[k] = head.indexOf(H[k]); });
  ['renter', 'dress', 'event', 'pickup', 'ret', 'status', 'co'].forEach(function (k) {
    if (col[k] < 0) throw new Error('Tracker is missing the "' + H[k] + '" column.');
  });
  var t = { sheet: sh, head: head, col: col, values: v, rules: loadRules_(), ix: brandIndex_(), rows: [] };
  for (var i = 1; i < v.length; i++) {
    var r = parseRow_(t, i + 1, v[i]);
    if (r) t.rows.push(r);
  }
  return t;
}

function parseRow_(t, rowNum, a) {
  var c = t.col, g = function (k) { return c[k] >= 0 ? a[c[k]] : ''; };
  var dress = String(g('dress') || '').trim(), renter = String(g('renter') || '').trim();
  if (!dress && !renter && !g('calId')) return null;
  var ev = dayNum_(g('event')), pu = dayNum_(g('pickup')), rt = dayNum_(g('ret'));
  var co = String(g('co') || '').trim();
  var w = dress ? computeWindow_(ev, pu, rt, co, t.rules) : null;
  var hold = g('holdUntil');
  return {
    row: rowNum, id: String(g('id') || ''), renter: renter, dress: dress, event: ev, pickupRaw: pu, retRaw: rt,
    status: String(g('status') || '').trim(), co: co, rate: g('rate'), clientPrice: g('clientPrice'), pay: String(g('pay') || ''),
    contact: String(g('contact') || ''), clsId: String(g('clsId') || ''), notes: String(g('notes') || ''),
    calId: String(g('calId') || ''), syncKey: String(g('syncKey') || ''),
    holdUntil: hold instanceof Date ? hold : null, createdBy: String(g('createdBy') || ''),
    start: w ? w.start : null, ret: w ? w.ret : null, blockEnd: w ? w.blockEnd : null
  };
}

function rowFromSheet_(t, rowNum) {
  var a = t.sheet.getRange(rowNum, 1, 1, t.head.length).getValues()[0];
  var r = parseRow_(t, rowNum, a);
  var idx = t.rows.findIndex(function (x) { return x.row === rowNum; });
  if (r) { if (idx >= 0) t.rows[idx] = r; else t.rows.push(r); }
  return r;
}

// ───────────────────────────── RULES ─────────────────────────────
/** Same logic as the sheet's Block Start / Block End formulas. */
function computeWindow_(ev, pickup, ret, co, rules) {
  var rule = rules.co[co] || rules.co[CONFIG.DEFAULT_CO] || rules.co[Object.keys(rules.co)[0]];
  var start = pickup !== null ? pickup : ev !== null ? ev - rule.lead : null;
  if (start === null) return null;
  var r = ret !== null ? ret : start + rule.days - 1;
  return { start: start, ret: r, blockEnd: r + rules.cleaning };
}

function windowFor_(eventDay, co, rules) { return computeWindow_(eventDay, null, null, co, rules); }

function isActive_(r, now) {
  if (r.start === null || CONFIG.BLOCKING.indexOf(r.status) < 0) return false;
  return !(r.status === 'Reserved' && r.holdUntil && r.holdUntil < now);
}

/** Is `dress` free from `start` to `blockEnd` (inclusive)? */
function checkFree_(rows, dress, start, blockEnd, excludeId, now) {
  var clashes = rows.filter(function (r) {
    return r.dress === dress && (!excludeId || r.id !== excludeId) && isActive_(r, now) &&
           start <= r.blockEnd && blockEnd >= r.start;
  });
  if (!clashes.length) return { free: true };
  return { free: false, clashes: clashes, nextFree: Math.max.apply(null, clashes.map(function (r) { return r.blockEnd; })) + 1 };
}

/** A clear "why it's taken" message: which booking, its dates, the cleaning day, and the earliest event date that works. */
function clashMsg_(dress, chk, co, rules, role) {
  var md = function (n) { return Utilities.formatDate(dayToDate_(n), CONFIG.TZ, 'MMM d'); };
  var parts = chk.clashes.map(function (r) {
    var who = role === 'staff' ? (r.renter || 'a booking') + (r.id ? ' (' + r.id + ')' : '')
            : isCls_(r.co) ? 'a CLS booking' : (r.renter || 'another booking');
    var out = r.start === r.ret ? md(r.start) : md(r.start) + '–' + md(r.ret);
    return who + ': out ' + out + (r.blockEnd > r.ret ? ', cleaning until ' + md(r.blockEnd) : '');
  });
  var rule = rules.co[co] || rules.co[CONFIG.DEFAULT_CO];
  var earliestEvent = chk.nextFree + (rule ? rule.lead : 0);
  return 'Sorry, ' + dress + ' is already booked around those dates — ' + parts.join('; ') + '. ' +
         'Earliest pick-up is ' + md(chk.nextFree) + ', so the earliest event date is ' + md(earliestEvent) +
         ' (' + dayToIso_(earliestEvent) + ').';
}

// ───────────────────────────── WRITING ─────────────────────────────
function writeRow_(t, rowNum, vals) {
  Object.keys(vals).forEach(function (name) {
    var ci = t.head.indexOf(name);
    if (ci < 0) {
      if (OPTIONAL_HEADERS.indexOf(name) >= 0) { console.warn('Tracker has no "' + name + '" column — run setup() to add it.'); return; }
      throw new Error('Column "' + name + '" not found. Run setup().');
    }
    t.sheet.getRange(rowNum, ci + 1).setValue(vals[name]);
  });
}

/** First row after the last booking. Extends the sheet (and its formulas) if needed. */
function nextEmptyRow_(t) {
  var last = 1;
  for (var i = t.values.length - 1; i >= 1; i--) {
    var a = t.values[i];
    if (String(a[t.col.renter] || '').trim() || String(a[t.col.dress] || '').trim()) { last = i + 1; break; }
  }
  var target = last + 1, sh = t.sheet;
  if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
  var bs = t.col.blockStart;
  if (bs >= 0 && !sh.getRange(target, bs + 1).getFormula()) {
    sh.getRange(2, bs + 1, 1, 4).copyTo(sh.getRange(target, bs + 1, 1, 4)); // Block Start … Conflict formulas
  }
  // keep t.values in sync so a second write in the same run doesn't pick the same row
  while (t.values.length < target) t.values.push(new Array(t.head.length).fill(''));
  t.values[target - 1][t.col.renter] = '(new)';
  return target;
}

// ───────────────────────────── HELPERS ─────────────────────────────
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy. Please try again in a few seconds.');
  try { var out = fn(); SpreadsheetApp.flush(); return out; } finally { lock.releaseLock(); }
}

/**
 * Who is calling: { role: 'staff' | 'partner' | 'public', brand: '' | 'House of Roan' | … }
 *
 * Partner PINs live in the PARTNER_PINS script property as JSON, one per
 * house:  {"House of Roan":"123456","Serra":"654321"}
 * The older single PARTNER_PIN still works and means CONFIG.DEFAULT_BRAND, so
 * nothing breaks the moment this file is pasted in.
 */
function partnerPins_() {
  var props = PropertiesService.getScriptProperties(), out = {};
  var raw = props.getProperty('PARTNER_PINS');
  if (raw) { try { out = JSON.parse(raw) || {}; } catch (e) { out = {}; } }
  var legacy = props.getProperty('PARTNER_PIN');
  if (legacy && !Object.keys(out).some(function (b) { return out[b] === legacy; })) {
    out[CONFIG.DEFAULT_BRAND] = legacy;
  }
  return out;
}

function roleFor_(pin) {
  var props = PropertiesService.getScriptProperties(), p = String(pin || '');
  if (!p) return null;
  if (p === props.getProperty('STAFF_PIN')) return { role: 'staff', brand: '' };
  var pins = partnerPins_();
  var brands = Object.keys(pins);
  for (var i = 0; i < brands.length; i++) {
    // Compared one by one rather than by lookup, so an empty or missing PIN
    // for one house can never match an empty PIN typed by someone else.
    if (String(pins[brands[i]] || '') && p === String(pins[brands[i]])) {
      return { role: 'partner', brand: brands[i] };
    }
  }
  return null;
}

function checkPin_(pin) {
  var cache = CacheService.getScriptCache(), fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw new Error('Too many wrong PIN attempts. Try again in 10 minutes.');
  var ctx = roleFor_(pin);
  if (!ctx) { cache.put('pinFails', String(fails + 1), 600); throw new Error('Wrong PIN.'); }
  return ctx;
}

/** A CLS booking — the partner may see the dates but not who or how much. */
function isCls_(co) { return String(co || '').indexOf('CLS') === 0; }

/** May this caller book under this channel? A house gets its own and the
 *  shared ones; a channel belonging to another house is not even listed. */
function coAllowed_(ctx, co, rules) {
  if (ctx.role !== 'partner') return true;
  if (isCls_(co)) return false;
  var rule = rules.co[co];
  if (!rule) return false;
  return !rule.house || rule.house === ctx.brand;
}

/**
 * Two separate questions, and they used to be one.
 *   - is this rental on one of MY dresses?   → ownsRental_
 *   - is it a CLS booking of my dress?       → mine to see, not mine to change
 * A rental on another house's dress is neither, and is refused as "not found"
 * rather than "belongs to someone else", which would confirm it exists.
 */
function requireOwn_(r, ctx, ix) {
  if (ctx.role !== 'partner') return;
  if (!ownsRental_(ctx, r, ix)) throw new Error('Rental not found: ' + (r.id || ''));
  if (isCls_(r.co)) throw new Error('That booking belongs to Click Lounge Studio.');
}

function findDress_(name) {
  var n = String(name || '').trim().toLowerCase();
  var d = loadDresses_().filter(function (x) { return x.name.toLowerCase() === n; })[0];
  if (!d) throw new Error('Dress not found: ' + name);
  return d;
}

function findRental_(t, id) {
  var r = t.rows.filter(function (x) { return x.id && x.id === String(id || ''); })[0];
  if (!r) throw new Error('Rental not found: ' + id);
  return r;
}

/**
 * What the CLIENT pays for one dress under one C/O:
 *   1. a flat rate on the Lists tab wins   (CLS Prestige = 0, i.e. the dress is included)
 *   2. any other CLS package  → the dress's Photographer Rate
 *   3. everyone else (the house's own renters) → the dress's Regular Rate
 */
function priceFor_(dress, co, rules) {
  var r = rules.co[co];
  if (r && r.rate !== null && r.rate !== undefined) return r.rate;
  if (String(co || '').indexOf('CLS') === 0 && dress.photographerRate !== null) return dress.photographerRate;
  return dress.regularRate;
}

/**
 * What CLS owes the rental house: a flat partnership rate on every CLS
 * booking, nothing on the house's own renters. Each house may have its own
 * deal ("Partnership rate: Serra" on the Lists tab); without one it falls
 * back to the single rate that covered everyone before there was more than
 * one house.
 */
function partnerCostFor_(co, rules, brand) {
  if (String(co || '').indexOf('CLS') !== 0) return null;
  var by = rules.partnerRateBy || {};
  if (brand && by[brand] !== undefined && by[brand] !== null) return by[brand];
  return rules.partnerRate;
}

function coOrDefault_(co, rules) { return rules.co[co] ? co : (rules.co[CONFIG.DEFAULT_CO] ? CONFIG.DEFAULT_CO : Object.keys(rules.co)[0]); }

function publicDress_(d) {
  return { name: d.name, category: d.category, color: d.color, size: d.size, regularRate: d.regularRate,
           photographerRate: d.photographerRate, photo: d.photo, notes: d.notes, brand: d.brand };
}

/** Dress name (lowercased) → the house that owns it. Built once per call. */
function brandIndex_() {
  var ix = {};
  loadDresses_().forEach(function (d) { ix[d.name.toLowerCase()] = d.brand; });
  return ix;
}

/** The house that owns the dress on this rental row. */
function brandOfRental_(r, ix) {
  return (ix || brandIndex_())[String(r.dress || '').toLowerCase()] || CONFIG.DEFAULT_BRAND;
}

/** May this caller see / touch this rental at all? */
function ownsRental_(ctx, r, ix) {
  if (ctx.role !== 'partner') return true;              // staff and public are not brand-limited
  return brandOfRental_(r, ix) === ctx.brand;
}

function rentalOut_(r, ctx) {
  var role = ctx && ctx.role ? ctx.role : ctx;   // tolerates an old-style string
  if (role === 'partner' && isCls_(r.co)) {   // dates only — no client, no money
    return { rentalId: r.id, dress: r.dress, co: 'CLS booking', renter: 'CLS client', status: r.status, mine: false,
             pickup: r.start !== null ? dayToIso_(r.start) : null, returnBy: r.ret !== null ? dayToIso_(r.ret) : null,
             availableAgain: r.blockEnd !== null ? dayToIso_(r.blockEnd + 1) : null };
  }
  return { rentalId: r.id, row: r.row, mine: true, renter: r.renter, dress: r.dress, co: r.co, status: r.status, payStatus: r.pay,
           eventDate: r.event !== null ? dayToIso_(r.event) : null, pickup: r.start !== null ? dayToIso_(r.start) : null,
           returnBy: r.ret !== null ? dayToIso_(r.ret) : null, availableAgain: r.blockEnd !== null ? dayToIso_(r.blockEnd + 1) : null,
           rate: r.rate, clientPrice: r.clientPrice,
           margin: r.clientPrice !== '' && r.rate !== '' ? Number(r.clientPrice) - Number(r.rate) : null,
           contact: r.contact, clsBookingId: r.clsId, notes: r.notes,
           holdUntil: r.holdUntil ? r.holdUntil.toISOString() : null, createdBy: r.createdBy };
}

function newId_() {
  return 'HR-' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyMMdd') + '-' +
         Math.random().toString(36).slice(2, 6).toUpperCase();
}

/**
 * Package add-on price. A number in the Photographer Rate column wins (use it for a special price).
 * Otherwise it's worked out from the Regular Rate: Regular Rate − ADDON_LESS, but never below ADDON_MIN.
 * So 799 → 700, 999 → 900, 499 → 500. The Photographer Rate column is optional.
 */
function addOnRate_(photoCell, regularCell) {
  var own = numOrNull_(photoCell);
  if (own !== null) return own;
  var reg = numOrNull_(regularCell);
  return reg === null ? null : Math.max(CONFIG.ADDON_MIN, reg - CONFIG.ADDON_LESS);
}

function numOrNull_(v) {                         // accepts 799, "799", "₱799", "P1,000"
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[₱Pp,\s]/g, '');
  return s === '' || isNaN(Number(s)) ? null : Number(s);
}

// Dates are handled as "day numbers" (days since 1970-01-01) in Manila time, so there are no timezone surprises.
function dayNum_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (v instanceof Date) return isoToDay_(Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd'));
  return isoToDay_(String(v));
}
function isoToDay_(s) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(s).trim());
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 864e5 : null;
}
function requireDay_(s, field) {
  var d = isoToDay_(s || '');
  if (d === null) throw new Error(field + ' must be a date like 2026-10-10');
  return d;
}
function dayToIso_(n) { return new Date(n * 864e5).toISOString().slice(0, 10); }
function dayToDate_(n) { var p = dayToIso_(n).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); } // script TZ = Manila
function todayDay_() { return isoToDay_(Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd')); }
// ───────────────────────────── RENTAL HOUSES ─────────────────────────────
/**
 * ▶▶ RUN THIS to add or change a house's PIN ▶▶  setMyPartnerPin()
 *
 * Put the house's name and the PIN you want between the quotes, save, pick
 * setMyPartnerPin from the dropdown and press Run. The Run button can't pass
 * arguments, which is why the values live here rather than in setPartnerPin().
 *
 * The name must match the Brand column on the Dresses tab exactly.
 */
function setMyPartnerPin() {
  return setPartnerPin('Serra', 'CHANGE-THIS');
}

function setPartnerPin(brand, pin) {
  if (brand === undefined || pin === undefined) {
    var u = 'Nothing happened — this needs a house and a PIN, and the Run button ' +
            'cannot type them. Use setMyPartnerPin() just above instead.';
    console.log(u); return { ok: false, error: u };
  }
  brand = String(brand || '').trim();
  pin = String(pin || '').trim();
  if (!brand) { console.log('Give the house a name.'); return { ok: false, error: 'No house name.' }; }
  if (/CHANGE-THIS/.test(pin)) {
    var m = 'The PIN still says CHANGE-THIS. Edit setMyPartnerPin() and run it again.';
    console.log(m); return { ok: false, error: m };
  }

  var props = PropertiesService.getScriptProperties(), pins = partnerPins_();
  if (!pin) {                                  // blank = remove this house's access
    delete pins[brand];
    props.setProperty('PARTNER_PINS', JSON.stringify(pins));
    if (props.getProperty('PARTNER_PIN') && brand === CONFIG.DEFAULT_BRAND) props.deleteProperty('PARTNER_PIN');
    console.log(brand + ' can no longer sign in.');
    return { ok: true, brand: brand, removed: true };
  }
  if (pin.length < 6) {
    var s = 'Use at least 6 characters — a 4-digit PIN is ten thousand guesses.';
    console.log(s); return { ok: false, error: s };
  }
  if (pin === props.getProperty('STAFF_PIN')) {
    var c = 'That is the CLS staff PIN. A house must not be able to sign in as staff.';
    console.log(c); return { ok: false, error: c };
  }
  var clash = Object.keys(pins).filter(function (b) { return b !== brand && String(pins[b]) === pin; })[0];
  if (clash) {
    var d = 'That PIN already belongs to ' + clash + '. Two houses sharing a PIN would share a view.';
    console.log(d); return { ok: false, error: d };
  }

  pins[brand] = pin;
  props.setProperty('PARTNER_PINS', JSON.stringify(pins));
  // Retire the old single-partner property so there are not two live answers.
  if (props.getProperty('PARTNER_PIN') && brand === CONFIG.DEFAULT_BRAND) props.deleteProperty('PARTNER_PIN');

  var known = {};
  loadDresses_().forEach(function (x) { known[x.brand] = (known[x.brand] || 0) + 1; });
  var cal = calendar_(brand);                  // the house's own calendar, ready to share
  console.log('PIN set for ' + brand + '.\n' +
              (known[brand] ? known[brand] + ' dress(es) on the Dresses tab belong to them.'
                            : '⚠ No dress has "' + brand + '" in its Brand column yet, so they will ' +
                              'sign in and see nothing. Check the spelling matches exactly.') +
              '\nTheir calendar: "' + cal.getName() + '" — share it with ' + brand + ' only.');
  return { ok: true, brand: brand, dresses: known[brand] || 0 };
}

/** Which houses can sign in, and how many dresses each has. PINs are not printed. */
function listPartners() {
  var pins = partnerPins_(), counts = {};
  loadDresses_().forEach(function (d) { counts[d.brand] = (counts[d.brand] || 0) + 1; });
  var lines = ['Houses that can sign in:'];
  Object.keys(pins).forEach(function (b) {
    lines.push('  • ' + b + ' — ' + (counts[b] || 0) + ' dress(es)' +
               (counts[b] ? '' : '  ⚠ no dresses carry this name'));
  });
  if (!Object.keys(pins).length) lines.push('  (none)');
  var orphan = Object.keys(counts).filter(function (b) { return !pins[b]; });
  if (orphan.length) {
    lines.push('Houses with dresses but no PIN (they cannot sign in):');
    orphan.forEach(function (b) { lines.push('  • ' + b + ' — ' + counts[b] + ' dress(es)'); });
  }
  var out = lines.join('\n');
  console.log(out);
  return out;
}
