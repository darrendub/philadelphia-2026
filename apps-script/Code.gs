/**
 * Philadelphia 2026 itinerary backend.
 * Paste this whole file into Extensions > Apps Script in the trip Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * The first request creates an "Itinerary" tab and fills it with the starting plan.
 */
const SHEET_ID = '1-HrgUqvTkK7HSaC0bQH7HybjUlG5OTRclz8QzqrtsUQ';
const TAB = 'Itinerary';
const HEADERS = ['id','day','time','endTime','area','title','place','address','notes','who',
  'flight','from','to','arrDay','arrTime','utc','arrUtc','updated'];
const NUMERIC = ['utc','arrUtc'];

function doGet(e) {
  try {
    return json_({ ok: true, items: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();
    if (body.action === 'add') {
      const item = body.item || {};
      if (!item.id) item.id = newId_();
      sh.appendRow(toRow_(item));
    } else if (body.action === 'update') {
      const row = findRow_(sh, body.id);
      if (!row) return json_({ ok: false, error: 'not_found' });
      const current = rowToObj_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0]);
      const merged = Object.assign({}, current, body.patch || {}, { id: body.id });
      sh.getRange(row, 1, 1, HEADERS.length).setValues([toRow_(merged)]);
    } else if (body.action === 'delete') {
      const row = findRow_(sh, body.id);
      if (row) sh.deleteRow(row);
    } else {
      return json_({ ok: false, error: 'unknown_action' });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function newId_() {
  return 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB);
  if (!sh) sh = ss.insertSheet(TAB);
  if (sh.getLastRow() === 0) {
    // Plain text everywhere so times like 16:55 and days like 2026-11-12 stay exactly as typed.
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    const rows = seed_().map(toRow_);
    sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }
  return sh;
}

function toRow_(item) {
  return HEADERS.map(function (h) {
    if (h === 'updated') return new Date().toISOString();
    let v = item[h];
    if (h === 'who') v = Array.isArray(v) ? v.join(',') : (v || '');
    if (v === null || v === undefined) v = '';
    return String(v);
  });
}

function rowToObj_(row) {
  const o = {};
  HEADERS.forEach(function (h, i) {
    let v = row[i];
    if (v instanceof Date) v = '';
    v = (v === null || v === undefined) ? '' : String(v).trim();
    if (h === 'who') o[h] = v ? v.split(',').map(function (s) { return s.trim(); }).filter(Boolean) : [];
    else if (NUMERIC.indexOf(h) >= 0) o[h] = v === '' ? null : Number(v);
    else o[h] = v;
  });
  return o;
}

function readAll_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, HEADERS.length);
  const values = range.getValues();
  let fixed = false;
  const out = [];
  values.forEach(function (row) {
    const o = rowToObj_(row);
    if (!o.title) return;
    if (!o.id) { o.id = newId_(); row[0] = o.id; fixed = true; }
    out.push(o);
  });
  if (fixed) range.setValues(values); // rows added by hand in Sheets get an id
  return out;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function seed_() {
  return [
    { id: 'f-as834', area: 'flight', flight: 'AS 834', title: 'Honolulu to San Francisco', from: 'HNL', to: 'SFO', day: '2026-11-12', time: '22:15', arrDay: '2026-11-13', arrTime: '05:25', utc: -10, arrUtc: -8, who: ['Maya', 'Emma'] },
    { id: 'f-aa2553', area: 'flight', flight: 'AA 2553', title: 'San Francisco to Philadelphia', from: 'SFO', to: 'PHL', day: '2026-11-13', time: '08:30', arrDay: '2026-11-13', arrTime: '16:54', utc: -8, arrUtc: -5, who: ['Maya', 'Emma'] },
    { id: 'f-as822', area: 'flight', flight: 'AS 822', title: 'Honolulu to Seattle', from: 'HNL', to: 'SEA', day: '2026-11-12', time: '22:25', arrDay: '2026-11-13', arrTime: '06:08', utc: -10, arrUtc: -8, who: ['Scott', 'Gail'] },
    { id: 'f-as442', area: 'flight', flight: 'AS 442', title: 'Seattle to Philadelphia', from: 'SEA', to: 'PHL', day: '2026-11-13', time: '08:32', arrDay: '2026-11-13', arrTime: '16:45', utc: -8, arrUtc: -5, who: ['Scott', 'Gail'] },
    { id: 'f-ua218', area: 'flight', flight: 'UA 218', title: 'Honolulu to Chicago', from: 'HNL', to: 'ORD', day: '2026-11-12', time: '16:55', arrDay: '2026-11-13', arrTime: '05:01', utc: -10, arrUtc: -6, who: ['Darren', 'Sophia', 'Rose'] },
    { id: 'f-ua423', area: 'flight', flight: 'UA 423', title: 'Chicago to Philadelphia', from: 'ORD', to: 'PHL', day: '2026-11-13', time: '06:42', arrDay: '2026-11-13', arrTime: '09:52', utc: -6, arrUtc: -5, who: ['Darren', 'Sophia', 'Rose'] },
    { id: 'cheesesteaks', area: 'food', title: 'Cheesesteak run', day: '2026-11-13', time: '12:00', notes: "With Mike Kuykendall\n\nOptions to choose from:\nPat's or Geno's (touristy)\nDalessandro's (Roxborough)\nJim's Steaks (cash only)\nJohn's Roast Pork\nSkinny Joe's\nAngelo's" },
    { id: 'airbnb-drinks', area: 'family', title: 'Drinks at Haworth Airbnb', day: '2026-11-13', time: '18:00' },
    { id: 'welcome-drinks', area: 'wedding', title: 'Aloha Welcome Drinks', day: '2026-11-14', time: '15:00', endTime: '17:00', place: 'Franklin Station Clubhouse', address: '806 Hoopes Lane, Media, PA 19063', notes: 'Attire: Hawaiian shirts and aloha gear.\nCelebratory cocktails and light apps provided. All are welcome.' },
    { id: 'ceremony', area: 'wedding', title: 'Ceremony', day: '2026-11-15', time: '16:30', endTime: '17:00', place: "Terrain at Styer's", address: '914 Baltimore Pike, Glen Mills, PA 19342', notes: 'Attire: semi-formal.\nThe ceremony is outside, so bring a warm jacket. No smoking on the property.\nAbout 20 minutes from the Airbnb.' },
    { id: 'reception', area: 'wedding', title: 'Reception', day: '2026-11-15', time: '17:00', endTime: '22:00', place: "Terrain at Styer's", address: '914 Baltimore Pike, Glen Mills, PA 19342' },
    { id: 'leave-airbnb', area: 'logistics', title: 'Leave the Airbnb for the airport', day: '2026-11-17', time: '04:00', place: 'Airbnb', address: '2 Yorktown Avenue, West Chester, PA 19382', notes: 'About 45 minutes to the airport.' },
    { id: 'f-as370', area: 'flight', flight: 'AS 370', title: 'Philadelphia to Seattle', from: 'PHL', to: 'SEA', day: '2026-11-17', time: '07:02', arrDay: '2026-11-17', arrTime: '10:08', utc: -5, arrUtc: -8, who: ['Darren', 'Sophia', 'Rose', 'Scott', 'Gail'] },
    { id: 'f-as626', area: 'flight', flight: 'AS 626', title: 'Seattle to Honolulu', from: 'SEA', to: 'HNL', day: '2026-11-17', time: '11:55', arrDay: '2026-11-17', arrTime: '16:30', utc: -8, arrUtc: -10, who: ['Darren', 'Sophia', 'Rose', 'Scott', 'Gail'] }
  ];
}
