/**
 * SHIFT REPORT - Google Apps Script backend
 *
 * Lives inside your Google Sheet (Extensions > Apps Script).
 * Run setupSheets() once, then Deploy > New deployment > Web app.
 *
 * Sheets used:
 *   Users      Name | PIN | Active          (one row per supervisor)
 *   Equipment  Equipment | Area | Supervisor (which supervisor looks after which machine)
 *   Jobs       one row per job (written by the app, do not type here)
 */

var SHEETS = { JOBS: 'Jobs', USERS: 'Users', EQUIP: 'Equipment' };

var JOB_COLS = [
  'JobID', 'SubmittedAt', 'Date', 'Shift', 'Supervisor', 'Type', 'Equipment',
  'Description', 'StartTime', 'EndTime', 'DowntimeHrs', 'Cause', 'ActionTaken',
  'SpareParts', 'Status', 'CompletedDate', 'LastUpdated',
  'MachineHours', 'Mechanics'   // added later: always keep new columns at the END
];
var C = {};
JOB_COLS.forEach(function (c, i) { C[c] = i; });

var TYPES = ['Breakdown', 'Planned maintenance', 'Unplanned maintenance', 'Concern'];
var STATUSES = ['Completed', 'Ongoing', 'Waiting for parts'];
var SHIFTS = ['Morning', 'Afternoon', 'Night'];
var HISTORY_DAYS = 60;      // completed jobs older than this are not sent to the app
var MAX_PER_REQUEST = 50;
var MAX_TEXT = 2000;

// ---------------------------------------------------------------- setup

/** Run this ONCE from the editor. Safe to run again (it never deletes data). */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jobs = ensureSheet_(ss, SHEETS.JOBS, JOB_COLS);
  var users = ensureSheet_(ss, SHEETS.USERS, ['Name', 'PIN', 'Active']);
  var equip = ensureSheet_(ss, SHEETS.EQUIP, ['Equipment', 'Area', 'Supervisor']);

  fmtCol_(jobs, C.Date + 1, 'yyyy-mm-dd');
  fmtCol_(jobs, C.CompletedDate + 1, 'yyyy-mm-dd');
  fmtCol_(jobs, C.SubmittedAt + 1, 'yyyy-mm-dd hh:mm');
  fmtCol_(jobs, C.LastUpdated + 1, 'yyyy-mm-dd hh:mm');
  fmtCol_(jobs, C.StartTime + 1, '@');
  fmtCol_(jobs, C.EndTime + 1, '@');
  fmtCol_(users, 2, '@'); // keep PINs like 0123 intact

  if (users.getLastRow() < 2) {
    users.getRange(2, 1, 2, 3).setValues([
      ['Supervisor One', '1111', 'Yes'],
      ['Supervisor Two', '2222', 'Yes']
    ]);
  }
  if (equip.getLastRow() < 2) {
    equip.getRange(2, 1, 3, 3).setValues([
      ['Pump 1', 'Utilities', 'Supervisor One'],
      ['Conveyor A', 'Packing', 'Supervisor One'],
      ['Compressor 2', 'Utilities', 'Supervisor Two']
    ]);
  }
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  SpreadsheetApp.flush();
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else {
    // Sheet already exists: add any new headers at the end (never touches existing data)
    var have = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (String(have[i]).trim() === '') {
        sh.getRange(1, i + 1).setValue(headers[i]).setFontWeight('bold');
      }
    }
  }
  return sh;
}

function fmtCol_(sh, col, fmt) {
  var rows = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, col, rows, 1).setNumberFormat(fmt);
}

// ---------------------------------------------------------------- web app

function doGet() {
  return json_({ ok: true, service: 'Shift Report API' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents);
    lock.waitLock(25000);
    try {
      return json_(handle_(req));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle_(req) {
  var auth = authenticate_(req.name, req.pin);
  if (auth.error) return { ok: false, code: 'AUTH', error: auth.error };
  var user = auth.user;

  var results = { jobs: [], updates: [] };
  if (req.jobs && req.jobs.length) results.jobs = saveJobs_(user, req.jobs);
  if (req.updates && req.updates.length) results.updates = applyUpdates_(user, req.updates);

  return {
    ok: true,
    user: { name: user.name },
    machines: machinesFor_(user),
    data: { jobs: buildJobs_(user) },
    results: results
  };
}

// ---------------------------------------------------------------- helpers

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '". Run setupSheets() once.');
  return sh;
}
function tz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
function norm_(s) { return String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' '); }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }

function readRows_(sh, ncols) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ncols).getValues();
}

function fmtDate_(v, tz) {
  if (isDate_(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}
function fmtTime_(v, tz) {
  if (isDate_(v)) return Utilities.formatDate(v, tz, 'HH:mm');
  return String(v == null ? '' : v);
}

/** Stops text from being read as a spreadsheet formula. */
function clean_(v) {
  var s = String(v == null ? '' : v).trim();
  if (s.length > MAX_TEXT) s = s.slice(0, MAX_TEXT);
  if (/^[=+\-@]/.test(s)) s = ' ' + s;
  return s;
}

// ---------------------------------------------------------------- auth

function authenticate_(name, pin) {
  var cache = CacheService.getScriptCache();
  var key = 'fail_' + norm_(name).replace(/[^a-z0-9]/g, '_').slice(0, 60);
  var fails = Number(cache.get(key) || 0);
  if (fails >= 5) return { error: 'Too many wrong tries. Wait 10 minutes and try again.' };

  var rows = readRows_(sheet_(SHEETS.USERS), 3);
  for (var i = 0; i < rows.length; i++) {
    var rn = String(rows[i][0]).trim();
    if (!rn) continue;
    if (norm_(rn) === norm_(name) &&
        String(rows[i][1]).trim() === String(pin == null ? '' : pin).trim() &&
        norm_(rows[i][2]) !== 'no') {
      cache.remove(key);
      return { user: { name: rn } };
    }
  }
  cache.put(key, String(fails + 1), 600);
  return { error: 'Wrong name or PIN.' };
}

function machinesFor_(user) {
  var me = norm_(user.name);
  var out = [];
  readRows_(sheet_(SHEETS.EQUIP), 3).forEach(function (r) {
    var name = String(r[0]).trim();
    if (!name) return;
    var owners = String(r[2]).split(/[,;]/).map(norm_);
    if (owners.indexOf(me) >= 0) out.push({ name: name, area: String(r[1]).trim() });
  });
  return out;
}

// ---------------------------------------------------------------- jobs

function validateJob_(j, tz) {
  if (!j || !j.id) return 'Missing job id';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(j.date || ''))) return 'Bad date';
  if (SHIFTS.indexOf(j.shift) < 0) return 'Bad shift';
  if (TYPES.indexOf(j.type) < 0) return 'Bad type';
  if (STATUSES.indexOf(j.status) < 0) return 'Bad status';
  if (!String(j.equipment || '').trim()) return 'Equipment is required';
  if (!String(j.description || '').trim()) return 'Description is required';
  if (j.start && !/^\d{2}:\d{2}$/.test(j.start)) return 'Bad start time';
  if (j.end && !/^\d{2}:\d{2}$/.test(j.end)) return 'Bad end time';
  if (j.downtime !== '' && j.downtime != null && !(Number(j.downtime) >= 0)) return 'Bad downtime';
  if (j.machineHours !== '' && j.machineHours != null && !(Number(j.machineHours) >= 0)) return 'Bad machine hours';
  return null;
}

function saveJobs_(user, jobs) {
  var sh = sheet_(SHEETS.JOBS);
  var tz = tz_();
  var last = sh.getLastRow();
  var seen = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = true; });
  }
  var now = new Date();
  var rows = [];
  var results = [];

  jobs.slice(0, MAX_PER_REQUEST).forEach(function (j) {
    var bad = validateJob_(j, tz);
    if (bad) { results.push({ id: j && j.id, ok: false, error: bad }); return; }
    var id = String(j.id);
    if (seen[id]) { results.push({ id: id, ok: true, duplicate: true }); return; }
    seen[id] = true;

    var date = Utilities.parseDate(j.date, tz, 'yyyy-MM-dd');
    var completed = j.status === 'Completed' ? date : '';
    var down = (j.downtime === '' || j.downtime == null) ? '' : Number(j.downtime);
    var row = new Array(JOB_COLS.length);
    row[C.JobID] = id;
    row[C.SubmittedAt] = now;
    row[C.Date] = date;
    row[C.Shift] = j.shift;
    row[C.Supervisor] = user.name;            // always the logged-in person
    row[C.Type] = j.type;
    row[C.Equipment] = clean_(j.equipment);
    row[C.Description] = clean_(j.description);
    row[C.StartTime] = j.start || '';
    row[C.EndTime] = j.end || '';
    row[C.DowntimeHrs] = down;
    row[C.Cause] = clean_(j.cause);
    row[C.ActionTaken] = clean_(j.action);
    row[C.SpareParts] = clean_(j.parts);
    row[C.Status] = j.status;
    row[C.CompletedDate] = completed;
    row[C.LastUpdated] = now;
    row[C.MachineHours] = (j.machineHours === '' || j.machineHours == null) ? '' : Number(j.machineHours);
    row[C.Mechanics] = clean_(j.mechanics);
    rows.push(row);
    results.push({ id: id, ok: true });
  });

  if (rows.length) sh.getRange(last + 1, 1, rows.length, JOB_COLS.length).setValues(rows);
  return results;
}

function applyUpdates_(user, updates) {
  var sh = sheet_(SHEETS.JOBS);
  var tz = tz_();
  var last = sh.getLastRow();
  var data = last > 1 ? sh.getRange(2, 1, last - 1, JOB_COLS.length).getValues() : [];
  var mine = {};
  machinesFor_(user).forEach(function (m) { mine[norm_(m.name)] = true; });
  var me = norm_(user.name);
  var now = new Date();
  var results = [];

  updates.slice(0, MAX_PER_REQUEST).forEach(function (u) {
    var out = { qid: u.qid, ok: false };
    results.push(out);
    if (STATUSES.indexOf(u.status) < 0) { out.error = 'Bad status'; return; }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(u.date || ''))) { out.error = 'Bad date'; return; }

    var idx = -1;
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][C.JobID]) === String(u.jobId)) { idx = i; break; }
    }
    if (idx < 0) { out.error = 'Job not found'; return; }
    var row = data[idx];
    if (!mine[norm_(row[C.Equipment])] && norm_(row[C.Supervisor]) !== me) {
      out.error = 'Not your machine'; return;
    }

    var r = idx + 2;
    var note = clean_(u.note);
    if (note) {
      var entry = u.date + ': ' + note;
      var existing = String(row[C.ActionTaken] || '');
      if (existing.indexOf(entry) < 0) {
        var merged = existing ? existing + ' | ' + entry : entry;
        row[C.ActionTaken] = merged;
        sh.getRange(r, C.ActionTaken + 1).setValue(merged);
      }
    }
    row[C.Status] = u.status;
    sh.getRange(r, C.Status + 1).setValue(u.status);
    var done = u.status === 'Completed' ? Utilities.parseDate(u.date, tz, 'yyyy-MM-dd') : '';
    row[C.CompletedDate] = done;
    sh.getRange(r, C.CompletedDate + 1).setValue(done);
    row[C.LastUpdated] = now;
    sh.getRange(r, C.LastUpdated + 1).setValue(now);
    out.ok = true;
  });
  return results;
}

/** Jobs this supervisor should see: their own reports + anything on their machines. */
function buildJobs_(user) {
  var tz = tz_();
  var me = norm_(user.name);
  var mine = {};
  machinesFor_(user).forEach(function (m) { mine[norm_(m.name)] = true; });
  var cutoff = Utilities.formatDate(new Date(Date.now() - HISTORY_DAYS * 86400000), tz, 'yyyy-MM-dd');

  var sh = sheet_(SHEETS.JOBS);
  var out = [];
  readRows_(sh, JOB_COLS.length).forEach(function (r) {
    if (!r[C.JobID]) return;
    var eq = String(r[C.Equipment]);
    if (!mine[norm_(eq)] && norm_(r[C.Supervisor]) !== me) return;
    var date = fmtDate_(r[C.Date], tz);
    var status = String(r[C.Status]);
    if (status === 'Completed' && date < cutoff) return;
    out.push({
      id: String(r[C.JobID]),
      date: date,
      shift: String(r[C.Shift]),
      supervisor: String(r[C.Supervisor]),
      type: String(r[C.Type]),
      equipment: eq,
      description: String(r[C.Description]),
      start: fmtTime_(r[C.StartTime], tz),
      end: fmtTime_(r[C.EndTime], tz),
      downtime: r[C.DowntimeHrs] === '' ? '' : Number(r[C.DowntimeHrs]),
      cause: String(r[C.Cause]),
      action: String(r[C.ActionTaken]),
      parts: String(r[C.SpareParts]),
      status: status,
      completed: r[C.CompletedDate] === '' ? '' : fmtDate_(r[C.CompletedDate], tz),
      machineHours: r[C.MachineHours] === '' ? '' : Number(r[C.MachineHours]),
      mechanics: String(r[C.Mechanics])
    });
  });
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return out;
}
