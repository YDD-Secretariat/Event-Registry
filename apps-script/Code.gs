/**
 * EVENT REGISTRY — backend (Google Apps Script, bound to the database Google Sheet)
 *
 * One script serves every programme. Each request carries a programId; every row
 * written carries that ProgramID. Programme settings (name, groups, labels, which
 * modules are on) are read live from the "Programs" sheet.
 *
 * Deploy: Extensions ▸ Apps Script ▸ paste this file ▸ Deploy ▸ New deployment ▸
 * Web app ▸ Execute as: Me ▸ Who has access: Anyone ▸ copy the /exec URL into config.js.
 */

var DUPLICATE_THRESHOLD = 80;       // % name similarity that triggers the duplicate-school warning
var IDEMPOTENCY_SECONDS = 21600;    // remember confirmed writes for 6h so retries never double-count

var SHEETS = {
  Programs: ['ProgramID','Name','Tagline','Status','Groups','OtherGroup','VisitorsInMembers',
    'ModSchools','ModMembers','ModVisitors','ModWorkersBulk','ModIndividuals','ModUploads',
    'MemberLabel','WorkerLabel','SchoolLabel','StudentLabel','GroupLabel','StationHint','Accent'],
  WorkerRoster: ['ProgramID','RosterID','Full Name','Station','Phone Number','Details'],
  RegisteredPeople: ['ProgramID','RegistrationID','RosterRowId','FullName','Station','Phone','OtherDetails','RegisteredAt','SubmittedBy'],
  Schools: ['ProgramID','SchoolID','SchoolName','NormalizedName','Location','CoordinatorName','CoordinatorPhone','CoordinatorLocation','Male','Female','Total','Station','Timestamp','PhotoFolderUrl','PhotoCount'],
  BulkRegistrations: ['ProgramID','EntryID','Group','Category','Count','RegisteredAt','SubmittedBy','Male','Female','Station','RequestId'],
  WorkerBulk: ['ProgramID','EntryID','Male','Female','Total','RegisteredAt','SubmittedBy','Station','RequestId'],
  VisitorRegistrations: ['ProgramID','EntryID','Category','Male','Female','Total','Notes','Station','RegisteredAt','RequestId'],
  ActivityLog: ['ProgramID','Timestamp','Station','Action','Details']
};

/* ------------------------------------------------------------------ routing */

function doGet(e) { return route_((e && e.parameter) || {}); }

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ success: false, error: 'Bad request body.' }); }
  return route_(body);
}

var READ_ACTIONS = {
  getPrograms: getPrograms_, getProgram: getProgram_,
  getStats: getStats_, getSchools: getSchools_, getRecent: getRecent_,
  getBulkStats: getBulkStats_, getRecentBulk: getRecentBulk_,
  getWorkerStats: getWorkerStats_, getRecentWorkerBulk: getRecentWorkerBulk_,
  getVisitorStats: getVisitorStats_, getRecentVisitors: getRecentVisitors_,
  getKitchenSummary: getKitchenSummary_, searchWorkers: searchWorkers_
};
var WRITE_ACTIONS = {
  registerSchool: registerSchool_, registerBulkMembers: registerBulkMembers_,
  registerBulkWorkersMinisters: registerBulkWorkers_, registerVisitor: registerVisitor_,
  registerWorker: registerWorker_, registerWalkIn: registerWalkIn_, uploadPhoto: uploadPhoto_
};

function route_(req) {
  try {
    var action = String(req.action || '');
    if (action === 'getPrograms') return json_(getPrograms_());

    var isWrite = !!WRITE_ACTIONS[action];
    var fn = isWrite ? WRITE_ACTIONS[action] : READ_ACTIONS[action];
    if (!fn) return json_({ success: false, error: 'Unknown action: ' + action });

    var prog = loadProgram_(req.programId);
    if (!prog) return json_({ success: false, error: 'Unknown programme: ' + (req.programId || '(none)') });

    if (!isWrite) return json_(fn(req, prog));

    if (prog.status === 'Closed') return json_({ success: false, error: 'This programme is closed for registration.' });
    var cache = CacheService.getScriptCache();
    var key = req.clientRequestId ? 'req:' + prog.id + ':' + req.clientRequestId : null;
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (key) { var hit = cache.get(key); if (hit) return json_(JSON.parse(hit)); }
      var result = fn(req, prog);
      if (key && result && result.success) cache.put(key, JSON.stringify(result), IDEMPOTENCY_SECONDS);
      return json_(result);
    } finally { lock.releaseLock(); }
  } catch (err) {
    return json_({ success: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------ sheet helpers */

function ss_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Script is not bound to a spreadsheet. Set Script Property SHEET_ID.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet "' + name + '". Run Event Registry ▸ Set up / repair sheets.');
  return sh;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
}

// All rows of a sheet as objects keyed by header, with _row (1-based sheet row).
function rows_(name, programId) {
  var sh = sheet_(name), last = sh.getLastRow();
  if (last < 2) return [];
  var h = headers_(sh);
  var vals = sh.getRange(2, 1, last - 1, h.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _row: i + 2 };
    for (var c = 0; c < h.length; c++) if (h[c]) o[h[c]] = vals[i][c];
    if (programId !== undefined && String(o.ProgramID) !== String(programId)) continue;
    if (programId === undefined || o.ProgramID !== '') out.push(o);
  }
  return out;
}

function append_(name, obj) {
  var sh = sheet_(name), h = headers_(sh);
  sh.appendRow(h.map(function (k) { return obj[k] === undefined ? '' : obj[k]; }));
}

function log_(prog, station, action, details) {
  try { append_('ActivityLog', { ProgramID: prog.id, Timestamp: new Date(), Station: station || '', Action: action, Details: details }); } catch (e) {}
}

function num_(v) { var n = parseInt(v, 10); return isNaN(n) || n < 0 ? 0 : n; }
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function id_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); }
function byTimeDesc_(field) { return function (a, b) { return new Date(b[field] || 0) - new Date(a[field] || 0); }; }
function norm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); }

/* --------------------------------------------------------------- programmes */

function programFromRow_(r) {
  var groups = String(r.Groups || '').split(',').map(function (g) { return g.trim(); }).filter(String);
  var other = String(r.OtherGroup || '').trim();
  if (other && groups.indexOf(other) < 0) groups.push(other);
  var on = function (v) { return v === '' || v === undefined ? false : bool_(v); };
  return {
    id: String(r.ProgramID), name: String(r.Name || r.ProgramID), tagline: String(r.Tagline || ''),
    status: String(r.Status || 'Active'), groups: groups, otherGroup: other,
    visitorsInMembers: on(r.VisitorsInMembers),
    modules: { schools: on(r.ModSchools), members: on(r.ModMembers), visitors: on(r.ModVisitors),
               workersBulk: on(r.ModWorkersBulk), individuals: on(r.ModIndividuals), uploads: on(r.ModUploads) },
    labels: { member: String(r.MemberLabel || 'Members'), worker: String(r.WorkerLabel || 'Workers'),
              school: String(r.SchoolLabel || 'School'), student: String(r.StudentLabel || 'Students'),
              group: String(r.GroupLabel || 'Group'), stationHint: String(r.StationHint || 'e.g. Desk 3') },
    accent: String(r.Accent || '#4f46e5')
  };
}

function loadProgram_(programId) {
  if (!programId) return null;
  var all = rows_('Programs');
  for (var i = 0; i < all.length; i++) if (String(all[i].ProgramID) === String(programId)) return programFromRow_(all[i]);
  return null;
}

function getPrograms_() {
  return { programs: rows_('Programs').filter(function (r) { return r.ProgramID; }).map(function (r) {
    var p = programFromRow_(r); return { id: p.id, name: p.name, tagline: p.tagline, status: p.status, accent: p.accent };
  }) };
}
function getProgram_(req, prog) { return { program: prog }; }

/* ------------------------------------------------------------------ schools */

function levenshtein_(a, b) {
  if (a === b) return 0; if (!a.length) return b.length; if (!b.length) return a.length;
  var prev = [], j, i;
  for (j = 0; j <= b.length; j++) prev[j] = j;
  for (i = 1; i <= a.length; i++) {
    var diag = prev[0]; prev[0] = i;
    for (j = 1; j <= b.length; j++) {
      var tmp = prev[j], cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      prev[j] = Math.min(prev[j] + 1, prev[j - 1] + 1, diag + cost); diag = tmp;
    }
  }
  return prev[b.length];
}
function similarity_(a, b) {
  var na = norm_(a), nb = norm_(b);
  if (!na || !nb) return 0; if (na === nb) return 100;
  return Math.round((1 - levenshtein_(na, nb) / Math.max(na.length, nb.length)) * 100);
}

function registerSchool_(req, prog) {
  var name = String(req.schoolName || '').trim();
  var male = num_(req.male), female = num_(req.female);
  if (name.length < 2) return { success: false, error: 'Enter the ' + prog.labels.school.toLowerCase() + ' name.' };
  if (male + female <= 0) return { success: false, error: 'Enter at least one male or female headcount.' };

  if (!req.confirmedNotDuplicate && !req.fromQueue) {
    var matches = rows_('Schools', prog.id).map(function (s) {
      return { schoolName: s.SchoolName, similarity: similarity_(name, s.SchoolName), coordinatorName: s.CoordinatorName,
               coordinatorPhone: s.CoordinatorPhone, male: s.Male, female: s.Female, total: s.Total, station: s.Station };
    }).filter(function (m) { return m.similarity >= DUPLICATE_THRESHOLD; })
      .sort(function (a, b) { return b.similarity - a.similarity; }).slice(0, 5);
    if (matches.length) return { success: false, duplicate: true, needsConfirmation: true, matches: matches };
  }
  var schoolId = id_('SCH-');
  append_('Schools', { ProgramID: prog.id, SchoolID: schoolId, SchoolName: name, NormalizedName: norm_(name),
    Location: String(req.location || '').trim(), CoordinatorName: String(req.coordinatorName || '').trim(),
    CoordinatorPhone: String(req.coordinatorPhone || '').trim(), CoordinatorLocation: String(req.coordinatorLocation || '').trim(),
    Male: male, Female: female, Total: male + female, Station: String(req.station || ''), Timestamp: new Date(), PhotoFolderUrl: '', PhotoCount: 0 });
  log_(prog, req.station, 'REGISTER_SCHOOL', name + ' (' + (male + female) + ')');
  return { success: true, schoolId: schoolId, schoolName: name, total: male + female };
}

function getSchools_(req, prog) { return { schools: rows_('Schools', prog.id).map(stripRow_) }; }
function getRecent_(req, prog) {
  var limit = num_(req.limit) || 8;
  return { recent: rows_('Schools', prog.id).sort(byTimeDesc_('Timestamp')).slice(0, limit).map(stripRow_) };
}
function getStats_(req, prog) {
  var s = rows_('Schools', prog.id), t = { totalStudents: 0, totalSchools: s.length, totalMale: 0, totalFemale: 0 };
  s.forEach(function (r) { t.totalMale += num_(r.Male); t.totalFemale += num_(r.Female); });
  t.totalStudents = t.totalMale + t.totalFemale;
  return t;
}
function stripRow_(r) { var o = {}; for (var k in r) if (k !== '_row' && k !== 'ProgramID') o[k] = r[k]; return o; }

/* ------------------------------------------------------------ members (bulk) */

function dupRequest_(sheetName, prog, requestId) {
  if (!requestId) return null;
  var rows = rows_(sheetName, prog.id);
  for (var i = rows.length - 1; i >= 0 && i >= rows.length - 300; i--) if (rows[i].RequestId === requestId) return rows[i];
  return null;
}

function registerBulkMembers_(req, prog) {
  var group = String(req.group || '').trim();
  if (prog.groups.indexOf(group) < 0) return { success: false, error: 'Choose a valid ' + prog.labels.group.toLowerCase() + '.' };
  var male = num_(req.male), female = num_(req.female);
  if (male + female <= 0) return { success: false, error: 'Enter at least one male or female count.' };
  if (dupRequest_('BulkRegistrations', prog, req.clientRequestId)) return { success: true, alreadyProcessed: true };
  append_('BulkRegistrations', { ProgramID: prog.id, EntryID: id_('BLK-'), Group: group, Category: 'Members', Count: male + female,
    RegisteredAt: new Date(), SubmittedBy: req.submittedBy || '', Male: male, Female: female, Station: req.station || '', RequestId: req.clientRequestId || '' });
  log_(prog, req.station, 'REGISTER_BULK', group + ' M' + male + ' F' + female);
  return { success: true, total: male + female };
}

function getRecentBulk_(req, prog) {
  var limit = num_(req.limit) || 8;
  return { recent: rows_('BulkRegistrations', prog.id).sort(byTimeDesc_('RegisteredAt')).slice(0, limit).map(stripRow_) };
}

// Per-group member totals. If the programme counts visitors inside Members, they are added to OtherGroup.
function memberBreakdown_(prog) {
  var map = {}, order = prog.groups.slice();
  order.forEach(function (g) { map[g] = { name: g, male: 0, female: 0, total: 0 }; });
  rows_('BulkRegistrations', prog.id).forEach(function (r) {
    var g = String(r.Group || '');
    if (!map[g]) { map[g] = { name: g, male: 0, female: 0, total: 0 }; order.push(g); }
    map[g].male += num_(r.Male); map[g].female += num_(r.Female);
  });
  if (prog.visitorsInMembers) {
    var vg = prog.otherGroup || 'Others';
    if (!map[vg]) { map[vg] = { name: vg, male: 0, female: 0, total: 0 }; order.push(vg); }
    rows_('VisitorRegistrations', prog.id).forEach(function (r) { map[vg].male += num_(r.Male); map[vg].female += num_(r.Female); });
  }
  var groups = order.map(function (g) { map[g].total = map[g].male + map[g].female; return map[g]; });
  var t = { male: 0, female: 0 };
  groups.forEach(function (g) { t.male += g.male; t.female += g.female; });
  return { groups: groups, male: t.male, female: t.female, total: t.male + t.female };
}

function getBulkStats_(req, prog) {
  var m = memberBreakdown_(prog);
  return { totalMembers: m.total, totalMale: m.male, totalFemale: m.female,
    groups: m.groups.map(function (g) { return { group: g.name, male: g.male, female: g.female, count: g.total }; }) };
}

/* ----------------------------------------------------------------- visitors */

function registerVisitor_(req, prog) {
  var male = num_(req.male), female = num_(req.female);
  if (male + female <= 0) return { success: false, error: 'Enter at least one male or female visitor.' };
  if (dupRequest_('VisitorRegistrations', prog, req.clientRequestId)) return { success: true, alreadyProcessed: true };
  append_('VisitorRegistrations', { ProgramID: prog.id, EntryID: id_('VIS-'), Category: req.category || 'Visitors', Male: male, Female: female,
    Total: male + female, Notes: String(req.notes || '').trim(), Station: req.station || '', RegisteredAt: new Date(), RequestId: req.clientRequestId || '' });
  log_(prog, req.station, 'REGISTER_VISITOR', 'M' + male + ' F' + female);
  return { success: true, total: male + female };
}

function getRecentVisitors_(req, prog) {
  var limit = num_(req.limit) || 8;
  return { recent: rows_('VisitorRegistrations', prog.id).sort(byTimeDesc_('RegisteredAt')).slice(0, limit).map(stripRow_) };
}
function getVisitorStats_(req, prog) {
  var v = rows_('VisitorRegistrations', prog.id).sort(byTimeDesc_('RegisteredAt')), t = { totalVisitors: 0, totalMale: 0, totalFemale: 0 };
  v.forEach(function (r) { t.totalMale += num_(r.Male); t.totalFemale += num_(r.Female); });
  t.totalVisitors = t.totalMale + t.totalFemale;
  t.visitors = v.map(stripRow_);
  return t;
}

/* ------------------------------------------------- workers: bulk + individual */

function registerBulkWorkers_(req, prog) {
  if (!req.confirmedExcludesIndividuals) return { success: false, error: 'Confirm that this count excludes people registered individually.' };
  var male = num_(req.male), female = num_(req.female);
  if (male + female <= 0) return { success: false, error: 'Enter at least one male or female count.' };
  if (dupRequest_('WorkerBulk', prog, req.clientRequestId)) return { success: true, alreadyProcessed: true };
  append_('WorkerBulk', { ProgramID: prog.id, EntryID: id_('WBK-'), Male: male, Female: female, Total: male + female,
    RegisteredAt: new Date(), SubmittedBy: req.submittedBy || '', Station: req.station || '', RequestId: req.clientRequestId || '' });
  log_(prog, req.station, 'REGISTER_WORKER_BULK', 'M' + male + ' F' + female);
  return { success: true, total: male + female };
}

function getRecentWorkerBulk_(req, prog) {
  var limit = num_(req.limit) || 8;
  return { recent: rows_('WorkerBulk', prog.id).sort(byTimeDesc_('RegisteredAt')).slice(0, limit).map(function (r) {
    var o = stripRow_(r); o.Count = o.Total; return o; }) };
}

function workerSummary_(prog) {
  var bulk = rows_('WorkerBulk', prog.id), people = rows_('RegisteredPeople', prog.id);
  var w = { bulkMale: 0, bulkFemale: 0, bulkTotal: 0, individualTotal: people.length, total: 0 };
  bulk.forEach(function (r) { w.bulkMale += num_(r.Male); w.bulkFemale += num_(r.Female); });
  w.bulkTotal = w.bulkMale + w.bulkFemale; w.total = w.bulkTotal + w.individualTotal;
  return w;
}

function getWorkerStats_(req, prog) {
  var w = workerSummary_(prog), by = {};
  rows_('RegisteredPeople', prog.id).forEach(function (p) { var s = String(p.Station || 'Unspecified'); by[s] = (by[s] || 0) + 1; });
  return { totalRegisteredWorkers: w.total, workersMinisters: w,
    byStation: Object.keys(by).sort().map(function (s) { return { station: s, count: by[s] }; }) };
}

function ensureRosterIds_() {
  var sh = sheet_('WorkerRoster'), h = headers_(sh), col = h.indexOf('RosterID') + 1, last = sh.getLastRow();
  if (!col || last < 2) return;
  var ids = sh.getRange(2, col, last - 1, 1).getValues(), changed = false;
  for (var i = 0; i < ids.length; i++) if (!ids[i][0]) { ids[i][0] = id_('RS-'); changed = true; }
  if (changed) sh.getRange(2, col, ids.length, 1).setValues(ids);
}

function searchWorkers_(req, prog) {
  var q = norm_(req.q);
  if (q.length < 2) return { workers: [] };
  var roster = rows_('WorkerRoster', prog.id);
  if (roster.some(function (r) { return !r.RosterID; })) {
    var lock = LockService.getScriptLock();
    if (lock.tryLock(8000)) { try { ensureRosterIds_(); } finally { lock.releaseLock(); } roster = rows_('WorkerRoster', prog.id); }
  }
  var done = {};
  rows_('RegisteredPeople', prog.id).forEach(function (p) { done[String(p.RosterRowId)] = true; });
  var tokens = q.split(' ');
  var hits = roster.filter(function (r) {
    var n = norm_(r['Full Name']); return n && tokens.every(function (t) { return n.indexOf(t) >= 0; });
  }).sort(function (a, b) {
    var sa = norm_(a['Full Name']).indexOf(q) === 0 ? 0 : 1, sb = norm_(b['Full Name']).indexOf(q) === 0 ? 0 : 1;
    return sa - sb || String(a['Full Name']).localeCompare(String(b['Full Name']));
  }).slice(0, 20);
  var core = { ProgramID: 1, RosterID: 1, 'Full Name': 1, Station: 1, 'Phone Number': 1, _row: 1 };
  return { workers: hits.map(function (r) {
    var other = {};
    for (var k in r) if (!core[k] && r[k] !== '' && r[k] !== null) other[k] = r[k];
    return { rowId: String(r.RosterID), fullName: String(r['Full Name']), station: String(r.Station || ''),
      phone: String(r['Phone Number'] || ''), otherDetails: other, alreadyRegistered: !!done[String(r.RosterID)] };
  }) };
}

function registerWorker_(req, prog) {
  var group = String(req.group || '').trim();
  if (prog.groups.indexOf(group) < 0) return { success: false, error: 'Choose a valid ' + prog.labels.group.toLowerCase() + '.' };
  var phone = String(req.phone || '').replace(/\D/g, '');
  if (phone.length < 10) return { success: false, error: 'Enter a valid phone number.' };
  var row = rows_('WorkerRoster', prog.id).filter(function (r) { return String(r.RosterID) === String(req.rowId); })[0];
  if (!row) return { success: false, error: 'That person is not on this programme’s roster.' };
  var already = rows_('RegisteredPeople', prog.id).filter(function (p) { return String(p.RosterRowId) === String(req.rowId); })[0];
  if (already) return { success: false, duplicate: true, fullName: row['Full Name'], error: row['Full Name'] + ' was already registered.' };
  var details = { Details: row.Details || '' };
  if (row.Station) details.RosterStation = row.Station;
  append_('RegisteredPeople', { ProgramID: prog.id, RegistrationID: id_('REG-'), RosterRowId: String(req.rowId), FullName: row['Full Name'],
    Station: group, Phone: phone, OtherDetails: JSON.stringify(details), RegisteredAt: new Date(), SubmittedBy: req.submittedBy || '' });
  log_(prog, req.submittedBy, 'REGISTER_INDIVIDUAL', row['Full Name'] + ' (' + req.rowId + ')');
  return { success: true, fullName: row['Full Name'] };
}

function registerWalkIn_(req, prog) {
  var name = String(req.name || '').trim(), group = String(req.group || '').trim();
  var phone = String(req.phone || '').replace(/\D/g, '');
  if (name.length < 2) return { success: false, error: 'Enter a full name.' };
  if (phone.length < 10) return { success: false, error: 'Enter a valid phone number.' };
  if (prog.groups.indexOf(group) < 0) return { success: false, error: 'Choose a valid ' + prog.labels.group.toLowerCase() + '.' };
  var dup = rows_('RegisteredPeople', prog.id).filter(function (p) { return norm_(p.FullName) === norm_(name) && String(p.Phone) === phone; })[0];
  if (dup) return { success: false, duplicate: true, fullName: name, error: name + ' was already registered.' };
  var wid = id_('WALKIN-');
  append_('RegisteredPeople', { ProgramID: prog.id, RegistrationID: id_('REG-'), RosterRowId: wid, FullName: name, Station: group, Phone: phone,
    OtherDetails: JSON.stringify({ Details: 'Walk-in' }), RegisteredAt: new Date(), SubmittedBy: req.submittedBy || '' });
  log_(prog, req.submittedBy, 'REGISTER_WALKIN', name + ' (' + wid + ')');
  return { success: true, fullName: name };
}

/* ------------------------------------------------------------------ uploads */

function folderIn_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function uploadPhoto_(req, prog) {
  var sheetRow = rows_('Schools', prog.id).filter(function (s) {
    return String(s.SchoolID) === String(req.schoolId) || String(s.SchoolName) === String(req.schoolName); })[0];
  if (!sheetRow) return { success: false, error: 'Register the ' + prog.labels.school.toLowerCase() + ' before uploading its sheets.' };
  var data = String(req.base64Data || '');
  var comma = data.indexOf(','); if (data.indexOf('data:') === 0 && comma > 0) data = data.slice(comma + 1);
  if (!data) return { success: false, error: 'No image data received.' };

  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('UPLOAD_ROOT_FOLDER_ID'), root;
  if (rootId) root = DriveApp.getFolderById(rootId);
  else { root = DriveApp.createFolder('Event Registry Uploads'); props.setProperty('UPLOAD_ROOT_FOLDER_ID', root.getId()); }
  var folder = folderIn_(folderIn_(root, prog.name), String(sheetRow.SchoolName).replace(/[\\\/:*?"<>|]/g, ' ').slice(0, 80));
  var blob = Utilities.newBlob(Utilities.base64Decode(data), req.mimeType || 'image/jpeg', req.fileName || ('sheet-' + Date.now() + '.jpg'));
  var file = folder.createFile(blob);

  var sh = sheet_('Schools'), h = headers_(sh);
  var cnt = num_(sheetRow.PhotoCount) + 1;
  sh.getRange(sheetRow._row, h.indexOf('PhotoCount') + 1).setValue(cnt);
  sh.getRange(sheetRow._row, h.indexOf('PhotoFolderUrl') + 1).setValue(folder.getUrl());
  log_(prog, req.station, 'UPLOAD_PHOTO', sheetRow.SchoolName + ' #' + cnt);
  return { success: true, fileId: file.getId(), photoCount: cnt };
}

/* ------------------------------------------------------------ live summary */

function getKitchenSummary_(req, prog) {
  var st = getStats_(req, prog), m = memberBreakdown_(prog), w = workerSummary_(prog), v = getVisitorStats_(req, prog);
  var visitorsExtra = prog.visitorsInMembers ? 0 : v.totalVisitors;
  var visitorsExtraM = prog.visitorsInMembers ? 0 : v.totalMale, visitorsExtraF = prog.visitorsInMembers ? 0 : v.totalFemale;
  return {
    grandTotal: st.totalStudents + m.total + w.total + visitorsExtra,
    studentTotal: st.totalStudents, memberTotal: m.total, groupTotal: m.total,
    visitorTotal: v.totalVisitors, workerTotal: w.total,
    totalMale: st.totalMale + m.male + w.bulkMale + visitorsExtraM,
    totalFemale: st.totalFemale + m.female + w.bulkFemale + visitorsExtraF,
    groups: m.groups,
    schools: rows_('Schools', prog.id).sort(function (a, b) { return String(a.SchoolName).localeCompare(String(b.SchoolName)); })
      .map(function (s) { return { name: s.SchoolName, total: num_(s.Total) }; }),
    workers: w
  };
}

/* ------------------------------------------------------- admin menu + setup */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Event Registry')
    .addItem('Set up / repair sheets', 'setup')
    .addItem('Assign roster IDs', 'assignRosterIds')
    .addToUi();
}

// Creates any missing sheet and any missing column header. Never deletes data.
function setup() {
  var ss = ss_(), created = [];
  Object.keys(SHEETS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); created.push(name); }
    var want = SHEETS[name], have = sh.getLastColumn() ? headers_(sh) : [];
    want.forEach(function (h) {
      if (have.indexOf(h) < 0) { sh.getRange(1, have.length + 1).setValue(h).setFontWeight('bold'); have.push(h); }
    });
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.getActive().toast(created.length ? 'Created: ' + created.join(', ') : 'All sheets are in place.', 'Event Registry');
}

function assignRosterIds() { ensureRosterIds_(); SpreadsheetApp.getActive().toast('Roster IDs assigned.', 'Event Registry'); }
