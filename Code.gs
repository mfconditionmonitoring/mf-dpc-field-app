/**
 * MF-DPC-132 Cloud Inbox – Google Apps Script web app
 * Runs inside the MF condition monitoring Google account (mf.condition.monitoring@gmail.com). Receives tower data and original photos from the field app,
 * files them in Google Drive, keeps a live Sheet of progress and defects, and e-mails a Sev 3 alert at once.
 * The laptop engine downloads the records through this web app with its engine key (no Google sign-in on the laptop);
 * older set-ups can still read the same Drive folder through Google Drive for desktop.
 *
 * Personal sign-in: every upload carries the person's ID and secret. The secret is checked against
 * config/users.json, which the processing engine writes from the team register (it holds only a checker
 * value, never the key). A person whose access is withdrawn in the engine is refused at once.
 *
 * Script properties (Project settings → Script properties):
 *   TEAM_KEY      only for the first link test before the team register exists (optional)
 *   ALERT_EMAILS  comma-separated addresses for Sev 3 alerts (E&I Engineer first)
 *   SHEET_ID      filled automatically on first run
 */
var ROOT_NAME = 'DPC Inspection Inbox';
var VERSION = '2.4.0';

/* Processing-engine laptops: ID -> SHA-256 of its engine key. The key itself is typed once into the engine and never stored here.
   Remove a line (and deploy a new version) to cut a laptop off at once. */
var ENGINES = {
  'ENG-HQ': '02569f4818f5a59ee1dfac1c950e1fe9784b69e875a6ef3a514190877ef0cee9',
  'ENG-DPC1': 'db4a2598e63bc7281dc95961c34ecc79b3634188d8272f04c5a404d0a04e0509'
};

function props_() { return PropertiesService.getScriptProperties(); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function folder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function root_() {
  var it = DriveApp.getRootFolder().getFoldersByName(ROOT_NAME);
  return it.hasNext() ? it.next() : DriveApp.getRootFolder().createFolder(ROOT_NAME);
}

function safe_(s) { return String(s || '').replace(/[^A-Za-z0-9_.\- ]/g, '_'); }

/** Write (or replace) a file of the given name in a folder. */
function putFile_(folder, name, blob) {
  var old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  return folder.createFile(blob.setName(name));
}

/* ---- tower status: who holds each record (status.json in the round folder) ---- */
function readStatus_(rf) {
  var f = rf.getFilesByName('status.json');
  if (f.hasNext()) { try { return JSON.parse(f.next().getBlob().getDataAsString()); } catch (e) {} }
  // first use in a round: build it from the packages already received
  var st = {}, subs = rf.getFolders();
  while (subs.hasNext()) {
    var sf = subs.next(), pf = sf.getFilesByName('package.json');
    if (!pf.hasNext()) continue;
    try { st[sf.getName()] = entry_(JSON.parse(pf.next().getBlob().getDataAsString()), null, null); } catch (e) {}
  }
  writeStatus_(rf, st);
  return st;
}
function writeStatus_(rf, st) { putFile_(rf, 'status.json', Utilities.newBlob(JSON.stringify(st), 'application/json', 'status.json')); }
function entry_(pkg, who, old, summary) {
  var by = who ? who.id : ((pkg.seal || {}).by || pkg.lastUser || pkg.createdBy || '');
  var users = users_() || [], nm = by;
  for (var i = 0; i < users.length; i++) if (users[i].id === by) nm = users[i].name;
  return { structure: pkg.structure, form: pkg.form, status: pkg.status || 'open', holder: by, holderName: who ? who.name : nm,
    since: old && old.holder === by && old.since ? old.since : new Date().toISOString(), updated: new Date().toISOString(),
    rev: pkg.rev || 0, pkgId: pkg.id || '', completedBy: pkg.completedBy || '', completed: pkg.completed || '', startedBy: pkg.createdBy || (old || {}).startedBy || '',
    defects: (pkg.defects || []).length, highest: summary && summary.highest !== undefined ? +summary.highest || 0 : highest_(pkg) };
}
function highest_(pkg) { var h = 0; (pkg.defects || []).forEach(function (d) { if (+d.sev > h) h = +d.sev; }); return h;
}
function heldError_(e) { return new Error('HELD BY ' + (e.holderName || e.holder) + ' (' + e.holder + ') – continuing on their phone since ' + String(e.since || '').replace('T', ' ').slice(0, 16)); }

function sheet_() {
  var p = props_(), id = p.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('DPC Inspection Live Status');
    DriveApp.getFileById(ss.getId()).moveTo(root_());
    p.setProperty('SHEET_ID', ss.getId());
  }
  var tabs = {
    'Inspections': ['Key', 'Round', 'Structure', 'Form', 'Date', 'Inspector', 'Team', 'Points checked', 'Y', 'N', 'NA', 'NV',
      'Highest Sev', 'Defects', 'Photos listed', 'Photos received', 'Open items', 'Status', 'Last update', 'Distance from tower m'],
    'Defects': ['Key', 'Round', 'Structure', 'Form', 'Point', 'Position', 'Component', 'Sev', 'Description', 'Location',
      'Close-up photo', 'Context photo', 'Date', 'Last update'],
    'Alerts': ['Time', 'Round', 'Structure', 'Point', 'Sev', 'Description', 'Sent to', 'Photo attached'],
    'Uploads': ['Time', 'Round', 'Folder', 'File', 'Size kB', 'Device', 'Person', 'Name'],
    'Sign-ins': ['Time', 'Person', 'Name', 'Device', 'Result'],
    'Files': ['Time', 'Round', 'Path', 'File ID', 'Size']
  };
  Object.keys(tabs).forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) {
      sh = ss.insertSheet(n);
      sh.getRange(1, 1, 1, tabs[n].length).setValues([tabs[n]]).setFontWeight('bold').setBackground('#0E5A73').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
  var s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  return ss;
}

/** Replace the row whose first column equals key, or append it. */
function upsert_(sh, key, row) {
  var last = sh.getLastRow();
  if (last > 1) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  sh.appendRow(row);
}

function hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function sha256_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)); }

/** Team register written by the engine: [{id, name, role, active, checker}] */
function users_() {
  var cfg = root_().getFoldersByName('config');
  if (!cfg.hasNext()) return null;
  var f = cfg.next().getFilesByName('users.json');
  return f.hasNext() ? JSON.parse(f.next().getBlob().getDataAsString()) : null;
}

/** Returns the person, or throws. Before the team register exists the TEAM_KEY allows a link test only. */
function person_(id, secret) {
  var list = users_();
  if (!list) {
    var k = props_().getProperty('TEAM_KEY');
    if (k && String(secret) === k) return { id: 'SETUP', name: 'Link test', role: 'setup', active: true };
    throw new Error('Team register not published yet – the E&I Engineer sends it from the engine');
  }
  var uid = String(id || '').toUpperCase(), p = null;
  for (var i = 0; i < list.length; i++) if (list[i].id === uid) p = list[i];
  if (!p) throw new Error('Unknown person ' + uid);
  if (sha256_(String(secret || '')) !== p.checker) throw new Error('Wrong personal key for ' + uid);
  if (!p.active) throw new Error('ACCESS WITHDRAWN for ' + uid + ' – see the E&I Engineer');
  return p;
}

/* ---------------------------------------------------------------- GET ---- */
function doGet(e) {
  try {
    var p = e.parameter || {};
    if (p.action === 'ping') return json_({ ok: true, version: VERSION, time: new Date().toISOString() });
    if (p.action === 'check') {
      var ul = users_() || [], uid = String(p.user || '').toUpperCase(), hit = null;
      for (var i = 0; i < ul.length; i++) if (ul[i].id === uid) hit = ul[i];
      return json_({ ok: true, exists: !!hit, active: hit ? hit.active !== false : false });
    }
    var who = person_(p.user, p.auth);
    if (p.action === 'config') {
      var cfg = root_().getFoldersByName('config');
      var out = { ok: true, version: VERSION, time: new Date().toISOString(), person: { id: who.id, name: who.name, role: who.role } };
      if (cfg.hasNext()) {
        var f = cfg.next().getFilesByName('config.json');
        if (f.hasNext()) out.config = JSON.parse(f.next().getBlob().getDataAsString());
      }
      return json_(out);
    }
    if (p.action === 'status') {
      var round = safe_(p.round);
      var rf = root_().getFoldersByName(round);
      var list = [];
      if (rf.hasNext()) {
        var subs = rf.next().getFolders();
        while (subs.hasNext()) list.push(subs.next().getName());
      }
      var st = {};
      var rf2 = root_().getFoldersByName(round);
      if (rf2.hasNext()) st = readStatus_(rf2.next());
      return json_({ ok: true, folders: list, status: st, time: new Date().toISOString() });
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* --------------------------------------------------------------- POST ---- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'engine') return json_(engine_(body));
    // Presence & calling: own auth, no record lock (heartbeats must not wait on uploads)
    if (['beat', 'callchief', 'callclear', 'call', 'presence'].indexOf(body.action) >= 0) return json_(presenceOp_(body));
    var who;
    try { who = person_(body.user, body.auth); }
    catch (x) { if (body.action === 'login') sheet_().getSheetByName('Sign-ins').appendRow([new Date(), body.user || '', '', body.device || '', String(x.message || x)]); throw x; }
    if (body.action === 'login') {
      sheet_().getSheetByName('Sign-ins').appendRow([new Date(), who.id, who.name, body.device || '', 'signed in']);
      return json_({ ok: true, time: new Date().toISOString(), person: { id: who.id, name: who.name, role: who.role, title: who.title || '' } });
    }
    if (who.role === 'setup') throw new Error('Link test only – sign in with a personal key to send records');
    body.who = who;
    lock.waitLock(30000);
    var round = safe_(body.round), folderName = safe_(body.folder);
    var rf = folder_(root_(), round);
    var tf = folder_(rf, folderName);
    body.rf = rf;
    if (body.action === 'photo') return json_(savePhoto_(body, round, folderName, tf));
    if (body.action === 'package') return json_(savePackage_(body, round, folderName, tf));
    if (body.action === 'takeover') return json_(takeover_(body, round, folderName, tf));
    if (body.action === 'handback') return json_(handback_(body, round, folderName, tf));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function savePhoto_(body, round, folderName, tf) {
  var name = safe_(body.name);
  if (!/^DPC132_[A-Z0-9-]+_[A-Z]{1,2}\d{1,2}(_[A-Za-z0-9-]+)?\.(jpg|jpeg|png)$/i.test(name)) throw new Error('Photo name not in PRO-401 format: ' + name);
  var st = readStatus_(body.rf), cur = st[folderName];
  if (cur && cur.holder && cur.holder !== body.who.id) {
    var ex = tf.getFoldersByName('photos');
    if (ex.hasNext() && ex.next().getFilesByName(name).hasNext()) throw heldError_(cur);
  }
  var bytes = Utilities.base64Decode(body.b64);
  var blob = Utilities.newBlob(bytes, body.mime || 'image/jpeg', name);
  var file = putFile_(folder_(tf, 'photos'), name, blob);
  fileLog_(round, folderName + '/photos/' + name, file);
  sheet_().getSheetByName('Uploads').appendRow([new Date(), round, folderName, name, Math.round(bytes.length / 1024), body.device || '', body.who.id, body.who.name]);
  return { ok: true, id: file.getId(), name: name, bytes: bytes.length, time: new Date().toISOString() };
}

function savePackage_(body, round, folderName, tf) {
  var pkg = body.pkg;
  if (!pkg || !pkg.structure || !pkg.form) throw new Error('Package incomplete');
  pkg.received = new Date().toISOString();
  if (!pkg.seal || pkg.seal.by !== body.who.id) throw new Error('Record not sealed by the person sending it');
  var st = readStatus_(body.rf), cur = st[folderName];
  if (cur && cur.holder && cur.holder !== body.who.id) throw heldError_(cur);
  var pfile = putFile_(tf, 'package.json', Utilities.newBlob(JSON.stringify(pkg, null, 1), 'application/json', 'package.json'));
  fileLog_(round, folderName + '/package.json', pfile);
  st[folderName] = entry_(pkg, body.who, cur, body.summary);
  writeStatus_(body.rf, st);

  var ss = sheet_();
  var s = body.summary || {};
  var key = round + '|' + pkg.structure + '|' + pkg.form;
  var received = 0, ph = tf.getFoldersByName('photos');
  if (ph.hasNext()) { var it = ph.next().getFiles(); while (it.hasNext()) { it.next(); received++; } }
  upsert_(ss.getSheetByName('Inspections'), key, [key, round, pkg.structure, pkg.form, (pkg.header || {}).date || '',
    (pkg.signoff || {}).inspector || '', (pkg.header || {}).team || '', s.checked || 0, s.Y || 0, s.N || 0, s.NA || 0, s.NV || 0,
    s.highest || 0, (pkg.defects || []).length, Object.keys(pkg.photos || {}).length, received, s.open || 0,
    pkg.status || '', new Date(), s.distance === undefined ? '' : s.distance]);

  var dsh = ss.getSheetByName('Defects');
  (pkg.defects || []).forEach(function (d) {
    var dk = key + '|' + d.key;
    var loc = [d.face ? 'Face ' + d.face : '', d.panel ? 'panel ' + d.panel : '', d.member ? 'member ' + d.member : ''].filter(String).join(', ');
    upsert_(dsh, dk, [dk, round, pkg.structure, pkg.form, d.point, d.pos || d.loc || '', d.component, d.sev, d.desc || '', loc,
      d.photo || '', d.photo2 || '', (pkg.header || {}).date || '', new Date()]);
  });

  var alerts = alertSev3_(pkg, round, tf);
  return { ok: true, key: key, photosReceived: received, alerts: alerts, time: new Date().toISOString() };
}

/** Another team member (or the same person on a new phone) continues a record: they become its holder. */
function takeover_(body, round, folderName, tf) {
  var st = readStatus_(body.rf), cur = st[folderName];
  var pf = tf.getFilesByName('package.json');
  if (!pf.hasNext()) throw new Error('Nothing received yet for ' + folderName);
  var pkg = JSON.parse(pf.next().getBlob().getDataAsString());
  if (!cur) cur = entry_(pkg, null, null);
  var who = body.who;
  if (cur.status === 'complete' && !(who.role === 'lead' || who.role === 'engineer')) throw new Error('Completed by ' + (cur.holderName || cur.holder) + ' – only a team leader or the E&I Engineer can take it over');
  var names = [], ph = tf.getFoldersByName('photos');
  if (ph.hasNext()) { var it = ph.next().getFiles(); while (it.hasNext()) names.push(it.next().getName()); }
  var prev = cur.holder;
  cur.holder = who.id; cur.holderName = who.name; cur.since = new Date().toISOString(); cur.updated = cur.since;
  cur.takenFrom = prev; cur.takeoverReason = body.reason || '';
  st[folderName] = cur; writeStatus_(body.rf, st);
  sheet_().getSheetByName('Uploads').appendRow([new Date(), round, folderName, 'TAKEOVER from ' + prev + (body.reason ? ' – ' + body.reason : ''), 0, body.device || '', who.id, who.name]);
  return { ok: true, pkg: pkg, photos: names, time: new Date().toISOString() };
}

/** After a review the team leader / E&I Engineer returns the record to the inspector. */
function handback_(body, round, folderName, tf) {
  var st = readStatus_(body.rf), cur = st[folderName];
  if (!cur || cur.holder !== body.who.id) throw new Error('Only the person holding the record can return it');
  var to = String(body.to || '').toUpperCase(), ul = users_() || [], u = null;
  for (var i = 0; i < ul.length; i++) if (ul[i].id === to) u = ul[i];
  if (!u) throw new Error('Unknown person ' + to);
  cur.holder = u.id; cur.holderName = u.name; cur.since = new Date().toISOString(); cur.updated = cur.since; cur.returnedBy = body.who.id;
  st[folderName] = cur; writeStatus_(body.rf, st);
  sheet_().getSheetByName('Uploads').appendRow([new Date(), round, folderName, 'RETURNED to ' + u.id + ' after review', 0, body.device || '', body.who.id, body.who.name]);
  return { ok: true, time: new Date().toISOString() };
}

/* ---------------------------------------------------------- ENGINE ---- */
/** The laptop engine: list changed files, download them, publish config/users, clear completed records. */
function engine_(b) {
  var h = ENGINES[String(b.engine || '')];
  if (!h || sha256_(String(b.ekey || '')) !== h) throw new Error('ENGINE NOT AUTHORISED (' + (b.engine || '?') + ') – contact the MF E&I Engineer');
  var root = root_();
  if (b.op === 'hello') return { ok: true, version: VERSION, engine: b.engine, time: new Date().toISOString() };
  if (b.op === 'changes') return changes_(root, b);
  if (b.op === 'get') {
    var out = [], total = 0;
    (b.ids || []).forEach(function (id) {
      var f = DriveApp.getFileById(id);
      if (!inTree_(root, f)) throw new Error('Not an inbox file');
      var bytes = f.getBlob().getBytes(); total += bytes.length;
      out.push({ id: id, b64: Utilities.base64Encode(bytes) });
    });
    return { ok: true, files: out, bytes: total };
  }
  if (b.op === 'put') {
    if (['config.json', 'users.json'].indexOf(b.name) < 0) throw new Error('Only config.json and users.json can be published');
    var parsed = JSON.parse(b.text);
    if (b.name === 'users.json' && (!parsed.length || !parsed.some(function (u) { return u.role === 'engineer' && u.active !== false; })))
      throw new Error('Team register refused: it must list at least one active E&I Engineer');
    putFile_(folder_(root, 'config'), b.name, Utilities.newBlob(b.text, 'application/json', b.name));
    return { ok: true, time: new Date().toISOString() };
  }
  if (b.op === 'trash') {
    var rf = root.getFoldersByName(safe_(b.round));
    if (!rf.hasNext()) return { ok: true, trashed: 0 };
    var r = rf.next(), n = 0;
    (b.folders || []).forEach(function (name) { var it = r.getFoldersByName(safe_(name)); while (it.hasNext()) { it.next().setTrashed(true); n++; } });
    return { ok: true, trashed: n };
  }
  throw new Error('Unknown engine operation');
}
function inTree_(root, f) {
  var rid = root.getId(), p = f.getParents();
  for (var d = 0; d < 4 && p.hasNext(); d++) { var x = p.next(); if (x.getId() === rid) return true; p = x.getParents(); }
  return false;
}
/** Every file the phones send is noted here, so the engine asks only for what is new. */
function fileLog_(round, path, file) {
  sheet_().getSheetByName('Files').appendRow([new Date(), round, path, file.getId(), file.getSize()]);
}
/** New files of a round. With no 'after' row: the whole round folder (resumable with a cursor), then the log from its end. */
function changes_(root, b) {
  var round = safe_(b.round), sh = sheet_().getSheetByName('Files'), last = sh.getLastRow(), list = [];
  if (b.after === undefined || b.after === null) {
    var rf = root.getFoldersByName(round), t0 = Date.now();
    if (!rf.hasNext()) return { ok: true, files: [], next: last };
    var recs = [], it = rf.next().getFolders();
    while (it.hasNext()) recs.push(it.next());
    recs.sort(function (x, y) { return x.getName() < y.getName() ? -1 : 1; });
    for (var i = +b.cursor || 0; i < recs.length; i++) {
      var r = recs[i], fs = r.getFiles();
      while (fs.hasNext()) { var f = fs.next(); list.push({ path: round + '/' + r.getName() + '/' + f.getName(), id: f.getId(), size: f.getSize() }); }
      var ph = r.getFoldersByName('photos');
      if (ph.hasNext()) { var ps = ph.next().getFiles(); while (ps.hasNext()) { var g = ps.next(); list.push({ path: round + '/' + r.getName() + '/photos/' + g.getName(), id: g.getId(), size: g.getSize() }); } }
      if (Date.now() - t0 > 200000 && i < recs.length - 1) return { ok: true, files: list, cursor: i + 1, next: b.next || last };
    }
    return { ok: true, files: list, next: b.next || last };
  }
  var from = Math.max(2, +b.after + 1);
  if (last >= from) {
    var v = sh.getRange(from, 1, last - from + 1, 5).getValues();
    for (var k = 0; k < v.length; k++) if (String(v[k][1]) === round) list.push({ path: round + '/' + v[k][2], id: v[k][3], size: v[k][4] });
  }
  return { ok: true, files: list, next: Math.max(last, +b.after) };
}

/** E-mail every Sev 3 defect once, with its close-up photo when it has already arrived. */
function alertSev3_(pkg, round, tf) {
  var p = props_(), to = p.getProperty('ALERT_EMAILS');
  if (!to) return 0;
  var sent = JSON.parse(p.getProperty('ALERTED') || '{}'), n = 0;
  (pkg.defects || []).forEach(function (d) {
    if (+d.sev !== 3) return;
    var k = round + '|' + pkg.structure + '|' + pkg.form + '|' + d.key;
    if (sent[k]) return;
    var attach = [], photoName = d.photo ? 'DPC132_' + pkg.structure + '_' + d.photo + '.jpg' : '';
    if (photoName) {
      var ph = tf.getFoldersByName('photos');
      if (ph.hasNext()) { var f = ph.next().getFilesByName(photoName); if (f.hasNext()) attach.push(f.next().getBlob()); }
    }
    var g = (pkg.photos || {})[d.photo] || {};
    var map = g.lat ? 'https://maps.google.com/?q=' + g.lat + ',' + g.lon : '';
    var loc = [d.pos || d.loc || '', d.face ? 'Face ' + d.face : '', d.member ? 'member ' + d.member : ''].filter(String).join(', ');
    var subject = 'SEV 3 – ' + pkg.structure + ' – ' + d.point + ' ' + d.component + ' – DPC 132 kV ' + round;
    var html = '<p><b>Sev 3 defect recorded in the field.</b> Inform DPC the same day (PRO-401 step 7).</p>' +
      '<table cellpadding="4" style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
      '<tr><td><b>Structure</b></td><td>' + pkg.structure + '</td></tr>' +
      '<tr><td><b>Point</b></td><td>' + d.point + ' – ' + d.component + '</td></tr>' +
      '<tr><td><b>Location</b></td><td>' + (loc || '–') + '</td></tr>' +
      '<tr><td><b>Description</b></td><td>' + (d.desc || '') + '</td></tr>' +
      '<tr><td><b>Inspector</b></td><td>' + ((pkg.signoff || {}).inspector || '') + ' – ' + ((pkg.header || {}).team || '') + '</td></tr>' +
      '<tr><td><b>Recorded</b></td><td>' + (g.time || pkg.updated || '') + '</td></tr>' +
      (map ? '<tr><td><b>Photo position</b></td><td><a href="' + map + '">' + g.lat + ', ' + g.lon + '</a></td></tr>' : '') +
      '</table><p>' + (attach.length ? 'Close-up photo attached: ' + photoName : 'Close-up photo not yet received; it will be in the Drive inbox when the phone uploads it.') + '</p>' +
      '<p style="color:#666">MF-DPC-132 Cloud Inbox ' + VERSION + '</p>';
    MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, attachments: attach });
    sheet_().getSheetByName('Alerts').appendRow([new Date(), round, pkg.structure, d.point + (d.pos ? ' ' + d.pos : ''), 3, d.desc || '', to, attach.length ? 'Yes' : 'No']);
    sent[k] = new Date().toISOString(); n++;
  });
  p.setProperty('ALERTED', JSON.stringify(sent));
  return n;
}

/* ---------------------------------------------- PRESENCE & CALLING ----
   Lightweight team safety layer. Phones send a quiet heartbeat (beat) with
   GPS + battery while the app is open; the Chief's board reads who is online
   (presence) and can ring a phone (call); a phone can call the Chief
   (callchief), which e-mails the Chief with the person's last-known position.
   State lives in Script Properties (tiny, ~6 people) so there is no Drive
   contention with record uploads, and these ops never take the record lock. */
var PRESENCE_TTL = 10 * 60 * 1000;   // treat as online if seen within 10 min
var CALL_TTL     = 2  * 60 * 1000;   // an incoming call rings for up to 2 min
var MEET_DEFAULT = 'https://meet.google.com/sbi-dasn-rvt';

function meetUrl_()   { return props_().getProperty('MEET_URL')   || MEET_DEFAULT; }
function chiefEmail_(){ return props_().getProperty('CHIEF_EMAIL') || 'shfqtali25@gmail.com'; }
function readMap_(k)  { try { return JSON.parse(props_().getProperty(k) || '{}'); } catch (e) { return {}; } }
function writeMap_(k, m) { props_().setProperty(k, JSON.stringify(m)); }

function engineOK_(b) { var h = ENGINES[String(b.engine || '')]; return !!h && sha256_(String(b.ekey || '')) === h; }

/** Chief-side ops (presence, call) accept an engine key OR an engineer/lead personal key. */
function chief_(b) {
  if (engineOK_(b)) return { id: b.engine, name: 'Engine ' + b.engine, role: 'engine' };
  var who = person_(b.user, b.auth);
  if (who.role !== 'engineer' && who.role !== 'lead') throw new Error('Only the E&I Engineer or a team leader can do this');
  return who;
}

function presenceOp_(b) {
  var now = Date.now();

  if (b.action === 'beat') {
    var who = person_(b.user, b.auth);
    var pm = readMap_('PRESENCE'), prev = pm[who.id] || {};
    pm[who.id] = { id: who.id, name: who.name, role: who.role,
      lat: b.lat, lon: b.lon, acc: b.acc, batt: b.batt, device: b.device || '',
      ts: now, calling: prev.calling || false, callingSince: prev.callingSince || 0 };
    writeMap_('PRESENCE', pm);
    var cm = readMap_('CALLS'), c = cm[who.id] || null;
    if (c && now - (c.ts || 0) > CALL_TTL) { delete cm[who.id]; writeMap_('CALLS', cm); c = null; }
    return { ok: true, time: new Date().toISOString(),
      call: c ? { from: c.from, fromName: c.fromName, url: c.url, ts: c.ts } : null };
  }

  if (b.action === 'callclear') {
    var w2 = person_(b.user, b.auth), cm2 = readMap_('CALLS');
    if (cm2[w2.id]) { delete cm2[w2.id]; writeMap_('CALLS', cm2); }
    var pm2 = readMap_('PRESENCE'); if (pm2[w2.id]) { pm2[w2.id].calling = false; writeMap_('PRESENCE', pm2); }
    return { ok: true };
  }

  if (b.action === 'callchief') {
    var w3 = person_(b.user, b.auth), pm3 = readMap_('PRESENCE');
    var e = pm3[w3.id] || { id: w3.id, name: w3.name, role: w3.role };
    e.name = w3.name; e.role = w3.role;
    if (b.lat !== undefined && b.lat !== null) { e.lat = b.lat; e.lon = b.lon; e.acc = b.acc; }
    if (b.batt !== undefined && b.batt !== null) e.batt = b.batt;
    e.device = b.device || e.device || ''; e.ts = now; e.calling = true; e.callingSince = now;
    pm3[w3.id] = e; writeMap_('PRESENCE', pm3);
    try { sheet_().getSheetByName('Sign-ins').appendRow([new Date(), w3.id, w3.name, b.device || '', 'CALLED CHIEF']); } catch (x) {}
    try { emailCall_(w3, e); } catch (x) {}
    return { ok: true, url: meetUrl_(), time: new Date().toISOString() };
  }

  if (b.action === 'presence') {
    chief_(b);
    var pm4 = readMap_('PRESENCE'), online = [], calling = [], DAY = 24 * 60 * 60 * 1000;
    Object.keys(pm4).forEach(function (id) {
      var e = pm4[id]; if (now - (e.ts || 0) > DAY) return;   // keep last-known so a man-down stays visible; board flags STALE
      e.seconds = Math.round((now - e.ts) / 1000);
      e.stale = e.seconds > 600;                               // no ping in 10 min: app closed, phone off, or signal loss
      online.push(e); if (e.calling) calling.push(e);
    });
    online.sort(function (a, c) { return a.ts < c.ts ? 1 : -1; });
    return { ok: true, time: new Date().toISOString(), meet: meetUrl_(), online: online, calling: calling };
  }

  if (b.action === 'call') {
    var me = chief_(b), to = String(b.to || '').toUpperCase();
    var ul = users_() || [], u = null;
    for (var i = 0; i < ul.length; i++) if (ul[i].id === to) u = ul[i];
    if (!u) throw new Error('Unknown person ' + to);
    var cm5 = readMap_('CALLS');
    cm5[to] = { from: me.id, fromName: me.name, url: meetUrl_(), ts: now };
    writeMap_('CALLS', cm5);
    return { ok: true, to: to, name: u.name, url: meetUrl_(), time: new Date().toISOString() };
  }

  throw new Error('Unknown presence action');
}

function emailCall_(who, e) {
  var to = chiefEmail_(); if (!to) return;
  var map = (e.lat !== undefined && e.lat !== null) ? 'https://maps.google.com/?q=' + e.lat + ',' + e.lon : '';
  var subject = 'CALL CHIEF – ' + who.name + ' (' + who.id + ') – DPC 132 kV field';
  var html = '<p><b>' + who.name + ' (' + who.id + ') is calling the Chief from the field.</b></p>' +
    '<table cellpadding="4" style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
    '<tr><td><b>Person</b></td><td>' + who.name + ' (' + who.id + ')</td></tr>' +
    (map ? '<tr><td><b>Last-known location</b></td><td><a href="' + map + '">' + e.lat + ', ' + e.lon + '</a>' + (e.acc ? ' (±' + Math.round(e.acc) + ' m)' : '') + '</td></tr>'
         : '<tr><td><b>Location</b></td><td>not available yet</td></tr>') +
    (e.batt !== undefined && e.batt !== null ? '<tr><td><b>Battery</b></td><td>' + e.batt + '%</td></tr>' : '') +
    '<tr><td><b>Time</b></td><td>' + new Date().toISOString().replace('T', ' ').slice(0, 19) + ' UTC</td></tr>' +
    '</table><p><a href="' + meetUrl_() + '" style="font-size:15px">Join the call line</a></p>' +
    '<p style="color:#666">MF-DPC-132 Cloud Inbox ' + VERSION + '</p>';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
}

/** Run once from the editor: creates the folder and the live Sheet, and asks for permissions. */
function setup() {
  var r = root_();
  folder_(r, 'config');
  var ss = sheet_();
  Logger.log('Inbox folder: ' + r.getUrl());
  Logger.log('Live status sheet: ' + ss.getUrl());
}
