/**
 * แบบก่อสร้างง่าย 3D — Backend (Google Apps Script)  v6
 * วางโค้ดนี้ใน Google Sheet > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web app
 * (Execute as: Me, Who has access: Anyone)
 *
 * รหัสองค์กร (org key): แก้ค่า ORG_KEY ด้านล่าง หรือใส่ใน Script Properties ชื่อ ORG_KEY
 */
var ORG_KEY = 'เปลี่ยนรหัสนี้';          // <-- เปลี่ยนเป็นรหัสลับขององค์กร
var FOLDER_NAME = 'drafting-app-projects'; // โฟลเดอร์ใน Google Drive ที่เก็บไฟล์โปรเจค

function key_() {
  return PropertiesService.getScriptProperties().getProperty('ORG_KEY') || ORG_KEY;
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, header) {
  var s = ss_().getSheetByName(name);
  if (!s) { s = ss_().insertSheet(name); s.appendRow(header); s.setFrozenRows(1); }
  return s;
}
function folder_() {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(FOLDER_NAME); p.setProperty('FOLDER_ID', f.getId()); return f;
}
var EV_H = ['time', 'date', 'userId', 'name', 'version', 'event', 'device', 'detail'];
var PR_H = ['id', 'name', 'ownerName', 'ownerId', 'year', 'createdAt', 'updatedAt', 'rev', 'lastEditor', 'editNote', 'parts', 'fileId', 'thumb'];
var RV_H = ['time', 'id', 'name', 'rev', 'editor', 'note'];

function out_(obj, cb) {
  var s = JSON.stringify(obj);
  if (cb && /^[\w.$]+$/.test(cb))
    return ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!p.action) return out_({ ok: true, app: 'drafting-app backend', v: 6 }, p.callback);
  return out_(handle_(p), p.callback);
}
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { return out_({ ok: false, error: 'bad json' }); }
  return out_(handle_(body));
}

function handle_(b) {
  if (String(b.key || '') !== String(key_())) return { ok: false, error: 'รหัสองค์กรไม่ถูกต้อง' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    switch (b.action) {
      case 'ping': return { ok: true };
      case 'event': return logEvents_(b);
      case 'saveProject': return saveProject_(b);
      case 'listProjects': return listProjects_(b);
      case 'getProject': return getProject_(b);
      default: return { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

// ---------- สถิติการใช้งาน (ไม่มีเนื้อหาโปรเจค) ----------
function logEvents_(b) {
  var list = b.events || [b], s = sheet_('Events', EV_H), rows = [];
  list.slice(0, 200).forEach(function (ev) {
    var t = ev.t ? new Date(ev.t) : new Date();
    rows.push([t, Utilities.formatDate(t, 'Asia/Bangkok', 'yyyy-MM-dd'), String(ev.uid || b.uid || '').slice(0, 40),
      String(ev.name || b.name || '').slice(0, 60), String(ev.ver || b.ver || '').slice(0, 30),
      String(ev.ev || '').slice(0, 40), String(ev.dev || b.dev || '').slice(0, 20), String(ev.d || '').slice(0, 200)]);
  });
  if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, EV_H.length).setValues(rows);
  ensureDashboard_();
  return { ok: true, n: rows.length };
}

// ---------- คลังโปรเจคกลาง ----------
function findRow_(s, id) {
  var n = s.getLastRow(); if (n < 2) return -1;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}
function saveProject_(b) {
  var p = b.project || {}; if (!p.id || !p.data) return { ok: false, error: 'missing project' };
  var s = sheet_('Projects', PR_H), r = findRow_(s, p.id), now = new Date();
  var json = JSON.stringify(p.data), old = r > 0 ? s.getRange(r, 1, 1, PR_H.length).getValues()[0] : null;
  var fileId = old ? old[11] : '', rev = old ? (Number(old[7]) || 1) + 1 : 1, file = null;
  if (fileId) { try { file = DriveApp.getFileById(fileId); file.setContent(json); } catch (e) { file = null; } }
  if (!file) { file = folder_().createFile(p.id + '.json', json, 'application/json'); fileId = file.getId(); }
  var thumb = String(p.thumb || ''); if (thumb.length > 45000) thumb = '';
  var editor = String(b.name || ''), note = String(p.editNote || '');
  var row = [p.id, String(p.name || ''), old ? old[2] : String(p.ownerName || editor), old ? old[3] : String(b.uid || ''),
    Number(p.year) || now.getFullYear(), old ? old[5] : new Date(p.created || now), now, rev, editor, note,
    (p.data.parts || []).length, fileId, thumb];
  if (r > 0) s.getRange(r, 1, 1, PR_H.length).setValues([row]); else s.appendRow(row);
  if (note || !old || old[8] !== editor) sheet_('Revisions', RV_H).appendRow([now, p.id, row[1], rev, editor, note || (old ? 'แก้ไข' : 'สร้างใหม่')]);
  return { ok: true, id: p.id, rev: rev, updatedAt: now.getTime() };
}
function listProjects_(b) {
  var s = sheet_('Projects', PR_H), n = s.getLastRow(); if (n < 2) return { ok: true, projects: [] };
  var v = s.getRange(2, 1, n - 1, PR_H.length).getValues(), res = [];
  v.forEach(function (x) {
    if (!x[0]) return;
    if (b.year && Number(x[4]) !== Number(b.year)) return;
    if (b.owner && String(x[2]).indexOf(b.owner) < 0) return;
    res.push({ id: x[0], name: x[1], ownerName: x[2], year: x[4], createdAt: +new Date(x[5]), updatedAt: +new Date(x[6]),
      rev: x[7], lastEditor: x[8], editNote: x[9], parts: x[10], thumb: b.noThumb ? '' : x[12] });
  });
  res.sort(function (a, c) { return c.updatedAt - a.updatedAt; });
  return { ok: true, projects: res };
}
function getProject_(b) {
  var s = sheet_('Projects', PR_H), r = findRow_(s, b.id); if (r < 0) return { ok: false, error: 'not found' };
  var x = s.getRange(r, 1, 1, PR_H.length).getValues()[0];
  var data = JSON.parse(DriveApp.getFileById(x[11]).getBlob().getDataAsString());
  return { ok: true, project: { id: x[0], name: x[1], ownerName: x[2], year: x[4], createdAt: +new Date(x[5]), updatedAt: +new Date(x[6]), rev: x[7], lastEditor: x[8], editNote: x[9], thumb: x[12], data: data } };
}

// ---------- แดชบอร์ด (สูตรอัตโนมัติ) ----------
function ensureDashboard_() {
  if (ss_().getSheetByName('Dashboard')) return;
  setupDashboard();
}
function setupDashboard() {
  sheet_('Events', EV_H); sheet_('Projects', PR_H);
  var d = ss_().getSheetByName('Dashboard') || ss_().insertSheet('Dashboard');
  d.clear();
  d.getRange('A1').setValue('📊 สรุปการใช้งาน').setFontSize(16).setFontWeight('bold');
  d.getRange('A3').setValue('ผู้ใช้ทั้งหมด'); d.getRange('B3').setFormula('=IFERROR(COUNTUNIQUE(Events!C2:C),0)');
  d.getRange('A4').setValue('จำนวนโปรเจคในคลังกลาง'); d.getRange('B4').setFormula('=COUNTA(Projects!A2:A)');
  d.getRange('A5').setValue('เหตุการณ์ทั้งหมด'); d.getRange('B5').setFormula('=COUNTA(Events!A2:A)');
  d.getRange('A7').setValue('ผู้ใช้ต่อวัน (30 วันล่าสุด)').setFontWeight('bold');
  d.getRange('A8').setFormula('=IFERROR(QUERY(Events!B2:C,"select B, count(C) where B is not null group by B order by B desc limit 30 label B \'วันที่\', count(C) \'จำนวนครั้ง\'",0),"")');
  d.getRange('D7').setValue('ผู้ใช้ไม่ซ้ำต่อวัน').setFontWeight('bold');
  d.getRange('D8').setFormula('=IFERROR(SORT(UNIQUE(FILTER(Events!B2:B,Events!B2:B<>"")),1,FALSE),"")');
  d.getRange('E8').setFormula('=IFERROR(ARRAYFORMULA(IF(D8:D="","",COUNTUNIQUEIFS(Events!C2:C,Events!B2:B,D8:D))),"")');
  d.getRange('G7').setValue('เวอร์ชันที่ใช้อยู่').setFontWeight('bold');
  d.getRange('G8').setFormula('=IFERROR(QUERY(Events!C2:E,"select E, count(C) where E is not null group by E order by count(C) desc label E \'เวอร์ชัน\', count(C) \'ครั้ง\'",0),"")');
  d.getRange('J7').setValue('เหตุการณ์ยอดนิยม').setFontWeight('bold');
  d.getRange('J8').setFormula('=IFERROR(QUERY(Events!F2:F,"select F, count(F) where F is not null group by F order by count(F) desc label F \'เหตุการณ์\', count(F) \'ครั้ง\'",0),"")');
  d.getRange('M7').setValue('ผู้ใช้ (ชื่อ)').setFontWeight('bold');
  d.getRange('M8').setFormula('=IFERROR(QUERY(Events!D2:D,"select D, count(D) where D <> \'\' group by D order by count(D) desc label D \'ชื่อ\', count(D) \'ครั้ง\'",0),"")');
  return 'ok';
}
