// Local mock of the Google Apps Script backend (for testing only).
// node backend/mock-server.js [apiPort=8771] [appPort=8770]
const http = require('http'), fs = require('fs'), path = require('path');
const KEY = process.env.ORG_KEY || 'test-key';
const API = +(process.argv[2] || 8771), APP = +(process.argv[3] || 8770);
const db = { events: [], projects: {}, revisions: [] };
let offline = false;
function send(res, obj, cb) {
  const s = JSON.stringify(obj);
  res.writeHead(200, { 'Content-Type': cb ? 'application/javascript' : 'application/json', 'Access-Control-Allow-Origin': '*' });
  res.end(cb ? `${cb}(${s})` : s);
}
function handle(b) {
  if (b.key !== KEY) return { ok: false, error: 'รหัสองค์กรไม่ถูกต้อง' };
  switch (b.action) {
    case 'ping': return { ok: true };
    case 'event': (b.events || [b]).forEach(e => db.events.push(e)); return { ok: true, n: (b.events || [b]).length };
    case 'saveProject': {
      const p = b.project, old = db.projects[p.id], rev = old ? old.rev + 1 : 1;
      db.projects[p.id] = { id: p.id, name: p.name, ownerName: old ? old.ownerName : (p.ownerName || b.name), year: p.year, createdAt: old ? old.createdAt : p.created, updatedAt: Date.now(), rev, lastEditor: b.name, editNote: p.editNote || '', parts: (p.data.parts || []).length, thumb: p.thumb, data: p.data };
      if (p.editNote || !old || old.lastEditor !== b.name) db.revisions.push({ id: p.id, rev, editor: b.name, note: p.editNote });
      return { ok: true, id: p.id, rev };
    }
    case 'listProjects': return { ok: true, projects: Object.values(db.projects).filter(p => (!b.year || p.year == b.year) && (!b.owner || (p.ownerName || '').includes(b.owner))).map(({ data, ...m }) => m).sort((a, c) => c.updatedAt - a.updatedAt) };
    case 'getProject': { const p = db.projects[b.id]; return p ? { ok: true, project: p } : { ok: false, error: 'not found' }; }
  }
  return { ok: false, error: 'unknown action' };
}
http.createServer((req, res) => {
  const u = new URL(req.url, 'http://x');
  if (u.pathname === '/__db') return send(res, db);
  if (u.pathname === '/__offline') { offline = u.searchParams.get('v') === '1'; return send(res, { offline }); }
  if (offline) { req.socket.destroy(); return; }
  if (req.method === 'OPTIONS') { res.writeHead(400); return res.end('preflight not supported (like Apps Script)'); }
  if (req.method === 'GET') { const p = Object.fromEntries(u.searchParams); return send(res, p.action ? handle(p) : { ok: true, mock: true }, p.callback); }
  let body = ''; req.on('data', c => body += c); req.on('end', () => { let b = {}; try { b = JSON.parse(body); } catch {} send(res, handle(b)); });
}).listen(API, () => console.log('mock API on', API));
http.createServer((req, res) => {
  const fp = path.join(__dirname, '..', 'index.html');
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); res.end(fs.readFileSync(fp));
}).listen(APP, () => console.log('app on', APP));
