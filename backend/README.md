# Backend (Google Apps Script) — drafting-app v6
- `Code.gs` — paste into Google Sheet > Extensions > Apps Script; set `ORG_KEY` (or Script Property `ORG_KEY`); Deploy as Web app (Execute as: Me, Access: Anyone).
- Actions (POST text/plain JSON `{action,key,...}`, or GET `?action=..&key=..&callback=fn` for JSONP): `ping`, `event` (`events:[{t,ev,d,uid,name,ver,dev}]`), `saveProject` (`project:{id,name,ownerName,year,created,thumb,editNote,data}`; upsert, rev+1, JSON in Drive folder `drafting-app-projects`), `listProjects` (`year`,`owner`,`noThumb`), `getProject` (`id`).
- Sheet tabs auto-created: `Events`, `Projects`, `Revisions`, `Dashboard` (run `setupDashboard()` to rebuild).
- `mock-server.js` — local emulator for tests: `node backend/mock-server.js` (API :8771 key `test-key`, app :8770).
- Thai setup guide: `SETUP-TH.md`.
