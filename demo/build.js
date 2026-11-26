// Demo build (run from the repo root): node demo/build.js
// Wraps Index.html with an in-page copy of Code.gs and fake Google + Zoom services (state in localStorage).
const fs = require('fs');
const path = require('path');
const [, , dir = 'apps-script', out = 'demo/index.html'] = process.argv;
const code = [fs.readFileSync(path.join(dir, 'Images.gs'), 'utf8'), fs.readFileSync(path.join(dir, 'Code.gs'), 'utf8')].join('\n').replace(/^const /gm, 'var ');
const page = fs.readFileSync(path.join(dir, 'Index.html'), 'utf8');

const mock = `<meta name="viewport" content="width=device-width, initial-scale=1"><script>
const DB = JSON.parse(localStorage.getItem('demoDB') || '{}');
const saveDB = () => localStorage.setItem('demoDB', JSON.stringify(DB));
function sheetApi(name) {
  const rows = DB[name];
  return {
    getRange: (a, c, nr, nc) => typeof a === 'string' ? { setNumberFormat() {} } : { setValues: v => { rows[a - 1] = v[0].map(String); saveDB(); } },
    appendRow: r => { rows.push(r.map(String)); saveDB(); },
    setFrozenRows() {},
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    deleteRow: i => { rows.splice(i - 1, 1); saveDB(); },
  };
}
const ss = { getId: () => 'db1', getSheetByName: n => DB[n] ? sheetApi(n) : null, insertSheet: n => { DB[n] = []; saveDB(); return sheetApi(n); }, deleteSheet() {} };
const SpreadsheetApp = { openById: () => ss, create: () => ss };
const userProps = JSON.parse(localStorage.getItem('demoProps') || '{}');
const ZOOM = location.search.includes('nozoom') ? {} : { ZOOM_ACCOUNT_ID: 'acc', ZOOM_CLIENT_ID: 'id', ZOOM_CLIENT_SECRET: 'secret', ZOOM_HOST_EMAIL: 'rimon@example.com' };
const PropertiesService = {
  getScriptProperties: () => ({ getProperty: k => ZOOM[k] || null }),
  getUserProperties: () => ({ getProperty: k => userProps[k] ?? null, setProperty: (k, v) => { userProps[k] = v; localStorage.setItem('demoProps', JSON.stringify(userProps)); } }),
};
const CacheService = { getScriptCache: () => ({ get: k => sessionStorage.getItem(k), put: (k, v) => sessionStorage.setItem(k, v) }) };
const Utilities = { base64Encode: s => btoa(s), getUuid: () => crypto.randomUUID(), base64Decode: s => s.length, newBlob: (b, type, name) => ({ type, name, bytes: b }) };
const MailApp = { getRemainingDailyQuota: () => 100 };
const Session = { getEffectiveUser: () => ({ getEmail: () => 'teacher@example.com' }) };
window.sent = []; window.zoomCalls = [];
const GmailApp = { sendEmail: (...a) => sent.push(['send', ...a]), createDraft: (...a) => { if (location.search.includes('draftfail')) throw new Error('Service invoked too many times for one day: gmail'); sent.push(['draft', ...a]); return { getMessage: () => ({ getId: () => '18c4f2a9b7d3e1f0' }) }; } };
const res = (code, body) => ({ getResponseCode: () => code, getContentText: () => JSON.stringify(body) });
const UrlFetchApp = { fetch: (url, o) => {
  zoomCalls.push({ url, o });
  if (url.includes('/oauth/token')) return res(200, { access_token: 'tok', expires_in: 3600, scope: location.search.includes('noscope') ? 'cloud_recording:read:list_user_recordings:admin' : 'meeting:write:meeting:admin meeting:read:past_meeting:admin cloud_recording:read:list_user_recordings:admin' });
  if (url.includes('/meetings') && !url.includes('past_meetings') && o.method === 'post') { const b = JSON.parse(o.payload); return res(201, { id: 84512349876, password: 'r7Kq2x', join_url: 'https://us06web.zoom.us/j/84512349876?pwd=demo', topic: b.topic }); }
  if (url.includes('/past_meetings/')) return res(200, { duration: 97, participants_count: 29 });
  if (url.includes('/recordings')) return res(200, { meetings: [
    { topic: 'Reading: matching headings', start_time: '2026-11-24T14:00:00Z', duration: 58, share_url: 'https://us06web.zoom.us/rec/share/abc123', password: 'Rd5!x' },
    { topic: 'Listening Part 3 practice', start_time: '2026-11-21T14:05:00Z', duration: 72, share_url: 'https://us06web.zoom.us/rec/share/def456', recording_play_passcode: 'xyz' } ] });
  return res(404, { message: 'not mocked' });
} };
${code}
// Demo only: banners from this deployment, and sample batches on first visit.
BANNER_URL = location.origin + '/apps-script/assets/';
if (!localStorage.getItem('demoDB')) {
  const mails = (b, n) => Array.from({ length: n }, (_, i) => 'student' + String(i + 1).padStart(2, '0') + '.b' + b + '@example.com').join(', ');
  for (let b = 60; b <= 82; b++) saveBatch('', 'Batch ' + b, mails(b, 18 + (b % 13)));
  saveBatch('', 'Writing Saturday', mails('ws', 12));
  const day = d => new Date(Date.now() - d * 864e5).toISOString().slice(0, 10) + 'T20:00';
  [[12, 'live', 'Batch 82 · Class 1', 90, 95, 27], [9, 'live', 'Batch 82 · Class 2', 90, 88, 25], [8, 'recorded', 'Batch 82 · Recording 1', 60, 60, ''],
   [5, 'live', 'Batch 82 · Class 3', 90, 97, 29], [3, 'live', 'Batch 81 · Class 7', 90, 91, 22]].forEach(([d, type, title, plan, act, att]) =>
    writeRow_('Classes', 0, { ID: Utilities.getUuid(), Batch: title.split(' · ')[0], Type: type, Title: title, Start: day(d), 'Planned min': plan, 'Actual min': act, Attended: att,
      'Meeting ID': type === 'live' ? '84512349876' : '', Link: 'https://us06web.zoom.us/j/84512349876', 'Sent at': day(d + 1).replace('T', ' '), Students: 28 }));
}
function resetDemo() { localStorage.removeItem('demoDB'); localStorage.removeItem('demoProps'); location.reload(); }
function runner(ok, fail) { return new Proxy({}, { get: (t, p) => p === 'withSuccessHandler' ? f => runner(f, fail) : p === 'withFailureHandler' ? f => runner(ok, f) :
  (...a) => setTimeout(() => { let r; try { r = JSON.parse(JSON.stringify(window[p](...a) ?? null)); } catch (e) { return fail && fail(e); } ok && ok(r); }, 150) }); }
window.google = { script: { get run() { return runner(); } } };
<\/script>`;
// Make page-side globals not collide with server names inside the same window.
const strip = '<div style="font:600 13px/1.4 system-ui,sans-serif;background:#fff4d6;color:#7a5200;padding:8px 16px;text-align:center">Demo with sample data. Nothing is sent and no Zoom class is made. <a href="#" onclick="resetDemo();return false" style="color:inherit">Reset</a></div>';
fs.writeFileSync(out, page.replace('<head>', () => '<head>' + mock).replace(/<body[^>]*>/, m => m + strip));
console.log('wrote', out);
