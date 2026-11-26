// Rimons IELTS class mail: a private web page (Google Apps Script web app).
// Batches hold the student emails. Pick the batch, the kind of email, fill one or two things, Open in Gmail.
// Data lives in one Google Sheet the script makes in your Drive ("Rimons IELTS class mail data").

const SENDER_NAME = 'Rimons IELTS';
const TIMEZONE = 'Asia/Dhaka';
// Hosted copies of the banners, used only by "Copy email design" (Gmail drops pictures pasted as data).
// They load once apps-script/assets is on the repo's main branch.
const BANNER_URL = 'https://cdn.jsdelivr.net/gh/ihmorol/email-template-generator@main/apps-script/assets/';
const UTC_OFFSET_MIN = 360; // Bangladesh is UTC+6 all year
const TABS = {
  Batches: ['Name', 'Emails', 'Created'],
  // Every email sent goes in here; Type says which kind. Only live and recorded count as classes.
  Classes: ['ID', 'Batch', 'Type', 'Title', 'Start', 'Planned min', 'Actual min', 'Attended', 'Meeting ID', 'Link', 'Sent at', 'Students', 'Passcode', 'Note'],
};
// The kinds of email. banner = the picture in assets/ (packed in Images.gs).
const KINDS = {
  live: { label: 'Live class', dot: '#16A34A', banner: 'live', subject: 'Live class: ' },
  recorded: { label: 'Class recording', dot: '#B5660A', banner: 'recorded', subject: 'Class recording: ' },
  reminder: { label: 'Class reminder', dot: '#16A34A', banner: 'reminder', subject: 'Reminder: ' },
  test: { label: 'Mock test', dot: '#7C3AED', banner: 'test', subject: 'Mock test: ' },
  resources: { label: 'Class resources', dot: '#2563EB', banner: 'resources', subject: 'Resources: ' },
  noclass: { label: 'No class', dot: '#64748B', banner: 'noclass', subject: '' },
  message: { label: 'Notice', dot: '#0A1226', banner: 'notice', subject: '' },
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Rimons IELTS class mail')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---- Storage: one spreadsheet, made on first use ----

function db_() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('dbId');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Rimons IELTS class mail data');
    props.setProperty('dbId', ss.getId());
  }
  Object.keys(TABS).forEach(name => {
    const old = ss.getSheetByName(name);
    // A sheet made by an older version gets the new column names added to its header.
    if (old) { if (old.getLastColumn() < TABS[name].length) old.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]); return; }
    const sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // keep everything as plain text, no date guessing
    sh.appendRow(TABS[name]);
    sh.setFrozenRows(1);
  });
  const blank = ss.getSheetByName('Sheet1');
  if (blank) ss.deleteSheet(blank);
  return ss;
}

function rows_(tab) {
  const sh = db_().getSheetByName(tab);
  const head = TABS[tab];
  return sh.getDataRange().getValues().slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { o[h] = r[j] === undefined ? '' : String(r[j]); });
    return o;
  });
}

function writeRow_(tab, row, obj) {
  const sh = db_().getSheetByName(tab);
  const vals = [TABS[tab].map(h => (obj[h] === undefined || obj[h] === null ? '' : String(obj[h])))];
  if (row) sh.getRange(row, 1, 1, vals[0].length).setValues(vals);
  else sh.appendRow(vals[0]);
}

// ---- Emails: pulls addresses out of anything pasted ----

// Longer endings first so "com" wins over "co", "info" over "in", "dev" over "de".
const TLD_ = 'com|net|org|edu|gov|info|io|co|me|bd|uk|in|us|ac|biz|ca|au|dev|de|pk|my|sg|ae|xyz|app';
const EMAIL_OK_ = /^[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$/;

// "rahim@gmail.comkarim@yahoo.com" -> two addresses, cut after a known ending.
function splitGlued_(tok) {
  const parts = tok.split('@');
  if (parts.length <= 2) return [tok];
  const cut = new RegExp('^(.+?\\.(?:' + TLD_ + ')(?:\\.(?:' + TLD_ + '))?)(.+)$', 'i');
  const out = [];
  let local = parts[0];
  for (let i = 1; i < parts.length; i++) {
    if (i === parts.length - 1) { out.push(local + '@' + parts[i]); break; }
    const m = parts[i].match(cut);
    if (m) { out.push(local + '@' + m[1]); local = m[2]; } else { out.push(local + '@' + parts[i]); local = ''; }
  }
  return out;
}

// Works with commas, spaces, new lines, semicolons, tabs, "Name <a@b.com>", mailto:, or no separator at all.
function extractEmails_(text) {
  const t = String(text || '').replace(/mailto:/gi, ' ').replace(/[ \t]*@[ \t]*/g, '@');
  const ok = new Set();
  const bad = new Set();
  t.split(/[^A-Za-z0-9._%+@-]+/).filter(x => x.includes('@')).forEach(tok => {
    splitGlued_(tok).forEach(e => {
      e = e.toLowerCase().replace(/^[._-]+|[._-]+$/g, '');
      if (EMAIL_OK_.test(e) && !e.includes('..')) ok.add(e); else if (e) bad.add(e);
    });
  });
  return { emails: [...ok], skipped: [...bad] };
}

// Called while typing in the emails box, so nothing is a surprise on Save.
function checkEmails(text) {
  return extractEmails_(text);
}

// ---- Batches ----

function findBatch_(name) {
  return rows_('Batches').find(b => b.Name.toLowerCase() === String(name).trim().toLowerCase()) || null;
}

function emailsOf_(b) {
  return b ? b.Emails.split('\n').filter(Boolean) : [];
}

// oldName is empty for a new batch. Renaming also moves the batch's class history.
function saveBatch(oldName, name, text) {
  name = String(name || '').trim();
  if (!name) throw new Error('Give the batch a name.');
  const { emails } = extractEmails_(text);
  if (!emails.length) throw new Error('No email addresses found. Nothing was saved, your text is still in the box.');
  const same = findBatch_(name);
  const old = oldName ? findBatch_(oldName) : null;
  if (same && (!old || same._row !== old._row)) throw new Error('There is already a batch called ' + same.Name + '. Pick another name.');
  writeRow_('Batches', old ? old._row : 0, { Name: name, Emails: emails.join('\n'), Created: old ? old.Created : today_() });
  if (old && old.Name !== name) {
    rows_('Classes').filter(c => c.Batch === old.Name).forEach(c => { c.Batch = name; writeRow_('Classes', c._row, c); });
  }
  return { name, count: emails.length };
}

// The class history stays in the sheet.
function deleteBatch(name) {
  const b = findBatch_(name);
  if (b) db_().getSheetByName('Batches').deleteRow(b._row);
}

function recipients_(batch) {
  return emailsOf_(findBatch_(batch));
}

// ---- Classes and the tracker ----

// Only classes that already happened count; upcoming ones show in the list as "Upcoming".
function stats_(all) {
  const classes = all.filter(c => (c.Type === 'live' || c.Type === 'recorded') && (!c.Start || endedAt_(c)));
  const minutes = classes.reduce((s, c) => s + (Number(c['Actual min']) || Number(c['Planned min']) || 0), 0);
  return {
    classes: classes.length,
    live: classes.filter(c => c.Type === 'live').length,
    recorded: classes.filter(c => c.Type === 'recorded').length,
    minutes,
  };
}

// Everything the page needs on load.
function pageData() {
  const classes = rows_('Classes');
  // Most recently used first (last class sent, else when the batch was made); ties: Batch 82 before Batch 81.
  const batches = rows_('Batches').map(b => {
    const own = classes.filter(c => c.Batch === b.Name);
    const recent = own.reduce((m, c) => (c['Sent at'] > m ? c['Sent at'] : m), b.Created || '');
    // Next numbers for the automatic title ("Batch 82 · Class 14"); scheduled classes count too.
    const next = {};
    ['live', 'recorded', 'test', 'resources'].forEach(t => { next[t] = own.filter(c => c.Type === t).length + 1; });
    // Classes and tests not over yet, soonest first: what a reminder can be about.
    const upcoming = own.filter(c => (c.Type === 'live' || c.Type === 'test') && c.Link && START_RE_.test(c.Start) && !endedAt_(c))
      .sort((x, y) => x.Start.localeCompare(y.Start))
      .map(c => ({ id: c.ID, type: c.Type, title: c.Title, start: c.Start }));
    return { name: b.Name, count: emailsOf_(b).length, stats: stats_(own), recent, next, upcoming };
  }).sort((a, b) => b.recent.localeCompare(a.recent) || b.name.localeCompare(a.name, undefined, { numeric: true }));
  return { batches, zoom: zoomStatus_() };
}

// One batch: its emails, its classes (newest first) and totals. Fills in real lengths from Zoom.
function batchDetail(name) {
  const b = findBatch_(name);
  if (!b) throw new Error('That batch no longer exists.');
  let classes = rows_('Classes').filter(c => c.Batch === b.Name);
  if (zoomReady_()) classes = classes.map(fillFromZoom_);
  classes.sort((x, y) => (y.Start || y['Sent at']).localeCompare(x.Start || x['Sent at']));
  return {
    name: b.Name,
    emails: emailsOf_(b),
    stats: stats_(classes),
    classes: classes.map(c => ({
      type: c.Type, title: c.Title, when: whenText_(c.Start || c['Sent at'].replace(' ', 'T'), true),
      planned: Number(c['Planned min']) || 0, actual: Number(c['Actual min']) || 0,
      attended: Number(c.Attended) || 0, students: Number(c.Students) || 0,
      ended: c.Start ? endedAt_(c) : true, note: c.Type === 'noclass' ? c.Note : '',
    })),
  };
}

// ---- Zoom (Server-to-Server OAuth app; keys live in Project Settings > Script properties) ----

function zoomKeys_() {
  const p = PropertiesService.getScriptProperties();
  return {
    account: p.getProperty('ZOOM_ACCOUNT_ID'), id: p.getProperty('ZOOM_CLIENT_ID'), secret: p.getProperty('ZOOM_CLIENT_SECRET'),
    host: (p.getProperty('ZOOM_HOST_EMAIL') || '').trim(),
  };
}

// Zoom's docs: a Server-to-Server app must name the host (email or user id), not "me".
function zoomUser_() {
  const host = zoomKeys_().host;
  return host ? encodeURIComponent(host) : 'me';
}

// Checked on every page load, so problems show up in plain words before he needs Zoom.
// The token reply lists the permissions (scopes) the Zoom app was given.
function zoomStatus_() {
  if (!zoomReady_()) return { connected: false };
  try { zoomToken_(); } catch (e) { return { connected: true, create: false, recordings: false, problem: e.message }; }
  const scope = CacheService.getScriptCache().get('zoomScope') || '';
  const has = re => !scope || re.test(scope); // no scope list known: assume yes and let the call itself report
  const status = { connected: true, create: has(/meeting:write/), recordings: has(/recording:read/), past: has(/past_meeting|meeting:read/) };
  const notes = [];
  if (!status.create) notes.push('The Zoom app is missing the permission to create meetings (meeting:write:meeting:admin). Add it under Scopes on marketplace.zoom.us.');
  if (!zoomKeys_().host) notes.push('ZOOM_HOST_EMAIL is not set in Script Properties. Add the email of the Zoom account that hosts the classes.');
  if (status.create && !status.recordings) notes.push('The recordings list is off because the Zoom app cannot read cloud recordings (cloud_recording:read:list_user_recordings:admin).');
  status.problem = notes.join(' ');
  return status;
}

function zoomReady_() {
  const k = zoomKeys_();
  return Boolean(k.account && k.id && k.secret);
}

function zoomToken_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('zoomToken');
  if (hit) return hit;
  const k = zoomKeys_();
  const res = UrlFetchApp.fetch('https://zoom.us/oauth/token?grant_type=account_credentials&account_id=' + encodeURIComponent(k.account), {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(k.id + ':' + k.secret) },
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) {
    throw new Error('Zoom refused the keys (' + (body.reason || body.error || res.getResponseCode()) + '). Check ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID and ZOOM_CLIENT_SECRET.');
  }
  const ttl = Math.max(60, (body.expires_in || 3600) - 300);
  cache.put('zoomToken', body.access_token, ttl);
  cache.put('zoomScope', String(body.scope || ''), ttl);
  return body.access_token;
}

function zoom_(method, path, payload) {
  const res = UrlFetchApp.fetch('https://api.zoom.us/v2' + path, {
    method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + zoomToken_() },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  return { ok: res.getResponseCode() < 300, code: res.getResponseCode(), body };
}

// Creates a scheduled meeting on the Zoom account that owns the app, so it shows in his Zoom app.
// Meeting settings (waiting room, recording, ...) come from his Zoom account defaults.
function createZoomMeeting_(title, start, minutes) {
  const r = zoom_('post', '/users/' + zoomUser_() + '/meetings', { topic: title, type: 2, start_time: start + ':00', timezone: TIMEZONE, duration: minutes });
  if (!r.ok) throw new Error('Zoom could not create the class (' + (r.body.message || r.code) + '). Nothing was sent. Create the meeting in the Zoom app, then tap "Paste a Zoom invitation instead".');
  return r.body;
}

// Cloud recordings from the last 30 days, newest first, for the Recorded class picker.
function recentRecordings() {
  if (!zoomReady_()) return [];
  const to = new Date();
  const from = new Date(to.getTime() - 30 * 864e5);
  const day = d => d.toISOString().slice(0, 10);
  const r = zoom_('get', '/users/' + zoomUser_() + '/recordings?page_size=30&from=' + day(from) + '&to=' + day(to));
  if (!r.ok) throw new Error('Zoom could not list the recordings (' + (r.body.message || r.code) + ').');
  return (r.body.meetings || []).filter(m => m.share_url).map(m => ({
    title: m.topic || 'Class recording',
    start: localFromIso_(m.start_time),
    minutes: Number(m.duration) || 0,
    link: m.share_url,
    passcode: m.recording_play_passcode ? '' : (m.password || ''), // a play passcode is already inside share_url
  })).sort((a, b) => b.start.localeCompare(a.start));
}

// After a live class ends, ask Zoom how long it really ran and how many joined. Saved once.
function fillFromZoom_(c) {
  if (c.Type !== 'live' || !c['Meeting ID'] || c['Actual min'] || !endedAt_(c)) return c;
  try {
    const r = zoom_('get', '/past_meetings/' + encodeURIComponent(c['Meeting ID']));
    if (r.ok && r.body.duration) {
      c['Actual min'] = String(r.body.duration);
      c.Attended = String(r.body.participants_count || '');
      writeRow_('Classes', c._row, c);
    }
  } catch (e) { /* no past-meeting scope or Zoom is down: keep the planned length */ }
  return c;
}

// ---- Time helpers (all times are Bangladesh local, "2026-11-27T20:00") ----

function epochOfLocal_(s) {
  const [y, mo, d, h, mi] = s.split(/[-T:]/).map(Number);
  return Date.UTC(y, mo - 1, d, h, mi) - UTC_OFFSET_MIN * 6e4;
}

function localFromIso_(iso) {
  return new Date(new Date(iso).getTime() + UTC_OFFSET_MIN * 6e4).toISOString().slice(0, 16);
}

function endedAt_(c) {
  return epochOfLocal_(c.Start) + (Number(c['Planned min']) || 60) * 6e4 < Date.now();
}

function today_() {
  return localFromIso_(new Date().toISOString()).slice(0, 10);
}

// "2026-11-27T20:00" -> "Fri, 27 Nov 2026, 8:00 PM (Bangladesh time)"; short: "Fri, 27 Nov, 8:00 PM"
function whenText_(s, short) {
  const [y, mo, d, h, mi] = s.split(/[-T:]/).map(Number);
  const day = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][new Date(Date.UTC(y, mo - 1, d)).getUTCDay()];
  const mon = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][mo - 1];
  const clock = ((h % 12) || 12) + ':' + String(mi).padStart(2, '0') + (h < 12 ? ' AM' : ' PM');
  return short ? day + ', ' + d + ' ' + mon + ', ' + clock : day + ', ' + d + ' ' + mon + ' ' + y + ', ' + clock + ' (Bangladesh time)';
}

// ---- The email ----

// Reads Zoom's "Copy Invitation" text or a recording's "Share" text. Any https link works.
function parseInvite(text) {
  const pick = re => ((text || '').match(re) || [])[1]?.trim() || '';
  const link = pick(/(https?:\/\/[^\s<>"]+)/);
  const meetingId = pick(/Meeting ID:\s*([\d ]+\d)/i);
  return {
    link,
    meetingId,
    passcode: pick(/Pass(?:code|word):\s*(\S+)/i),
    topic: pick(/Topic:\s*(.+)/i),
    time: pick(/(?:Time|Date):\s*(.+)/i),
  };
}

const START_RE_ = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/;
const DAY_RE_ = /^\d{4}-\d{2}-\d{2}$/;
const DEFAULT_TIP = 'Consistency beats intensity. Progress over perfection.';
const TEST_CHECKLIST = [
  'Join 10 minutes early',
  'Sit somewhere quiet with steady internet',
  'Use headphones for Listening',
  'Keep paper, a pen and a charged phone or laptop ready',
];

const fmtId_ = id => String(id || '').replace(/\D/g, '').replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1 $2 $3');
const hoursText_ = m => m < 60 ? m + ' min' : Math.floor(m / 60) + ' h' + (m % 60 ? ' ' + (m % 60) + ' min' : '');
const firstUrl_ = s => ((String(s || '').match(/https?:\/\/[^\s<>"]+/) || [])[0] || '');
const minutes_ = v => Math.min(300, Math.max(15, Number(v) || 60));

// "2026-11-27" -> "Fri, 27 Nov"
function dayText_(ymd) {
  return whenText_(ymd + 'T00:00', true).replace(/, 12:00 AM$/, '');
}

function addDays_(ymd, n) {
  const [y, m, d] = ymd.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
}

// "Starts today at 8:00 PM" / "Starts tomorrow at ..." / "Started at ..." for the reminder.
function startsText_(start) {
  const today = today_(), day = start.slice(0, 10);
  const clock = whenText_(start, true).split(', ').pop();
  if (epochOfLocal_(start) <= Date.now()) return 'Started at ' + clock + '. Join now';
  if (day === today) return 'Starts today at ' + clock;
  if (day === addDays_(today, 1)) return 'Starts tomorrow at ' + clock;
  return 'Starts ' + whenText_(start, true);
}

// A meeting made on Send (live class or a test on Zoom). Before Send the preview shows placeholders.
function meeting_(title, start, minutes, real) {
  if (!real) return { link: 'https://zoom.us', meetingId: '••• •••• ••••', passcode: '••••••', id: '' };
  const m = createZoomMeeting_(title, start, minutes);
  return { link: m.join_url, meetingId: fmtId_(m.id), passcode: m.password || '', id: String(m.id) };
}

// One builder per kind of email. Each returns what the email shows (d) and what the history stores (log).
// form: { batch, type, title, tip, ... } plus the fields of its kind.
const MAKERS_ = {
  live(form, real) {
    const title = (form.title || '').trim();
    if (form.source === 'zoom') {
      if (!title) throw new Error('Give the class a title.');
      if (!START_RE_.test(form.start || '')) throw new Error('Pick the class date and time.');
      const minutes = minutes_(form.minutes);
      const m = meeting_(title, form.start, minutes, real);
      return {
        d: { title, time: whenText_(form.start), link: m.link, meetingId: m.meetingId, passcode: m.passcode, button: 'Join the class', inSubject: whenText_(form.start) },
        log: { Start: form.start, 'Planned min': minutes, 'Meeting ID': m.id, Link: real ? m.link : '', Passcode: real ? m.passcode : '' },
      };
    }
    const p = parseInvite(form.invite);
    if (!p.link) throw new Error('No link found. Paste the whole Zoom invitation.');
    const t = title || p.topic || 'IELTS live class';
    return {
      d: { title: t, time: p.time, link: p.link, meetingId: p.meetingId, passcode: p.passcode, button: 'Join the class', inSubject: p.time },
      log: { 'Meeting ID': p.meetingId.replace(/\D/g, ''), Link: p.link, Passcode: p.passcode },
    };
  },

  recorded(form) {
    const title = (form.title || '').trim();
    if (form.source === 'zoom') {
      const r = form.rec || {};
      if (!/^https:\/\//.test(r.link || '')) throw new Error('Pick a recording.');
      const ok = START_RE_.test(r.start || '');
      return {
        d: { title: title || r.title, time: ok ? whenText_(r.start) : '', link: r.link, passcode: r.passcode || '', button: 'Watch the recording' },
        log: { Start: ok ? r.start : '', 'Planned min': r.minutes || '', 'Actual min': r.minutes || '', Link: r.link },
      };
    }
    const p = parseInvite(form.invite);
    if (!p.link) throw new Error('No link found. Paste the recording link.');
    return {
      d: { title: title || p.topic || 'IELTS class recording', time: p.time, link: p.link, passcode: p.passcode, button: 'Watch the recording' },
      log: { Link: p.link },
    };
  },

  // Re-sends the link of a class or test already made for this batch. Nothing to type.
  reminder(form) {
    const c = rows_('Classes').find(x => x.ID === form.classId);
    if (!c || !c.Link || !START_RE_.test(c.Start)) throw new Error('Pick the class to remind them about.');
    const test = c.Type === 'test';
    const lead = startsText_(c.Start);
    return {
      d: {
        title: c.Title, time: whenText_(c.Start), note: lead, link: c.Link, meetingId: fmtId_(c['Meeting ID']), passcode: c.Passcode,
        button: test ? 'Open the test' : 'Join the class', subject: 'Reminder: ' + c.Title + ' ' + lead.charAt(0).toLowerCase() + lead.slice(1),
      },
      log: { Link: c.Link, Note: c.ID },
    };
  },

  test(form, real) {
    const title = (form.title || '').trim();
    if (!title) throw new Error('Give the test a title.');
    if (!START_RE_.test(form.start || '')) throw new Error('Pick the test date and time.');
    const minutes = minutes_(form.minutes);
    const onZoom = form.where !== 'link';
    let m = { link: firstUrl_(form.testLink), meetingId: '', passcode: '', id: '' };
    if (onZoom) m = meeting_(title, form.start, minutes, real);
    else if (!m.link) throw new Error('Paste the test link.');
    return {
      d: {
        title, time: whenText_(form.start) + ' · ' + hoursText_(minutes), link: m.link, meetingId: m.meetingId, passcode: m.passcode,
        button: onZoom ? 'Join the test room' : 'Open the test', checklist: TEST_CHECKLIST, inSubject: whenText_(form.start),
      },
      log: { Start: form.start, 'Planned min': minutes, 'Meeting ID': m.id, Link: real || !onZoom ? m.link : '', Passcode: real ? m.passcode : '' },
    };
  },

  // Files are already in Google Drive (uploaded when he picked them); links are added as they are.
  resources(form) {
    const title = (form.title || '').trim();
    if (!title) throw new Error('Give the resources a title.');
    const files = (form.files || []).filter(f => f && /^https:\/\//.test(f.open)).map(f => ({
      name: String(f.name), badge: fileBadge_(f.name, f.mime), meta: fileKind_(f.name, f.mime) + (f.size ? ' · ' + sizeText_(f.size) : ''), open: f.open, download: f.download,
    }));
    const links = (form.links || []).map(firstUrl_).filter(Boolean).map(u => ({ name: linkName_(u), badge: 'LINK', meta: u.replace(/^https?:\/\/(www\.)?/, '').slice(0, 60), open: u }));
    const items = files.concat(links);
    if (!items.length) throw new Error('Add a file or a link.');
    return {
      d: { title, message: (form.note || '').trim(), items, link: items.length === 1 ? items[0].open : '' },
      log: { Link: items[0].open, Note: items.map(i => i.name).join(', ') },
    };
  },

  noclass(form) {
    if (!DAY_RE_.test(form.from || '') || !DAY_RE_.test(form.back || '')) throw new Error('Pick the day off and the day classes start again.');
    if (form.back <= form.from) throw new Error('Classes must start again after the day off.');
    const last = addDays_(form.back, -1);
    const off = last === form.from ? 'on ' + dayText_(form.from) : 'from ' + dayText_(form.from) + ' to ' + dayText_(last);
    const reason = ['Holiday', 'Break', 'Class cancelled'].includes(form.reason) ? form.reason : '';
    return {
      d: { title: 'No class ' + off, time: reason, note: 'Classes start again on ' + dayText_(form.back), message: (form.note || '').trim() },
      log: { Start: form.from + 'T00:00', Note: 'back ' + form.back },
    };
  },

  message(form) {
    const title = (form.title || '').trim(), text = (form.message || '').trim();
    if (!title) throw new Error('Give the message a title.');
    if (!text) throw new Error('Write the message.');
    return { d: { title, message: text, link: '' }, log: { Note: text.slice(0, 200) } };
  },
};

function build_(form, real) {
  const kind = MAKERS_[form.type] ? form.type : 'live';
  const { d, log } = MAKERS_[kind](form, real);
  d.kind = kind;
  d.tip = kind === 'message' ? '' : (form.tip || '').trim() || DEFAULT_TIP;
  log.Type = kind; log.Title = d.title;
  const subject = d.subject || KINDS[kind].subject + d.title + (d.inSubject ? ' | ' + d.inSubject : '');
  const banner = KINDS[kind].banner;
  return { d, subject, text: plainText_(d), html: emailHtml_(d, bannerData_(banner)), mailHtml: emailHtml_(d, bannerData_(banner) && 'cid:banner'), banner, log };
}

// The plain-text part of the email (shown by mail apps that do not show designs).
function plainText_(d) {
  return [
    d.title, d.time, d.note, d.message,
    d.button && d.link && d.button + ': ' + d.link,
    d.meetingId && 'Meeting ID: ' + d.meetingId, d.passcode && 'Passcode: ' + d.passcode,
    ...(d.items || []).map(i => i.name + ': ' + i.open),
    ...(d.checklist || []).map(c => '- ' + c),
  ].filter(Boolean).join('\n');
}

// ---- Resources: files go to one Drive folder, shared as "anyone with the link can view" ----

function resourceFolder_() {
  const props = PropertiesService.getUserProperties();
  const id = props.getProperty('resFolder');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* deleted: make a new one */ } }
  const f = DriveApp.createFolder('Rimons IELTS class resources');
  props.setProperty('resFolder', f.getId());
  return f;
}

// Called once per file as soon as he picks it, so Send stays quick. data is base64.
function uploadResource(name, mime, data) {
  const file = resourceFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(data), mime || 'application/octet-stream', name));
  let shared = true;
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { shared = false; }
  const id = file.getId();
  return {
    id, name: file.getName(), size: file.getSize(), mime: file.getMimeType(), shared,
    open: 'https://drive.google.com/file/d/' + id + '/view',
    download: 'https://drive.google.com/uc?export=download&id=' + id,
  };
}

// Removing a file from the list moves it to the Drive bin (it can be restored there for 30 days).
function removeResource(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* already gone */ }
}

const EXT_ = n => (String(n).match(/\.([a-z0-9]{1,5})$/i) || [])[1] || '';
function fileBadge_(name, mime) {
  const ext = EXT_(name).toUpperCase();
  return ext ? ext.slice(0, 4) : /pdf/.test(mime) ? 'PDF' : /audio/.test(mime) ? 'MP3' : /video/.test(mime) ? 'MP4' : /image/.test(mime) ? 'IMG' : 'FILE';
}
function fileKind_(name, mime) {
  const e = EXT_(name).toLowerCase(), m = String(mime || '');
  if (e === 'pdf' || /pdf/.test(m)) return 'PDF';
  if (/audio/.test(m) || /^(mp3|m4a|wav|aac|ogg)$/.test(e)) return 'Audio';
  if (/video/.test(m) || /^(mp4|mov|mkv|webm)$/.test(e)) return 'Video';
  if (/image/.test(m) || /^(jpg|jpeg|png|webp|gif|heic)$/.test(e)) return 'Picture';
  if (/^(doc|docx|odt|rtf|txt)$/.test(e) || /word|text/.test(m)) return 'Document';
  if (/^(ppt|pptx|key)$/.test(e) || /presentation/.test(m)) return 'Slides';
  if (/^(xls|xlsx|csv)$/.test(e) || /sheet/.test(m)) return 'Spreadsheet';
  return 'File';
}
function sizeText_(b) {
  return b < 1024 * 1024 ? Math.max(1, Math.round(b / 1024)) + ' KB' : (b / 1048576).toFixed(b < 10485760 ? 1 : 0) + ' MB';
}
function linkName_(u) {
  if (/youtu\.?be/.test(u)) return 'YouTube video';
  if (/drive\.google|docs\.google/.test(u)) return 'Google Drive file';
  if (/zoom\.us\/rec/.test(u)) return 'Zoom recording';
  if (/rimonsielts\.com/.test(u)) return 'Rimons IELTS website';
  return u.replace(/^https?:\/\/(www\.)?/, '').split('/')[0];
}

// What each kind of email looks like before anything is filled in (shown as a sample on the page).
function sampleEmail(kind) {
  kind = MAKERS_[kind] ? kind : 'live';
  const t = new Date(Date.now() + UTC_OFFSET_MIN * 6e4 + 864e5).toISOString().slice(0, 10) + 'T20:00';
  const codes = { link: 'https://zoom.us', meetingId: '••• •••• ••••', passcode: '••••••' };
  const d = {
    live: Object.assign({ title: 'Your class title', time: whenText_(t), button: 'Join the class' }, codes),
    recorded: { title: 'Your recording title', link: 'https://zoom.us', passcode: '••••••', button: 'Watch the recording' },
    reminder: Object.assign({ title: 'Your class title', time: whenText_(t), note: 'Starts tomorrow at 8:00 PM', button: 'Join the class' }, codes),
    test: Object.assign({ title: 'Your test title', time: whenText_(t) + ' · 2 h 45 min', button: 'Join the test room', checklist: TEST_CHECKLIST }, codes),
    resources: { title: 'Your resources title', message: 'Your note to students', items: [
      { name: 'Reading practice test.pdf', badge: 'PDF', meta: 'PDF · 2.4 MB', open: 'https://drive.google.com', download: 'https://drive.google.com' },
      { name: 'YouTube video', badge: 'LINK', meta: 'youtube.com/watch', open: 'https://youtube.com' }] },
    noclass: { title: 'No class on Fri, 27 Nov', time: 'Holiday', note: 'Classes start again on Sat, 28 Nov' },
    message: { title: 'Your message title', message: 'Your message to students' },
  }[kind];
  d.kind = kind;
  d.tip = kind === 'message' ? '' : DEFAULT_TIP;
  return emailHtml_(d, bannerData_(KINDS[kind].banner));
}

// Called by the page whenever the form changes.
function preview(form) {
  const mail = build_(form, false);
  const to = recipients_(form.batch);
  return { subject: mail.subject, html: mail.html, count: to.length, copy: makesMeeting_(form) ? null : copyKit_(mail, to) };
}

// A new Zoom meeting does not exist until Send, so that email cannot be copied before it.
function makesMeeting_(form) {
  return (form.type === 'live' && form.source === 'zoom') || (form.type === 'test' && form.where !== 'link');
}

// What the emergency copy buttons put on the clipboard.
function copyKit_(mail, to) {
  return { html: emailHtml_(mail.d, BANNER_URL + mail.banner + '.jpg'), text: mail.text, subject: mail.subject, emails: to };
}

// Called by the Send button. Nothing is sent from here: the designed email is saved as a Gmail
// draft with every student in BCC, and the page opens Gmail so he presses the final Send himself.
// (Sending from Gmail also avoids the ~100-a-day limit on emails a script may send.)
function send(form) {
  const to = recipients_(form.batch);
  if (!to.length) throw new Error('Batch "' + form.batch + '" has no emails. Add them under Batches.');
  const mail = build_(form, true);
  const me = Session.getEffectiveUser().getEmail();
  const opts = { htmlBody: mail.mailHtml, bcc: to.join(','), name: SENDER_NAME };
  if (bannerData_(mail.banner)) opts.inlineImages = { banner: bannerBlob_(mail.banner) };
  let draft = null, draftError = '';
  try { draft = GmailApp.createDraft(me, mail.subject, mail.text, opts); } catch (e) { draftError = e.message || String(e); }
  writeRow_('Classes', 0, Object.assign(mail.log, {
    ID: Utilities.getUuid(), Batch: findBatch_(form.batch).Name, 'Sent at': localFromIso_(new Date().toISOString()).replace('T', ' '), Students: to.length,
  }));
  const gmail = 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(me);
  return {
    count: to.length,
    subject: mail.subject,
    zoom: makesMeeting_(form),
    html: mail.html,
    copy: copyKit_(mail, to),
    draftError,
    draftUrl: draft ? gmail + '#drafts?compose=' + draft.getMessage().getId() : '',
    draftsUrl: gmail + '#drafts',
  };
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

// Message text: escaped, line breaks kept, links clickable.
function richText_(s) {
  return esc_(s).replace(/https?:\/\/[^\s<>"&]+(?:&amp;[^\s<>"&]+)*/g, u => '<a href="' + u + '" style="color:#1D4ED8;word-break:break-all">' + u + '</a>').replace(/\n/g, '<br>');
}

// Inline styles only (Gmail drops most <style> rules) and email-safe tables.
// Colours and type follow the Rimons IELTS site (Aurora bento): page #F4F6FB, ink #0A1226, blue #2563EB.
// The banner travels inside the email as an inline image (cid:banner), so it needs no hosting and never breaks.
// The page preview shows the same picture as a data URL.
function bannerData_(kind) { return typeof IMAGES !== 'undefined' && IMAGES[kind] ? 'data:image/jpeg;base64,' + IMAGES[kind] : ''; }
function bannerBlob_(kind) { return Utilities.newBlob(Utilities.base64Decode(IMAGES[kind]), 'image/jpeg', kind + '.jpg'); }

function emailHtml_(d, img) {
  const k = KINDS[d.kind] || KINDS.live;
  const font = "'Plus Jakarta Sans',Arial,Helvetica,sans-serif";
  const mono = "SFMono-Regular,Consolas,'Liberation Mono',Menlo,monospace";
  const table = (style, inner) => '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"' + (style ? ' style="' + style + '"' : '') + '>' + inner + '</table>';
  const code = (label, value) => value
    ? '<tr><td style="padding:6px 0;color:#5B6478;font-size:14px">' + label + '</td>' +
      '<td align="right" style="padding:6px 0;color:#0A1226;font-size:17px;font-weight:600;font-family:' + mono + ';letter-spacing:0.02em">' + esc_(value) + '</td></tr>'
    : '';
  const pill = (href, text, dark) => '<a href="' + esc_(href) + '" style="display:inline-block;margin:0 8px 0 0;padding:9px 18px;border-radius:999px;font-size:14px;font-weight:600;text-decoration:none;' +
    (dark ? 'background:#2563EB;color:#FFFFFF' : 'background:#F4F6FB;color:#0A1226') + '">' + text + '</a>';
  const item = i => '<tr><td style="padding:14px 0;border-top:1px solid #E6E9F0">' +
    table('', '<tr><td width="44" valign="top"><div style="width:44px;height:44px;border-radius:10px;background:#EAF0FF;color:#1D4ED8;font-size:11px;font-weight:700;line-height:44px;text-align:center;letter-spacing:0.04em">' + esc_(i.badge) + '</div></td>' +
      '<td valign="top" style="padding-left:12px"><div style="font-size:15px;font-weight:600;color:#0A1226;word-break:break-word">' + esc_(i.name) + '</div>' +
      '<div style="margin-top:2px;font-size:13px;color:#5B6478;word-break:break-all">' + esc_(i.meta) + '</div>' +
      '<div style="margin-top:10px">' + pill(i.open, 'Open', true) + (i.download ? pill(i.download, 'Download') : '') + '</div></td></tr>') +
    '</td></tr>';
  const head = img ? '<img src="' + img + '" width="560" alt="' + k.label + ', Rimons IELTS" style="display:block;width:100%;max-width:560px;height:auto;border:0;border-radius:20px 20px 0 0">' : '';
  return '' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F4F6FB;font-family:' + font + '">' +
'<tr><td align="center" style="padding:16px 8px">' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px">' +
'<tr><td style="background:#FFFFFF;border-radius:20px;box-shadow:0 1px 2px rgba(10,18,38,0.06),0 8px 24px -8px rgba(10,18,38,0.12)">' +
(head && d.link ? '<a href="' + esc_(d.link) + '" style="display:block;text-decoration:none">' + head + '</a>' : head) +
'<div style="padding:24px 20px">' +
'<div style="font-size:24px;line-height:1.25;font-weight:700;color:#0A1226;letter-spacing:-0.02em">' + esc_(d.title) + '</div>' +
'<div style="margin-top:8px;font-size:15px;color:#3B4660">' +
'<span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:' + k.dot + ';margin-right:8px;vertical-align:1px"></span>' +
k.label + (d.time ? ' &middot; ' + esc_(d.time) : '') + '</div>' +
(d.note ? '<div style="margin-top:16px;padding:12px 16px;border-radius:12px;background:#EAF0FF;color:#1D4ED8;font-size:16px;font-weight:600">' + esc_(d.note) + '</div>' : '') +
(d.message ? '<div style="margin-top:16px;font-size:16px;line-height:1.6;color:#0A1226">' + richText_(d.message) + '</div>' : '') +
(d.meetingId || d.passcode
  ? table('margin-top:20px;background:#F4F6FB;border-radius:14px', '<tr><td style="padding:10px 16px">' + table('', code('Meeting ID', d.meetingId) + code('Passcode', d.passcode)) + '</td></tr>')
  : '') +
(d.button && d.link
  ? table('margin-top:24px', '<tr><td align="center" style="border-radius:999px;background:#2563EB">' +
    '<a href="' + esc_(d.link) + '" style="display:block;padding:15px 24px;color:#FFFFFF;font-size:16px;font-weight:600;text-decoration:none;border-radius:999px">' + d.button + '</a></td></tr>') +
    '<div style="margin-top:12px;font-size:13px;color:#5B6478;text-align:center;word-break:break-all">or open <a href="' + esc_(d.link) + '" style="color:#1D4ED8">' + esc_(d.link) + '</a></div>'
  : '') +
(d.items && d.items.length
  ? table('margin-top:20px', d.items.map(item).join('')) +
    '<div style="margin-top:4px;font-size:13px;line-height:1.5;color:#5B6478">Files open in Google Drive, on a phone or a computer. Tap Download to keep a copy.</div>'
  : '') +
(d.checklist
  ? '<div style="margin-top:24px;font-size:15px;font-weight:600;color:#0A1226">Before the test</div>' +
    table('margin-top:8px', d.checklist.map(c => '<tr><td width="24" valign="top" style="padding:4px 0;color:#16A34A;font-weight:700">&#10003;</td><td style="padding:4px 0;font-size:15px;color:#3B4660">' + esc_(c) + '</td></tr>').join(''))
  : '') +
(d.tip ? '<div style="margin-top:24px;padding-top:20px;border-top:1px solid #E6E9F0;font-size:15px;line-height:1.5;color:#3B4660"><span style="color:#B5660A;font-weight:600">Tip</span>&nbsp; ' + esc_(d.tip) + '</div>' : '') +
'</div></td></tr>' +
'<tr><td align="center" style="padding:16px 4px 0;font-size:13px;color:#5B6478">' +
'<a href="https://wa.me/+8801873405061" style="color:#3B4660;text-decoration:none;font-weight:600">WhatsApp</a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;' +
'<a href="https://www.facebook.com/RimonsIELTS" style="color:#3B4660;text-decoration:none;font-weight:600">Facebook</a>&nbsp;&nbsp;&middot;&nbsp;&nbsp;' +
'<a href="https://www.youtube.com/@RimonsIELTS" style="color:#3B4660;text-decoration:none;font-weight:600">YouTube</a>' +
'<div style="margin-top:8px">&copy; ' + new Date().getFullYear() + ' Rimons IELTS</div>' +
'</td></tr></table></td></tr></table>';
}

// Test copy only: run once from the editor. Makes two small batches whose "students" are your own
// Gmail with +tags (you+student1@gmail.com), so a real send lands in your own inbox.
function make_test_batches() {
  const [user, domain] = Session.getEffectiveUser().getEmail().split('@');
  const list = n => Array.from({ length: n }, (_, i) => user + '+student' + (i + 1) + '@' + domain).join('\n');
  [['Test batch 1', 3], ['Test batch 2', 2]].forEach(([name, n]) => { if (!findBatch_(name)) saveBatch('', name, list(n)); });
  db_().rename('Rimons IELTS class mail data (TEST)');
  Logger.log('Test batches ready. Reload the web app.');
}

// Run this from the Apps Script editor to check the readers. It touches no data.
function test_all() {
  const eq = (a, b, what) => { if (a !== b) throw new Error(what + ': expected "' + b + '", got "' + a + '"'); };
  const live = parseInvite('Rimon is inviting you to a scheduled Zoom meeting.\n\nTopic: Writing Task 2\nTime: Nov 27, 2026 08:00 PM Dhaka\n\nJoin Zoom Meeting\nhttps://us05web.zoom.us/j/81234567890?pwd=AbC123.1\n\nMeeting ID: 812 3456 7890\nPasscode: 4xYz9\n');
  eq(live.link, 'https://us05web.zoom.us/j/81234567890?pwd=AbC123.1', 'live link');
  eq(live.meetingId, '812 3456 7890', 'meeting id');
  eq(live.passcode, '4xYz9', 'passcode');
  eq(live.topic, 'Writing Task 2', 'topic');
  eq(live.time, 'Nov 27, 2026 08:00 PM Dhaka', 'time');
  const rec = parseInvite('Topic: Reading Class 5\nDate: Nov 20, 2026 08:00 PM Dhaka\n\nShare recording with viewers:\nhttps://us05web.zoom.us/rec/share/Xy_z-12.ab\nPasscode: a1$b2C!\n');
  eq(rec.link, 'https://us05web.zoom.us/rec/share/Xy_z-12.ab', 'rec link');
  eq(rec.passcode, 'a1$b2C!', 'rec passcode');
  eq(parseInvite('no link here').link, '', 'no link');
  eq(emailHtml_({ kind: 'live', title: '<b>x</b>', link: 'https://a.b/?q="1"', tip: 't' }).includes('<b>x</b>'), false, 'title escaped');
  eq(MAKERS_.noclass({ from: '2026-12-01', back: '2026-12-04', reason: 'Holiday' }).d.title, 'No class from Tue, 1 Dec to Thu, 3 Dec', 'days off');
  eq(MAKERS_.noclass({ from: '2026-12-01', back: '2026-12-02' }).d.note, 'Classes start again on Wed, 2 Dec', 'one day off');
  eq(fileBadge_('Reading Test 3.pdf', ''), 'PDF', 'file badge');
  eq(fileKind_('track 1.mp3', 'audio/mpeg') + ' ' + sizeText_(2621440), 'Audio 2.5 MB', 'file kind and size');
  eq(richText_('see https://a.b/x?y=1&z=2 now').includes('href="https://a.b/x?y=1&amp;z=2"'), true, 'links in messages');

  const ex = t => extractEmails_(t).emails.join(' ');
  eq(ex('Ana <Ana@x.com>, b@y.co\nana@x.com'), 'ana@x.com b@y.co', 'names, commas, repeats');
  eq(ex('a@gmail.com b@gmail.com\tc@yahoo.com;d@outlook.com|e@hotmail.com'), 'a@gmail.com b@gmail.com c@yahoo.com d@outlook.com e@hotmail.com', 'any separator');
  eq(ex('rahim@gmail.comkarim@yahoo.comsara@du.ac.bd'), 'rahim@gmail.com karim@yahoo.com sara@du.ac.bd', 'no separator at all');
  eq(ex('mailto:x@y.com. "z@w.org", (q@r.net)'), 'x@y.com z@w.org q@r.net', 'mailto and punctuation');
  eq(ex('rahim @ gmail.com'), 'rahim@gmail.com', 'spaces around @');
  eq(ex('১. rahim@gmail.com ২. karim@gmail.com'), 'rahim@gmail.com karim@gmail.com', 'bangla numbering');
  eq(extractEmails_('ok@a.com broken@gmail').skipped.join(), 'broken@gmail', 'broken address reported');

  eq(whenText_('2026-11-27T20:00'), 'Fri, 27 Nov 2026, 8:00 PM (Bangladesh time)', 'evening time');
  eq(whenText_('2026-12-01T00:30', true), 'Tue, 1 Dec, 12:30 AM', 'short time');
  eq(localFromIso_('2026-11-20T14:00:00Z'), '2026-11-20T20:00', 'utc to Dhaka');
  eq(stats_([{ Type: 'live', Start: '2026-01-01T10:00', 'Actual min': '95', 'Planned min': '90' }, { Type: 'recorded', 'Planned min': '60' }, { Type: 'live' }, { Type: 'live', Start: '2999-01-01T10:00', 'Planned min': '90' }]).minutes, 155, 'total minutes skip upcoming');
  return 'all checks passed';
}
