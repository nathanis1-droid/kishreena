/**
 * Kishreena Session Log — Google Apps Script backend
 *
 * Paste this into the Sheet's Apps Script editor (Extensions → Apps Script),
 * run setup() once, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone). Put the web app URL into
 * SCRIPT_URL at the top of mentors.html.
 *
 * Tabs it creates:
 *   Mentors  — one row per mentor: Email | Name | Code | Active (Y/N)
 *   Logs     — one row per logged session (written by the web page)
 *   Hours    — total minutes and sessions per mentor (formula)
 *   Pairs    — which calls each mentor–mentee pair has logged (formula)
 */

const NOTIFY_EMAIL = ''; // optional: where concern alerts go. Blank = the Sheet owner.

const LOG_HEADERS = ['Logged at', 'Mentor email', 'Mentor name', 'Mentee first name', 'Call',
                     'Call date', 'Minutes', 'Happened?', 'Her next step', 'Concern'];
const CALLS = ['Intro call (30 min)', 'Session 1', 'Session 2', 'Session 3', 'Wrap-up (30 min)', 'Extra session'];
const HAPPENED = ['Yes', 'Rescheduled', "She didn't join", "I couldn't join"];

function setup() {
  const ss = SpreadsheetApp.getActive();

  const mentors = getOrCreate_(ss, 'Mentors');
  if (mentors.getLastRow() === 0) {
    mentors.appendRow(['Email', 'Name', 'Code', 'Active (Y/N)']);
    mentors.getRange('A1:D1').setFontWeight('bold');
    mentors.setFrozenRows(1);
    mentors.getRange('F1').setValue('Add one row per mentor. Give each a 4-digit code (format column C as plain text so leading zeros stay). Set Active to N to switch someone off.');
  }
  mentors.getRange('C:C').setNumberFormat('@');

  const logs = getOrCreate_(ss, 'Logs');
  if (logs.getLastRow() === 0) {
    logs.appendRow(LOG_HEADERS);
    logs.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
    logs.setFrozenRows(1);
    // Highlight any row with a concern
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=LEN($J2)>0')
      .setBackground('#fbe3e3')
      .setRanges([logs.getRange('A2:J1000')])
      .build();
    logs.setConditionalFormatRules([rule]);
  }

  const hours = getOrCreate_(ss, 'Hours');
  hours.clear();
  hours.getRange('A1').setFormula(
    "=QUERY(Logs!A:J,\"select C, sum(G), count(G) where C is not null and H = 'Yes' group by C order by sum(G) desc label C 'Mentor', sum(G) 'Total minutes', count(G) 'Sessions'\",1)");

  const pairs = getOrCreate_(ss, 'Pairs');
  pairs.clear();
  pairs.getRange('A1').setFormula(
    "=QUERY(Logs!A:J,\"select C, D, count(E) where C is not null and H = 'Yes' group by C, D pivot E\",1)");

  const first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ---------- Web app ----------

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }

  const mentor = findMentor_(req.email, req.code);
  if (!mentor) return json_({ ok: false, error: 'That email and code don\'t match. Check your code or email Kishreena.' });

  if (req.action === 'login' || req.action === 'list') return json_({ ok: true, name: mentor.name, ...listFor_(mentor.email) });
  if (req.action === 'log') return logSession_(mentor, req.entry || {});
  return json_({ ok: false, error: 'Unknown action' });
}

function doGet() {
  return ContentService.createTextOutput('Kishreena session log is running.');
}

function findMentor_(email, code) {
  email = String(email || '').trim().toLowerCase();
  code = String(code || '').trim();
  if (!email || !code) return null;
  const rows = SpreadsheetApp.getActive().getSheetByName('Mentors').getDataRange().getValues().slice(1);
  for (const r of rows) {
    const active = String(r[3] || 'Y').trim().toUpperCase() !== 'N';
    if (active && String(r[0]).trim().toLowerCase() === email && String(r[2]).trim() === code) {
      return { email: email, name: String(r[1] || '').trim() };
    }
  }
  return null;
}

function logSession_(mentor, entry) {
  const mentee = clean_(entry.mentee, 40);
  const call = CALLS.indexOf(entry.call) >= 0 ? entry.call : '';
  const happened = HAPPENED.indexOf(entry.happened) >= 0 ? entry.happened : '';
  const date = /^\d{4}-\d{2}-\d{2}$/.test(entry.date || '') ? entry.date : '';
  const minutes = Math.round(Number(entry.minutes));
  if (!mentee || !call || !happened || !date || !(minutes >= 0 && minutes <= 240)) {
    return json_({ ok: false, error: 'Please fill in every required field.' });
  }
  const nextStep = clean_(entry.nextStep, 300);
  const concern = clean_(entry.concern, 2000);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName('Logs').appendRow([
      new Date(), mentor.email, mentor.name, mentee, call, date, minutes, happened, nextStep, concern
    ]);
  } finally {
    lock.releaseLock();
  }

  if (concern) {
    const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail(to, 'Kishreena: concern logged by ' + mentor.name,
      'Mentor: ' + mentor.name + ' (' + mentor.email + ')\nMentee: ' + mentee + '\nCall: ' + call + ' on ' + date +
      '\n\nConcern:\n' + concern + '\n\nOpen the Logs tab to follow up.');
  }
  return json_({ ok: true, name: mentor.name, ...listFor_(mentor.email) });
}

function listFor_(email) {
  const rows = SpreadsheetApp.getActive().getSheetByName('Logs').getDataRange().getValues().slice(1);
  const mine = rows.filter(r => String(r[1]).toLowerCase() === email).map(r => ({
    mentee: r[3], call: r[4],
    date: r[5] instanceof Date ? Utilities.formatDate(r[5], 'GMT', 'yyyy-MM-dd') : String(r[5]),
    minutes: Number(r[6]) || 0, happened: r[7], nextStep: r[8]
  })).sort((a, b) => (a.date < b.date ? 1 : -1));
  const totalMinutes = mine.filter(s => s.happened === 'Yes').reduce((t, s) => t + s.minutes, 0);
  return { sessions: mine, totalMinutes: totalMinutes };
}

// Strip leading characters that a spreadsheet would treat as a formula, and cap length.
function clean_(v, max) {
  let s = String(v || '').trim().slice(0, max);
  while (/^[=+\-@]/.test(s)) s = s.slice(1).trim();
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
