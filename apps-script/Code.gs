/**
 * The Poster Child — fan capture endpoint (v2: email + name + phone + IG + consent)
 *
 * SETUP (first time)
 * 1. Google Sheet → Extensions → Apps Script → replace everything with this file → save.
 * 2. Run `setup` once (approve permissions). It writes the header row.
 * 3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → copy the URL.
 *
 * UPDATING (you already have a deployment)
 * Paste this file over the old code, save, then Deploy → Manage deployments → pencil icon →
 * Version: "New version" → Deploy. The URL stays the same, so the site needs no change.
 */

const SHEET_NAME = 'Fans';
const NOTIFY_EMAIL = ''; // optional: your email to get a ping on every signup

const HEADERS = ['Timestamp', 'Email', 'Name', 'Phone', 'Instagram', 'Consent', 'Source', 'Referrer', 'User agent'];

function setup() {
  const sh = getSheet_();
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const email = String(d.email || d.value || '').trim().toLowerCase().slice(0, 120);
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, error: 'bad email' });

    const sh = getSheet_();
    if (sh.getLastRow() === 0) setup();

    // skip exact duplicate emails
    const existing = sh.getLastRow() > 1
      ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat().map(v => String(v).toLowerCase())
      : [];
    if (existing.includes(email)) return json_({ ok: true, duplicate: true });

    sh.appendRow([
      new Date(),
      email,
      String(d.name || '').slice(0, 80),
      String(d.phone || '').slice(0, 40),
      String(d.instagram || '').slice(0, 40),
      d.consent ? 'yes' : 'no',
      String(d.source || '').slice(0, 100),
      String(d.ref || '').slice(0, 200),
      String(d.ua || '').slice(0, 200)
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'New fan on the list: ' + email,
        [d.name, email, d.phone, d.instagram].filter(Boolean).join('\n'));
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, msg: 'the poster child fan capture is live' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
