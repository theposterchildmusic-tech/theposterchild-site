/**
 * The Poster Child — fan capture endpoint
 * Receives POSTs from the website and appends a row to the Google Sheet.
 *
 * SETUP
 * 1. Create a new Google Sheet. Rename the first tab to "Fans" (or change SHEET_NAME below).
 * 2. Extensions → Apps Script. Delete the default code, paste this whole file, save.
 * 3. Run the `setup` function once (select it in the toolbar, press ▶). Approve the permissions.
 *    This writes the header row.
 * 4. Deploy → New deployment → type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, copy the Web app URL.
 * 5. Paste that URL into index.html where it says PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE.
 *
 * Any time you edit this script you must Deploy → Manage deployments → Edit → "New version"
 * for the live URL to pick up the change.
 */

const SHEET_NAME = 'Fans';
const NOTIFY_EMAIL = ''; // optional: put your email here to get a ping on every new signup

function setup() {
  const sh = getSheet_();
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Type', 'Contact', 'Source', 'Referrer', 'User agent']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const type = String(data.type || '').slice(0, 20);
    const value = String(data.value || '').trim().slice(0, 120);

    if (!value || (type !== 'email' && type !== 'instagram')) {
      return json_({ ok: false, error: 'bad input' });
    }

    const sh = getSheet_();

    // skip exact duplicates
    const existing = sh.getLastRow() > 1
      ? sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().flat().map(v => String(v).toLowerCase())
      : [];
    if (existing.includes(value.toLowerCase())) {
      return json_({ ok: true, duplicate: true });
    }

    sh.appendRow([
      new Date(),
      type,
      value,
      String(data.source || '').slice(0, 100),
      String(data.ref || '').slice(0, 200),
      String(data.ua || '').slice(0, 200)
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'New fan on the list: ' + value, type + ': ' + value);
    }

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// lets you open the web app URL in a browser to confirm it's alive
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
