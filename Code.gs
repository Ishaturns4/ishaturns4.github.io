// Isha's 4th birthday RSVP backend. Paste into script.google.com (Extensions > Apps Script from a Google Sheet),
// then Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone. Copy the /exec URL into index.html ENDPOINT.
const SHEET = 'RSVPs';
const HEADERS = ['Submitted', 'Attending', 'Name', 'Phone', 'Email', 'Adults', 'Kids', 'Kids names', 'Message'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.attending !== 'yes' && d.attending !== 'no') throw new Error('Invalid response');
    const name = clean_(d.name, 100);
    if (!name) throw new Error('Name is required');
    const yes = d.attending === 'yes';
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow([new Date(), yes ? 'Yes' : 'No', name, clean_(d.phone, 30), clean_(d.email, 120),
        yes ? count_(d.adults, 1) : 0, yes ? count_(d.kids, 0) : 0, clean_(d.kidNames, 300), clean_(d.message, 1000)]);
    } finally { lock.releaseLock(); }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// Trims, caps length, and stops spreadsheet formula injection (=, +, -, @ prefixes).
function clean_(v, max) {
  const s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function count_(v, min) { const n = Math.floor(Number(v)); return isFinite(n) ? Math.max(min, Math.min(20, n)) : min; }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Totals for the host on the header row (extra rows here would make appendRow skip ahead).
    sh.getRange('K1:P1').setValues([['Families attending', '=COUNTIF(B2:B,"Yes")',
      'Adults', '=SUMIF(B2:B,"Yes",F2:F)', 'Kids', '=SUMIF(B2:B,"Yes",G2:G)']]);
  }
  return sh;
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Run once from the editor to verify: adds a test row you can delete.
function testPost() {
  Logger.log(doPost({ postData: { contents: JSON.stringify({ attending: 'yes', name: 'Test Family', phone: '7805551234', adults: 2, kids: 1 }) } }).getContent());
}
