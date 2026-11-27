/**
 * HT-KITE-01 · Καταγραφή διαδικασίας προσομοίωσης χαρταετού
 * Google Apps Script (Web app). Δέχεται POST από kite/index.html και γράφει σε δύο φύλλα:
 *   events    — μία γραμμή ανά ενέργεια (φάση, απάντηση, έλεγχος, δοκιμή, κείμενο)
 *   snapshots — μία γραμμή ανά «Υποβολή για την έρευνα» (τελική εικόνα συμμετέχοντα)
 *
 * Εγκατάσταση:
 * 1. Νέο Google Sheet (ΧΩΡΙΣΤΟ από το BRaL) → Επεκτάσεις → Apps Script → επικολλήστε αυτό το αρχείο.
 * 2. Ανάπτυξη → Νέα ανάπτυξη → Τύπος: Web app · Εκτέλεση ως: Εγώ · Πρόσβαση: Οποιοσδήποτε.
 * 3. Αντιγράψτε το URL (/exec) στο CONFIG.research.endpoint μέσα στο kite/index.html.
 *
 * Σύνδεση με BRaL και PRE/POST: η στήλη `code` χρησιμοποιεί την ίδια φόρμουλα κωδικού.
 */
const EVENT_COLS = ['received', 'sid', 'code', 'group', 'seq', 'ts', 'phase', 'type', 'data', 'study', 'version'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    if (body.kind === 'events') {
      const sh = sheet_(ss, 'events', EVENT_COLS);
      const rows = (body.events || []).map(ev => EVENT_COLS.map(c =>
        c === 'received' ? now : c === 'study' ? body.study : c === 'version' ? body.version : (ev[c] === undefined ? '' : ev[c])));
      if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, EVENT_COLS.length).setValues(rows);
    } else if (body.kind === 'snapshot') {
      const keys = ['received'].concat(Object.keys(body));
      const sh = sheet_(ss, 'snapshots', keys);
      let header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const missing = keys.filter(k => header.indexOf(k) < 0);
      if (missing.length) {
        sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
        header = header.concat(missing);
      }
      const row = header.map(h => h === 'received' ? now : (body[h] === undefined ? '' : body[h]));
      sh.appendRow(row);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  }
  return sh;
}
