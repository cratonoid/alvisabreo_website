// Google Apps Script backend for facts.html.
// Paste this into Extensions → Apps Script of the facts Google Sheet.
// Put your real password in ADMIN_KEY inside the Apps Script editor only —
// don't commit it here, this repo is public.
const ADMIN_KEY = 'CHANGE_ME';
const SHEET_NAME = 'Facts';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET: returns every fact in column A, with its row number (used for deleting).
function doGet() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return json({ ok: true, facts: [] });

  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  const facts = [];
  values.forEach(function(row, i) {
    const text = String(row[0]).trim();
    if (text) facts.push({ row: i + 1, text: text });
  });
  return json({ ok: true, facts: facts });
}

// POST body (JSON): { key, action: 'check' | 'add' | 'edit' | 'delete', text, row, newText }
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'Bad request' });
  }

  if (body.key !== ADMIN_KEY) return json({ ok: false, error: 'Wrong password' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet();

    if (body.action === 'check') {
      return json({ ok: true });
    }

    if (body.action === 'add') {
      const text = String(body.text || '').trim();
      if (!text) return json({ ok: false, error: 'Empty fact' });
      sheet.appendRow([text, new Date()]);
      return json({ ok: true });
    }

    if (body.action === 'delete' || body.action === 'edit') {
      const row = Number(body.row);
      // Only touch the row if it still holds the same fact, in case the sheet changed.
      if (!(row >= 1 && row <= sheet.getLastRow() &&
            String(sheet.getRange(row, 1).getValue()).trim() === body.text)) {
        return json({ ok: false, error: 'Fact not found. Refresh and try again.' });
      }

      if (body.action === 'delete') {
        sheet.deleteRow(row);
        return json({ ok: true });
      }

      const newText = String(body.newText || '').trim();
      if (!newText) return json({ ok: false, error: 'Empty fact' });
      sheet.getRange(row, 1).setValue(newText);
      return json({ ok: true });
    }

    return json({ ok: false, error: 'Unknown action' });
  } finally {
    lock.releaseLock();
  }
}
