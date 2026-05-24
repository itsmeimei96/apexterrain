// ═══════════════════════════════════════════════════════════════
//  APEX TERRAIN — Google Apps Script Bridge
//  Deploy as: Extensions → Apps Script → Deploy → Web App
//  Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

const INVOICE_TAB = 'Invoice_Database';
const FORM_TAB    = 'Form Submissions';

// ── GET ─────────────────────────────────────────────────────────
// ?tab=invoices  →  all rows from Invoice_Database
// ?tab=forms     →  all rows from Form Submissions (header-mapped)
function doGet(e) {
  const tab = (e.parameter.tab || 'invoices');
  const ss  = SpreadsheetApp.getActiveSpreadsheet();

  if (tab === 'forms') {
    const sheet = ss.getSheetByName(FORM_TAB);
    if (!sheet) return json({ rows: [] });
    const vals = sheet.getDataRange().getValues();
    if (vals.length < 2) return json({ rows: [] });
    const headers = vals[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
    const rows = vals.slice(1).map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
    return json({ rows });
  }

  // Default: Invoice_Database
  // Columns: A=Invoice ID, B=Name, C=Email, D=Address, E=Status, F=Total Amount, G=Date Created
  const sheet = ss.getSheetByName(INVOICE_TAB);
  if (!sheet) return json({ rows: [] });
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return json({ rows: [] });
  const rows = vals.slice(1)
    .filter(r => r[0] && r[1] && r[2] && r[3])  // require ID + name + email + address
    .map(r => ({
      id:      String(r[0] || ''),
      name:    String(r[1] || ''),
      email:   String(r[2] || ''),
      address: String(r[3] || ''),
      status:  String(r[4] || ''),
      total:   Number(r[5]  || 0),
      date:    r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    }));
  return json({ rows });
}

// ── POST ─────────────────────────────────────────────────────────
// Body: single row object OR array of row objects
// Upserts by Invoice ID (col A). Creates new row if ID not found.
function doPost(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INVOICE_TAB);
  if (!sheet) return json({ ok: false, error: 'Sheet "' + INVOICE_TAB + '" not found' });

  let payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch(err) { return json({ ok: false, error: 'Bad JSON' }); }

  const items = Array.isArray(payload) ? payload : [payload];
  let created = 0, updated = 0;
  items.forEach(item => {
    const result = upsertRow(sheet, item);
    if (result === 'created') created++;
    else updated++;
  });
  return json({ ok: true, created, updated });
}

// ── helpers ──────────────────────────────────────────────────────
function upsertRow(sheet, item) {
  const id    = String(item.id || '');
  if (!id || !item.name || !item.email || !item.address) return 'skipped';
  const vals  = sheet.getDataRange().getValues();
  const row   = [id, item.name || '', item.email || '', item.address || '', item.status || '', item.total || 0, item.date || ''];

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === id) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([row]);
      return 'updated';
    }
  }
  sheet.appendRow(row);
  return 'created';
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
