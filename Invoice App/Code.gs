// ═══════════════════════════════════════════════════════════════
//  APEX TERRAIN — Google Apps Script Bridge
//  Deploy as: Extensions → Apps Script → Deploy → New Deployment
//  Type: Web App | Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

const INVOICE_TAB = 'Invoice_Database';
const FORM_TAB    = 'Form Submissions';
const OWNER_EMAIL = 'apexterrainmanagementltd@gmail.com';

// ── GET ─────────────────────────────────────────────────────────
// ?tab=invoices  →  all rows from Invoice_Database
// ?tab=forms     →  all rows from Form Submissions
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
  const sheet = ss.getSheetByName(INVOICE_TAB);
  if (!sheet) return json({ rows: [] });
  const vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return json({ rows: [] });
  const rows = vals.slice(1)
    .filter(r => r[0] && r[1] && r[2])
    .map(r => ({
      id:       String(r[0] || ''),
      name:     String(r[1] || ''),
      email:    String(r[2] || ''),
      address:  String(r[3] || ''),
      status:   String(r[4] || 'New'),
      total:    Number(r[5]  || 0),
      date:     r[6] ? Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      phone:    String(r[7] || ''),
      services: String(r[8] || ''),
      details:  String(r[9] || '')
    }));
  return json({ rows });
}

// ── POST ─────────────────────────────────────────────────────────
// From apexterrain.ca form  → has firstName/lastName fields
// From Invoice App          → has id field (invoice update)
function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let payload;
  try { payload = JSON.parse(e.postData.contents); }
  catch(err) { return json({ ok: false, error: 'Bad JSON' }); }

  const items = Array.isArray(payload) ? payload : [payload];

  // Detect if this is a website form submission (has firstName) or invoice update (has id)
  const isFormSubmission = items.some(item => item.firstName !== undefined);

  if (isFormSubmission) {
    return handleFormSubmission(ss, items[0]);
  } else {
    // Invoice app update
    const sheet = ss.getSheetByName(INVOICE_TAB);
    if (!sheet) return json({ ok: false, error: 'Sheet "' + INVOICE_TAB + '" not found' });
    let created = 0, updated = 0;
    items.forEach(item => {
      const result = upsertInvoiceRow(sheet, item);
      if (result === 'created') created++;
      else if (result === 'updated') updated++;
    });
    return json({ ok: true, created, updated });
  }
}

// ── Handle website quote form submission ─────────────────────────
function handleFormSubmission(ss, data) {
  const timestamp = new Date();
  const name = (data.firstName || '') + ' ' + (data.lastName || '');
  const invoiceId = 'APX-' + Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Math.floor(Math.random() * 900 + 100);

  // 1. Write to Form Submissions tab
  const formSheet = ss.getSheetByName(FORM_TAB);
  if (formSheet) {
    // Add header row if sheet is empty
    if (formSheet.getLastRow() === 0) {
      formSheet.appendRow(['Timestamp', 'Invoice ID', 'First Name', 'Last Name', 'Email', 'Phone', 'Address', 'Services', 'Details']);
    }
    formSheet.appendRow([
      timestamp,
      invoiceId,
      data.firstName || '',
      data.lastName  || '',
      data.email     || '',
      data.phone     || '',
      data.address   || '',
      data.services  || '',
      data.details   || ''
    ]);
  }

  // 2. Write to Invoice_Database tab
  const invoiceSheet = ss.getSheetByName(INVOICE_TAB);
  if (invoiceSheet) {
    // Add header row if sheet is empty
    if (invoiceSheet.getLastRow() === 0) {
      invoiceSheet.appendRow(['Invoice ID', 'Name', 'Email', 'Address', 'Status', 'Total Amount', 'Date Created', 'Phone', 'Services', 'Details']);
    }
    invoiceSheet.appendRow([
      invoiceId,
      name.trim(),
      data.email    || '',
      data.address  || '',
      'New',
      0,
      timestamp,
      data.phone    || '',
      data.services || '',
      data.details  || ''
    ]);
  }

  // 3. Send confirmation email to customer
  if (data.email) {
    try {
      MailApp.sendEmail({
        to: data.email,
        subject: 'We received your inquiry — Apex Terrain Management',
        htmlBody: `
          <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">
            <div style="background:#1a1a1a;padding:32px 40px;text-align:center;">
              <h1 style="color:#f28c28;font-size:28px;margin:0;letter-spacing:2px;">APEX TERRAIN</h1>
              <p style="color:rgba(255,255,255,0.5);font-size:13px;margin:6px 0 0;">Management LTD.</p>
            </div>
            <div style="padding:40px;">
              <h2 style="color:#1a1a1a;font-size:22px;margin:0 0 16px;">Hi ${data.firstName || 'there'},</h2>
              <p style="color:#444;line-height:1.7;margin:0 0 20px;">Thank you for reaching out to Apex Terrain Management. We've received your inquiry and one of our team members will be in touch within <strong>48 hours</strong> to discuss your project and schedule a site assessment.</p>
              <div style="background:#f9f9f9;border-left:4px solid #f28c28;padding:20px 24px;margin:24px 0;border-radius:0 4px 4px 0;">
                <p style="margin:0 0 8px;font-size:13px;color:#888;text-transform:uppercase;letter-spacing:1px;">Your Inquiry Summary</p>
                <p style="margin:4px 0;color:#333;font-size:14px;"><strong>Reference:</strong> ${invoiceId}</p>
                <p style="margin:4px 0;color:#333;font-size:14px;"><strong>Address:</strong> ${data.address || 'Not provided'}</p>
                <p style="margin:4px 0;color:#333;font-size:14px;"><strong>Services:</strong> ${data.services || 'Not specified'}</p>
              </div>
              <p style="color:#444;line-height:1.7;margin:0 0 32px;">In the meantime, if you have any urgent questions, feel free to call us directly at <strong>403-418-5959</strong>.</p>
              <p style="color:#444;margin:0;">Best regards,<br><strong>Apex Terrain Management LTD.</strong></p>
            </div>
            <div style="background:#f5f5f5;padding:20px 40px;text-align:center;border-top:1px solid #eee;">
              <p style="color:#999;font-size:12px;margin:0;">Mission, BC &nbsp;·&nbsp; 403-418-5959 &nbsp;·&nbsp; apexterrain.ca</p>
            </div>
          </div>
        `
      });
    } catch(err) { Logger.log('Customer email failed: ' + err); }
  }

  // 4. Send notification email to owner
  try {
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: '🌲 New Quote Request — ' + name.trim() + ' (' + invoiceId + ')',
      htmlBody: `
        <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
          <div style="background:#1a1a1a;padding:24px 32px;">
            <h2 style="color:#f28c28;margin:0;font-size:20px;">New Quote Request</h2>
            <p style="color:rgba(255,255,255,0.5);margin:4px 0 0;font-size:13px;">Submitted via apexterrain.ca</p>
          </div>
          <div style="padding:32px;background:#fff;border:1px solid #eee;">
            <table style="width:100%;border-collapse:collapse;">
              <tr><td style="padding:8px 0;color:#888;font-size:13px;width:120px;">Reference</td><td style="padding:8px 0;color:#1a1a1a;font-weight:bold;">${invoiceId}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;">Name</td><td style="padding:8px 0;color:#1a1a1a;">${name.trim()}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;">Email</td><td style="padding:8px 0;color:#1a1a1a;">${data.email || '—'}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;">Phone</td><td style="padding:8px 0;color:#1a1a1a;">${data.phone || '—'}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;">Address</td><td style="padding:8px 0;color:#1a1a1a;">${data.address || '—'}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;">Services</td><td style="padding:8px 0;color:#1a1a1a;">${data.services || '—'}</td></tr>
              <tr><td style="padding:8px 0;color:#888;font-size:13px;vertical-align:top;">Details</td><td style="padding:8px 0;color:#1a1a1a;">${data.details || '—'}</td></tr>
            </table>
            <div style="margin-top:24px;padding-top:24px;border-top:1px solid #eee;">
              <a href="https://docs.google.com/spreadsheets/d/1Px2WzEXuSe6y87tDdKCA_7EgYhwQaaEQ5kot-7WUzKQ/edit" style="background:#f28c28;color:#fff;padding:12px 24px;text-decoration:none;font-weight:bold;font-size:14px;border-radius:4px;">View in Google Sheets</a>
              &nbsp;&nbsp;
              <a href="https://apexinvoice-six.vercel.app/" style="background:#1a1a1a;color:#fff;padding:12px 24px;text-decoration:none;font-weight:bold;font-size:14px;border-radius:4px;">Open Invoice App</a>
            </div>
          </div>
        </div>
      `
    });
  } catch(err) { Logger.log('Owner email failed: ' + err); }

  return json({ ok: true, invoiceId });
}

// ── Upsert invoice row (from Invoice App) ────────────────────────
function upsertInvoiceRow(sheet, item) {
  const id = String(item.id || '');
  if (!id) return 'skipped';
  const vals = sheet.getDataRange().getValues();
  const row = [
    id,
    item.name    || '',
    item.email   || '',
    item.address || '',
    item.status  || 'New',
    item.total   || 0,
    item.date    || new Date(),
    item.phone   || '',
    item.services|| '',
    item.details || ''
  ];
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === id) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([row]);
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
