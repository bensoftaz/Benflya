/**
 * Benson & Precious — QR RSVP + Wedding Day Check-In + Payment Tracking
 * Google Apps Script Web App backend.
 *
 * SETUP
 * 1. Open your Google Sheet > Extensions > Apps Script.
 * 2. Replace the ENTIRE contents of Code.gs with this file.
 * 3. Update CONFIG.PAYMENT_IMAGE_URL below if your filename differs.
 * 4. Deploy > Manage deployments > (pencil/edit icon) > Version: "New version" > Deploy.
 *    (Editing code alone does NOT update the live /exec URL — you must redeploy a new version.)
 * 5. Confirm "Who has access" is set to "Anyone".
 * 6. No login/token step is required — checkin.html's own on-page login
 *    (Admin / EventsC01 / EventsC02) is the only access gate. This script
 *    trusts any request that reaches it, same as before.
 */
const CONFIG = {
  SHEET_NAME: 'Guests',
  QR_BASE: 'https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=',
  PAYMENT_IMAGE_URL: 'https://benflya.online/images/payment-details.png',
  WEDDING_NAME: 'Benson & Precious Wedding',
  FROM_NAME: 'Benson & Precious',
  RSVP_AMOUNT: 'QAR 130'
};

// Column layout (1-based sheet columns / 0-based array indices):
// 1  Guest ID          (0)
// 2  Created At        (1)
// 3  Name              (2)
// 4  Email             (3)
// 5  Attending         (4)
// 6  Party Size        (5)
// 7  Plus One Name     (6)
// 8  Dietary           (7)
// 9  Message           (8)
// 10 Checked In At     (9)
// 11 QR URL            (10)
// 12 Phone             (11)
// 13 Payment Status    (12)
// 14 Payment Email Sent At (13)

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = p.action || '';
    if (action === 'list') return json_(listGuests_());
    if (action === 'stats') return json_(stats_());
    if (action === 'checkin') return json_(checkin_(p.id));
    if (action === 'sendPayment') return json_(sendPayment_(p.id));
    if (action === 'setPayment') return json_(setPayment_(p.id, p.status));
    return json_({ status: 'ok', service: 'Benson & Precious RSVP API' });
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = body.action || 'rsvp';
    if (action === 'rsvp') return json_(saveRsvp_(body));
    return json_({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

/* ---------------- RSVP submission ---------------- */

function saveRsvp_(d) {
  const name = String(d.name || '').trim();
  const email = String(d.email || '').trim();
  const phone = String(d.phone || '').trim();
  const attending = String(d.attending || '').toLowerCase();
  if (!name || !email || !['yes', 'no'].includes(attending)) {
    throw new Error('Name, email and attendance are required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheet_(ss);
  const guestId = uniqueGuestId_(sh);
  const partySize = Math.max(0, Math.min(10, Number(d.partySize || 0)));
  const plusOneName = String(d.plusOneName || '').trim();
  const qrUrl = CONFIG.QR_BASE + encodeURIComponent(guestId);
  const now = new Date();

  sh.appendRow([
    guestId, now, name, email, attending, partySize, plusOneName,
    String(d.dietary || '').trim(), String(d.message || '').trim(),
    '', qrUrl, phone, '', ''
  ]);

  if (attending === 'yes' && email) sendQrEmail_(email, name, guestId, partySize, qrUrl);
  return { status: 'success', guestId: guestId, qrUrl: qrUrl };
}

function sendQrEmail_(email, name, guestId, partySize, qrUrl) {
  const party = partySize ? ' + ' + partySize + (partySize === 1 ? ' guest' : ' guests') : '';
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName('wedding-qr.png');
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto">' +
    '<h2 style="color:#55162A">Your Wedding QR Pass</h2>' +
    '<p>Dear ' + escapeHtml_(name) + ',</p>' +
    '<p>Your RSVP for <b>' + CONFIG.WEDDING_NAME + '</b> has been received.</p>' +
    '<p><b>Guest ID:</b> ' + guestId + '<br><b>Party:</b> ' + escapeHtml_(name + party) + '</p>' +
    '<p>Show the QR code below at the entrance on October 10, 2026.</p>' +
    '<p><img src="cid:weddingQr" width="300" height="300" alt="Wedding QR code"></p>' +
    '<p>Please save this email or QR image. A second scan will be flagged automatically.</p>' +
    '<p>With love,<br>' + CONFIG.FROM_NAME + '</p></div>';
  MailApp.sendEmail({ to: email, subject: 'Your wedding QR pass — Benson & Precious', htmlBody: html, inlineImages: { weddingQr: qrBlob }, name: CONFIG.FROM_NAME });
}

/* ---------------- Payment details email + status ---------------- */

function sendPayment_(guestId) {
  guestId = String(guestId || '').trim().toUpperCase();
  if (!guestId) return { status: 'error', message: 'Missing guest ID.' };

  const sh = getSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === guestId) {
      const name = String(values[i][2]);
      const email = String(values[i][3] || '');
      if (!email) return { status: 'error', message: 'No email on file for this guest.' };
      sendPaymentEmail_(email, name, guestId);
      sh.getRange(i + 1, 14).setValue(new Date()); // column 14 = Payment Email Sent At
      return { status: 'success', email: email };
    }
  }
  return { status: 'error', message: 'Guest not found.' };
}

function sendPaymentEmail_(email, name, guestId) {
  const html = '<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto">' +
    '<h2 style="color:#55162A">Wedding Payment Details</h2>' +
    '<p>Dear ' + escapeHtml_(name) + ',</p>' +
    '<p>Thank you for RSVPing to <b>' + CONFIG.WEDDING_NAME + '</b>! To secure your seat, please complete payment of <b>' + CONFIG.RSVP_AMOUNT + '</b> using the details below.</p>' +
    '<p><img src="' + CONFIG.PAYMENT_IMAGE_URL + '" width="400" alt="Payment details"></p>' +
    '<p><b>Important:</b> please use your Guest ID as the payment reference: <b>' + guestId + '</b></p>' +
    '<p>Once we confirm your payment, your wedding QR pass will follow separately.</p>' +
    '<p>With love,<br>' + CONFIG.FROM_NAME + '</p></div>';
  MailApp.sendEmail({ to: email, subject: 'Payment details — Benson & Precious Wedding', htmlBody: html, name: CONFIG.FROM_NAME });
}

function setPayment_(guestId, status) {
  status = String(status || '').toLowerCase();
  if (!['paid', 'unpaid'].includes(status)) return { status: 'error', message: 'Invalid status.' };
  guestId = String(guestId || '').trim().toUpperCase();

  const sh = getSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === guestId) {
      sh.getRange(i + 1, 13).setValue(status); // column 13 = Payment Status
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Guest not found.' };
}

/* ---------------- Check-in (scanner) ---------------- */

function checkin_(guestId) {
  guestId = String(guestId || '').trim().toUpperCase();
  if (!guestId) return { status: 'not_found' };

  const sh = getSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).toUpperCase() === guestId) {
      const name = String(values[i][2]);
      const attending = String(values[i][4]).toLowerCase();
      const plusOneName = String(values[i][6] || '');
      const dietary = String(values[i][7] || '');
      const checkedAt = values[i][9];

      if (attending !== 'yes') return { status: 'not_attending', name: name };
      if (checkedAt) {
        return { status: 'already', name: name, time: new Date(checkedAt).toISOString(), plusOne: !!plusOneName, plusOneName: plusOneName, dietary: dietary };
      }
      const stamp = new Date();
      sh.getRange(i + 1, 10).setValue(stamp);
      return { status: 'ok', name: name, plusOne: !!plusOneName, plusOneName: plusOneName, dietary: dietary };
    }
  }
  return { status: 'not_found' };
}

/* ---------------- Stats + list (admin dashboard) ---------------- */

function stats_() {
  const v = getSheet_(SpreadsheetApp.getActiveSpreadsheet()).getDataRange().getValues();
  let attending = 0, declined = 0, checkedIn = 0, paid = 0;
  for (let i = 1; i < v.length; i++) {
    const att = String(v[i][4]).toLowerCase();
    if (att === 'yes') attending++;
    if (att === 'no') declined++;
    if (v[i][9]) checkedIn++;
    if (String(v[i][12] || '').toLowerCase() === 'paid') paid++;
  }
  return { status: 'success', attending: attending, declined: declined, totalGuests: v.length - 1, checkedIn: checkedIn, paid: paid };
}

function listGuests_() {
  const v = getSheet_(SpreadsheetApp.getActiveSpreadsheet()).getDataRange().getValues();
  const responses = v.slice(1).map(function (r) {
    return {
      GuestId: String(r[0] || ''),
      Name: String(r[2] || ''),
      Email: String(r[3] || ''),
      Attending: String(r[4] || ''),
      PartySize: Number(r[5] || 0),
      PlusOneName: String(r[6] || ''),
      PlusOne: r[6] ? 'yes' : 'no',
      Dietary: String(r[7] || ''),
      Message: String(r[8] || ''),
      CheckedIn: r[9] ? 'yes' : 'no',
      CheckedInTime: r[9] ? new Date(r[9]).toISOString() : '',
      Phone: String(r[11] || ''),
      PaymentStatus: String(r[12] || ''),
      PaymentEmailSentAt: r[13] ? new Date(r[13]).toISOString() : ''
    };
  });
  return { status: 'success', responses: responses };
}

/* ---------------- Sheet + helpers ---------------- */

function getSheet_(ss) {
  const headers = ['Guest ID', 'Created At', 'Name', 'Email', 'Attending', 'Party Size', 'Plus One Name', 'Dietary', 'Message', 'Checked In At', 'QR URL', 'Phone', 'Payment Status', 'Payment Email Sent At'];
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else {
    const lastCol = sh.getLastColumn();
    if (lastCol < headers.length) {
      sh.getRange(1, lastCol + 1, 1, headers.length - lastCol).setValues([headers.slice(lastCol)]);
    }
  }
  return sh;
}

function uniqueGuestId_(sh) {
  const existing = new Set(sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues().flat().map(String));
  let id;
  do { id = 'BP-' + Utilities.getUuid().replace(/-/g, '').substring(0, 6).toUpperCase(); } while (existing.has(id));
  return id;
}

function escapeHtml_(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
