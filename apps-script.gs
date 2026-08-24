/**
 * Benson & Precious — QR RSVP + Wedding Day Check-In
 * Google Apps Script Web App backend.
 *
 * Setup:
 * 1. Create a Google Sheet and open Extensions > Apps Script.
 * 2. Paste this file into Code.gs.
 * 3. Set CONFIG.ADMIN_USERNAME / ADMIN_PASSWORD before deployment.
 * 4. Deploy as Web app: Execute as Me, Who has access: Anyone.
 * 5. Put the /exec URL into the main site's RSVP_API_URL and checkin.html API_URL.
 */
const CONFIG = {
  SHEET_NAME: 'Guests',
  ADMIN_USERNAME: 'usher',
  ADMIN_PASSWORD: 'CHANGE-THIS-BEFORE-DEPLOY',
  QR_BASE: 'https://api.qrserver.com/v1/create-qr-code/?size=500x500&data=',
  WEDDING_NAME: 'Benson & Precious Wedding',
  FROM_NAME: 'Benson & Precious'
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'lookup') return json_(lookupGuest_(e.parameter.guestId, e.parameter.token));
  if (action === 'stats') return json_(stats_(e.parameter.token));
  if (action === 'list') return json_(listGuests_(e.parameter.token));
  return json_({status:'ok', service:'Benson & Precious RSVP API'});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || 'rsvp';
    if (action === 'login') return json_(login_(body.username, body.password));
    if (action === 'checkin') return json_(checkin_(body.guestId, body.token));
    if (action === 'rsvp') return json_(saveRsvp_(body));
    return json_({status:'error', message:'Unknown action'});
  } catch (err) {
    return json_({status:'error', message: String(err.message || err)});
  }
}

function saveRsvp_(d) {
  const name = String(d.name || '').trim();
  const email = String(d.email || '').trim();
  const attending = String(d.attending || '').toLowerCase();
  if (!name || !email || !['yes','no'].includes(attending)) throw new Error('Name, email and attendance are required.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = getSheet_(ss);
  const guestId = uniqueGuestId_(sh);
  const partySize = Math.max(0, Math.min(10, Number(d.partySize || 0)));
  const plusOneName = String(d.plusOneName || '').trim();
  const qrData = guestId;
  const qrUrl = CONFIG.QR_BASE + encodeURIComponent(qrData);
  const now = new Date();

  sh.appendRow([guestId, now, name, email, attending, partySize, plusOneName,
                String(d.dietary || '').trim(), String(d.message || '').trim(), '', qrUrl]);

  if (attending === 'yes' && email) sendQrEmail_(email, name, guestId, partySize, qrUrl);
  return {status:'success', guestId:guestId, qrUrl:qrUrl};
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
  MailApp.sendEmail({to:email, subject:'Your wedding QR pass — Benson & Precious', htmlBody:html, inlineImages:{weddingQr:qrBlob}, name:CONFIG.FROM_NAME});
}

function login_(username, password) {
  if (String(username) !== CONFIG.ADMIN_USERNAME || String(password) !== CONFIG.ADMIN_PASSWORD) {
    return {status:'error', message:'Invalid username or password.'};
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('auth:' + token, '1', 21600); // 6 hours
  return {status:'success', token:token};
}

function checkin_(guestId, token) {
  requireAuth_(token);
  guestId = String(guestId || '').trim().toUpperCase();
  if (!/^BP-[A-Z0-9]{6}$/.test(guestId)) return {status:'not_found', message:'Invalid guest ID.'};
  const sh = getSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]).toUpperCase() === guestId) {
      const name = String(values[i][2]);
      const attending = String(values[i][4]).toLowerCase();
      const partySize = Number(values[i][5] || 0);
      const dietary = String(values[i][7] || '');
      const checkedAt = values[i][9];
      if (attending !== 'yes') return {status:'not_attending', name:name, guestId:guestId};
      if (checkedAt) return {status:'already_checked_in', name:name, guestId:guestId, partySize:partySize, dietary:dietary, checkedAt:new Date(checkedAt).toISOString()};
      const stamp = new Date();
      sh.getRange(i+1, 10).setValue(stamp);
      return {status:'checked_in', name:name, guestId:guestId, partySize:partySize, dietary:dietary, checkedAt:stamp.toISOString()};
    }
  }
  return {status:'not_found', guestId:guestId};
}

function lookupGuest_(guestId, token) {
  requireAuth_(token);
  const sh = getSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sh.getDataRange().getValues();
  guestId = String(guestId || '').trim().toUpperCase();
  for (let i=1;i<values.length;i++) if (String(values[i][0]).toUpperCase() === guestId) {
    return {status:'found', guestId:guestId, name:String(values[i][2]), attending:String(values[i][4]), partySize:Number(values[i][5]||0), dietary:String(values[i][7]||''), checkedAt:values[i][9] ? new Date(values[i][9]).toISOString() : null};
  }
  return {status:'not_found', guestId:guestId};
}

function stats_(token) {
  requireAuth_(token);
  const v = getSheet_(SpreadsheetApp.getActiveSpreadsheet()).getDataRange().getValues();
  let expected=0, checked=0;
  for(let i=1;i<v.length;i++){ if(String(v[i][4]).toLowerCase()==='yes'){ expected++; if(v[i][9]) checked++; } }
  return {status:'success', checkedIn:checked, expected:expected};
}

function listGuests_(token) {
  requireAuth_(token);
  const v = getSheet_(SpreadsheetApp.getActiveSpreadsheet()).getDataRange().getValues();
  return {status:'success', guests:v.slice(1).map(r=>({guestId:r[0],createdAt:r[1],name:r[2],email:r[3],attending:r[4],partySize:r[5],plusOneName:r[6],dietary:r[7],message:r[8],checkedAt:r[9]}))};
}

function requireAuth_(token) {
  if (!token || CacheService.getScriptCache().get('auth:' + token) !== '1') throw new Error('Unauthorized');
}

function getSheet_(ss) {
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(['Guest ID','Created At','Name','Email','Attending','Party Size','Plus One Name','Dietary','Message','Checked In At','QR URL']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function uniqueGuestId_(sh) {
  const existing = new Set(sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),1).getValues().flat().map(String));
  let id;
  do { id = 'BP-' + Utilities.getUuid().replace(/-/g,'').substring(0,6).toUpperCase(); } while(existing.has(id));
  return id;
}
function escapeHtml_(s){ return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function json_(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
