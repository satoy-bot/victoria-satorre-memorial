// Betty's Place Guest Book + Photo Upload Backend
// Deploy as Google Apps Script Web App
// Execute as: Me
// Who has access: Anyone
// v2 (2026-06-10): adds api=public (JSONP) for the auto-posting tributes feed, a private
// review dashboard + api=submissions, and review_update with photo-sharing on approval.
// Preserves the existing api=dayof (Event Photos) endpoint. Private review needs Script
// Property REVIEW_TOKEN. Contact info (email/phone) is NEVER returned by the public feed.

const SHEET_ID = '1vQGaUjZbNqP3HEfQLIqsFWhqepiB_BTENu7nLorReFY';
const SUBMISSIONS_TAB = 'Submissions';
const PHOTO_FOLDER_ID = '1bC5KsfNPsNmVU_n4zZ249QLwtWJQO-sw';

// Column order in the Submissions tab (1-indexed).
const COLS = {
  timestamp: 1, name: 2, email: 3, phone: 4, relationship: 5, message: 6,
  updatePermission: 7, publicPermission: 8, photoUrl: 9, photoFilename: 10,
  source: 11, status: 12, familyNotes: 13
};

function doPost(e) {
  try {
    const data = e.parameter || {};

    if (data.action === 'review_update') {
      return handleReviewUpdate_(data);
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUBMISSIONS_TAB);
    let photoUrl = '';

    if (data.photo_base64 && data.photo_filename) {
      const bytes = Utilities.base64Decode(data.photo_base64);
      const mimeType = data.photo_mime_type || 'application/octet-stream';
      const safeName = String(data.photo_filename)
        .replace(/[\\/:*?"<>|]/g, '-')
        .slice(0, 120);
      const blob = Utilities.newBlob(bytes, mimeType, safeName);
      const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      const file = folder.createFile(blob);

      file.setDescription(
        "Uploaded through Betty's Place guest book by " + (data.name || 'Guest')
      );
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      photoUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.phone || '',
      data.relationship || '',
      data.message || '',
      data.update_permission || '',
      data.public_permission || '',
      photoUrl,
      data.photo_filename || '',
      data.source || 'bettys-place-guestbook',
      'New',
      ''
    ]);

    return json_({ ok: true, photoUrl: photoUrl });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.api === 'dayof') return listDayofPhotos_();

  // PUBLIC feed: only Approved entries whose author granted public permission.
  // JSONP via ?callback= so the static site reads it with no CORS setup.
  if (params.api === 'public') return listPublic_(params.callback || '');

  if (params.review === '1') return reviewDashboard_(params.token || '');
  if (params.api === 'submissions') return listSubmissions_(params.token || '', params.status || 'New');

  return json_({
    ok: true,
    service: "Betty's Place guest book backend"
  });
}

function listDayofPhotos_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUBMISSIONS_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_({ ok: true, items: [] });
  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
  const items = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var source = String(row[10] || '');
    var photoUrl = String(row[8] || '');

    if (source !== 'bettys-place-dayof' || !photoUrl) continue;

    var fileId = (photoUrl.match(/\/d\/([^/]+)/) || [])[1] || '';
    items.push({
      timestamp: String(row[0] || ''),
      caption: String(row[5] || ''),
      photoUrl: photoUrl,
      thumb: fileId
        ? 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800'
        : ''
    });
  }

  items.reverse();
  return json_({ ok: true, items: items });
}

function listPublic_(callback) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUBMISSIONS_TAB);
  const lastRow = sheet.getLastRow();
  let items = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, COLS.familyNotes).getValues();
    items = values.map(function (row) {
      return {
        status: String(row[COLS.status - 1] || 'New'),
        publicPermission: String(row[COLS.publicPermission - 1] || '').toLowerCase(),
        name: row[COLS.name - 1],
        relationship: row[COLS.relationship - 1],
        message: row[COLS.message - 1],
        photoUrl: row[COLS.photoUrl - 1],
        timestamp: row[COLS.timestamp - 1]
      };
    }).filter(function (it) {
      return it.status === 'Approved' && (it.publicPermission === 'yes' || it.publicPermission === 'true');
    }).map(function (it) {
      // Public projection ONLY — no email/phone ever.
      return {
        name: it.name || 'A friend',
        relationship: it.relationship || '',
        message: it.message || '',
        photoUrl: it.photoUrl || '',
        date: formatDate_(it.timestamp)
      };
    }).reverse();
  }
  const payload = { ok: true, items: items };
  if (callback && /^[\w$.]+$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function formatDate_(d) {
  if (!d) return '';
  try {
    return Utilities.formatDate(new Date(d), 'America/Los_Angeles', 'MMMM d, yyyy');
  } catch (e) { return ''; }
}

function reviewDashboard_(token) {
  if (!isValidReviewToken_(token)) {
    return HtmlService.createHtmlOutput('<h1>Review dashboard locked</h1><p>Missing or invalid review token.</p>');
  }
  var html = '<!doctype html><html><head><base target="_top">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1"><title>Review</title>' +
    '<style>body{font-family:system-ui,sans-serif;background:#f7f1ea;color:#2d2118;margin:0}' +
    'header{position:sticky;top:0;background:#fffaf4;border-bottom:1px solid #e2d3c2;padding:16px}' +
    'button,select{font:inherit}button{border:0;border-radius:999px;padding:9px 13px;background:#6d4c35;color:#fff;cursor:pointer}' +
    'button.reject{background:#9f3f36}button.private{background:#4f6475}main{padding:16px}' +
    '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px}' +
    '.card{background:#fffaf4;border:1px solid #e2d3c2;border-radius:18px;overflow:hidden}' +
    '.thumb{width:100%;aspect-ratio:1/1;object-fit:cover;background:#eadccd;display:block}' +
    '.body{padding:12px}.meta{font-size:.9rem;color:#6a5849;word-break:break-word}' +
    '.message{white-space:pre-wrap;margin:10px 0}.status{display:inline-block;border-radius:999px;padding:4px 8px;background:#eadccd;font-size:.8rem;margin-bottom:8px}' +
    'textarea{box-sizing:border-box;width:100%;min-height:60px;border:1px solid #d2bfad;border-radius:12px;padding:8px}' +
    '.actions{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}a{color:#6d4c35}</style></head><body>' +
    '<header><h1>Betty’s Place Review</h1><label>Status ' +
    '<select id="statusFilter"><option>New</option><option>Approved</option><option>Rejected</option><option>Private only</option><option>All</option></select></label> ' +
    '<button onclick="loadSubmissions()">Refresh</button> <span id="count"></span></header><main id="app"><p>Loading…</p></main><script>' +
    'var token=' + JSON.stringify(token) + ';var baseUrl=location.href.split("?")[0];' +
    'function pid(u){var m=String(u||"").match(/\\/d\\/([^/]+)/);return m?m[1]:"";}' +
    'function thumb(u){var id=pid(u);return id?"https://drive.google.com/thumbnail?id="+encodeURIComponent(id)+"&sz=w400":"";}' +
    'function safe(s){return String(s||"").replace(/[&<>\"\\047]/g,function(c){return({"&":"&amp;","<":"&lt;",">":"&gt;","\\042":"&quot;","\\047":"&#39;"})[c];});}' +
    'async function loadSubmissions(){var st=document.getElementById("statusFilter").value,app=document.getElementById("app");app.innerHTML="<p>Loading…</p>";' +
    'var res=await fetch(baseUrl+"?api=submissions&token="+encodeURIComponent(token)+"&status="+encodeURIComponent(st));var data=await res.json();' +
    'if(!data.ok){app.innerHTML="<p>"+(data.error||"Error")+"</p>";return;}document.getElementById("count").textContent=data.items.length+" item(s)";' +
    'if(!data.items.length){app.innerHTML="<p>No submissions.</p>";return;}app.innerHTML="<div class=grid>"+data.items.map(card).join("")+"</div>";}' +
    'function card(it){var t=thumb(it.photoUrl);return "<article class=card>"+(t?"<img class=thumb src=\\""+safe(t)+"\\">":"<div class=thumb></div>")+' +
    '"<div class=body><span class=status>"+safe(it.status||"New")+"</span><div class=meta><strong>"+safe(it.name||"Guest")+"</strong><br>"+safe(it.relationship)+"<br>"+safe(it.email)+" "+safe(it.phone)+"<br>public: "+safe(it.publicPermission)+"</div>"+' +
    '"<p class=message>"+safe(it.message)+"</p><div class=meta>"+safe(it.photoFilename)+(it.photoUrl?"<br><a href=\\""+safe(it.photoUrl)+"\\" target=_blank>Open photo</a>":"")+"</div>"+' +
    '"<textarea id=notes-"+it.row+">"+safe(it.familyNotes)+"</textarea><div class=actions>"+' +
    '"<button onclick=\\"upd("+it.row+",\\047Approved\\047)\\">Approve</button>"+' +
    '"<button class=reject onclick=\\"upd("+it.row+",\\047Rejected\\047)\\">Reject</button>"+' +
    '"<button class=private onclick=\\"upd("+it.row+",\\047Private only\\047)\\">Private</button></div></div></article>";}' +
    'async function upd(row,status){var notes=document.getElementById("notes-"+row).value;' +
    'var body=new URLSearchParams({action:"review_update",token:token,row:String(row),status:status,familyNotes:notes});' +
    'var res=await fetch(baseUrl,{method:"POST",body:body});var data=await res.json();if(!data.ok){alert(data.error||"Failed");return;}loadSubmissions();}' +
    'loadSubmissions();</script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle("Betty's Place Review")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function listSubmissions_(token, statusFilter) {
  if (!isValidReviewToken_(token)) return json_({ ok: false, error: 'Invalid review token' });
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUBMISSIONS_TAB);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_({ ok: true, items: [] });
  const values = sheet.getRange(2, 1, lastRow - 1, COLS.familyNotes).getValues();
  const items = values.map(function (row, index) {
    return {
      row: index + 2, timestamp: row[COLS.timestamp - 1], name: row[COLS.name - 1],
      email: row[COLS.email - 1], phone: row[COLS.phone - 1], relationship: row[COLS.relationship - 1],
      message: row[COLS.message - 1], updatePermission: row[COLS.updatePermission - 1],
      publicPermission: row[COLS.publicPermission - 1], photoUrl: row[COLS.photoUrl - 1],
      photoFilename: row[COLS.photoFilename - 1], source: row[COLS.source - 1],
      status: row[COLS.status - 1] || 'New', familyNotes: row[COLS.familyNotes - 1]
    };
  }).filter(function (item) {
    if (statusFilter === 'All') return true;
    return String(item.status || 'New') === statusFilter;
  }).reverse();
  return json_({ ok: true, items: items });
}

function handleReviewUpdate_(data) {
  if (!isValidReviewToken_(data.token || '')) return json_({ ok: false, error: 'Invalid review token' });
  const row = Number(data.row || 0);
  const status = String(data.status || '').trim();
  const allowed = ['New', 'Approved', 'Rejected', 'Private only'];
  if (!row || row < 2) return json_({ ok: false, error: 'Invalid row' });
  if (allowed.indexOf(status) === -1) return json_({ ok: false, error: 'Invalid status' });
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUBMISSIONS_TAB);
  sheet.getRange(row, COLS.status).setValue(status);
  sheet.getRange(row, COLS.familyNotes).setValue(data.familyNotes || '');

  // A photo becomes link-viewable only when Approved; otherwise return it to private.
  const photoUrl = String(sheet.getRange(row, COLS.photoUrl).getValue() || '');
  setPhotoSharing_(photoUrl, status === 'Approved');

  return json_({ ok: true, row: row, status: status });
}

function setPhotoSharing_(photoUrl, makePublic) {
  const m = String(photoUrl || '').match(/[-\w]{25,}/);
  if (!m) return;
  try {
    const file = DriveApp.getFileById(m[0]);
    if (makePublic) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } else {
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    }
  } catch (e) { /* ignore */ }
}

function isValidReviewToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('REVIEW_TOKEN');
  return Boolean(expected && token && String(token) === String(expected));
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
