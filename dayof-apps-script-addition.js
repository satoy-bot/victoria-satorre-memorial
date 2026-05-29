// ============================================================
// Betty's Place — Day-Of Photos endpoint
// ADD TO THE END of Code.gs in your Apps Script project.
// Then add ONE line inside doGet() before the final return:
//
//   if (params.api === 'dayof') return listDayofPhotos_();
//
// Then Redeploy → Manage deployments → Edit → New version → Deploy.
// ============================================================

function listDayofPhotos_() {
  const sheet = getSubmissionsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_({ ok: true, items: [] });

  const numCols = COLS.familyNotes;
  const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const items = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var source   = String(row[COLS.source   - 1] || '');
    var photoUrl = String(row[COLS.photoUrl - 1] || '');

    if (source !== 'bettys-place-dayof' || !photoUrl) continue;

    var fileId = (photoUrl.match(/\/d\/([^/]+)/) || [])[1] || '';
    items.push({
      timestamp: String(row[COLS.timestamp - 1] || ''),
      caption:   String(row[COLS.message   - 1] || ''),
      photoUrl:  photoUrl,
      thumb:     fileId
        ? 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800'
        : ''
    });
  }

  items.reverse(); // newest first
  return json_({ ok: true, items: items });
}
