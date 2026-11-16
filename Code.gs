// ══════════════════════════════════════════════════════════════════════════════
//  Dollar Tracker — Google Apps Script Backend
//
//  HOW TO DEPLOY:
//  1. Open Google Sheets → Extensions → Apps Script
//  2. Replace Code.gs with this entire file → Save
//  3. Deploy → New Deployment → Type: Web App
//     Execute as: Me   |   Who has access: Anyone
//  4. Copy the Web App URL → paste into the HTML app Settings page
// ══════════════════════════════════════════════════════════════════════════════

var SHEET_NAME = 'Transactions';

var HEADERS = ['id', 'date', 'time', 'bankAmount', 'dollarAmount', 'exchangeRate', 'sellPrice', 'note', 'createdAt'];

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    var hdr = sheet.getRange(1, 1, 1, HEADERS.length);
    hdr.setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);
  }
  return sheet;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── doGet — all operations via GET params (works anonymously, no CORS preflight) ──
function doGet(e) {
  try {
    var action = (e.parameter && e.parameter.action) ? e.parameter.action : 'getAll';
    var sheet  = getSheet();

    if (action === 'getAll') {
      var values = sheet.getDataRange().getValues();
      if (values.length <= 1) return respond({ ok: true, data: [] });
      var rows = values.slice(1).map(function(row) {
        var obj = {};
        HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
        return obj;
      }).sort(function(a, b) {
        return new Date(b.createdAt) - new Date(a.createdAt);
      });
      return respond({ ok: true, data: rows });
    }

    if (action === 'add') {
      var tx = JSON.parse(decodeURIComponent(e.parameter.data));
      tx.id        = tx.id        || (Date.now().toString(36) + Math.random().toString(36).slice(2));
      tx.createdAt = tx.createdAt || new Date().toISOString();
      sheet.appendRow(HEADERS.map(function(h) { return tx[h] || ''; }));
      return respond({ ok: true, data: tx });
    }

    if (action === 'delete') {
      var id     = e.parameter.id;
      var values = sheet.getDataRange().getValues();
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(id)) {
          sheet.deleteRow(i + 1);
          return respond({ ok: true });
        }
      }
      return respond({ ok: false, error: 'Row not found' });
    }

    return respond({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ ok: false, error: err.message, data: [] });
  }
}
