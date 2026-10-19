// الصق الكود ده كله في Apps Script (Extensions > Apps Script) ثم Deploy كـ Web app
var SHEET_NAME = 'البيانات';

function sh_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    var s = sh_();
    var v = s.getDataRange().getDisplayValues();
    if (!v.length || !v[0][0]) return out_({ ok: true, cols: [], rows: [] });
    var cols = [];
    for (var j = 1; j < v[0].length; j++) { if (v[0][j] === '') break; cols.push(v[0][j]); }
    var rows = [];
    for (var i = 1; i < v.length; i++) {
      if (!v[i][0]) continue;
      var o = { _id: v[i][0] };
      for (var k = 0; k < cols.length; k++) o[cols[k]] = v[i][k + 1];
      rows.push(o);
    }
    return out_({ ok: true, cols: cols, rows: rows });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var d = JSON.parse(e.postData.contents);
    var s = sh_();
    var cols = d.cols || [];
    var head = ['_id'].concat(cols);
    var cur = s.getRange(1, 1, 1, head.length).getValues()[0];
    if (cur.join('|') !== head.join('|')) {
      var maxC = s.getMaxColumns();
      if (maxC > head.length) s.getRange(1, head.length + 1, 1, maxC - head.length).clearContent();
      s.getRange(1, 1, 1, head.length).setNumberFormat('@').setValues([head]).setFontWeight('bold');
      s.setFrozenRows(1);
      s.setRightToLeft(true);
      s.hideColumns(1);
    }
    var last = s.getLastRow();
    var ids = last > 1 ? s.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var idx = ids.indexOf(String(d.id));
    if (d.action === 'delete') {
      if (idx > -1) s.deleteRow(idx + 2);
    } else if (d.action === 'add' || d.action === 'update') {
      var vals = [d.id].concat(cols.map(function (c) { return d.row[c] == null ? '' : String(d.row[c]); }));
      var r = idx > -1 ? idx + 2 : last + 1;
      var maxC2 = s.getMaxColumns();
      if (maxC2 > vals.length) s.getRange(r, vals.length + 1, 1, maxC2 - vals.length).clearContent();
      s.getRange(r, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.release();
  }
}
