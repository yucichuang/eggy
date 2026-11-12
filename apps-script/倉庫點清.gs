const SHEET = '清單';

function doGet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const n = Math.max(sh.getLastRow() - 1, 1);
  const v = sh.getRange(2, 1, n, 11).getValues();
  const rows = [];
  v.forEach((r, i) => {
    if (r[2] !== '') rows.push({ row: i + 2, pos: r[0], bag: r[1], name: r[2], ser: r[3], qty: r[4], pts: r[5], go: r[6], box: r[7], st: r[8], note: r[10] });
  });
  return out({ ok: true, rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
    if (d.action === 'add') {
      const names = sh.getRange(2, 3, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
      let last = 1;
      names.forEach((r, i) => { if (r[0] !== '') last = i + 2; });
      const row = last + 1;
      const x = d.item;
      sh.getRange(row, 1, 1, 7).setValues([[x.pos, x.bag, x.name, x.ser, x.qty, x.pts, x.go]]);
      return out({ ok: true, row });
    }
    if (d.action === 'update') {
      const col = { bag: 2, name: 3, ser: 4, qty: 5, pts: 6, go: 7, box: 8, note: 11 };
      Object.keys(d.data).forEach(k => { if (col[k]) sh.getRange(d.row, col[k]).setValue(d.data[k]); });
      return out({ ok: true });
    }
    if (d.action === 'delete') {
      sh.getRange(d.row, 1, 1, 11).clearContent();
      return out({ ok: true });
    }
    return out({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
