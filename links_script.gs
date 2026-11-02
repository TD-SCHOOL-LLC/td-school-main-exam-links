/**
 * TD School – Linklər səhifəsi üçün Google Apps Script
 * Cədvəldə "Links" vərəqi yaradır: ID, Səviyyə, Ad, Link, Tarix
 */
const SHEET_NAME = 'Links';
const ADMIN_PASSWORD = 'td.link2026';
const LEVELS = ['Pre-A1', 'A1', 'A2/B1', 'B1+/B2', 'C1/C2'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['ID', 'Səviyyə', 'Ad', 'Link', 'Tarix']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function safe_(v) {
  v = String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doGet() {
  const v = sheet_().getDataRange().getValues();
  const links = [];
  for (let i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    links.push({ id: String(v[i][0]), level: v[i][1], title: v[i][2], url: v[i][3] });
  }
  return json_({ links: links });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d.password !== ADMIN_PASSWORD) return json_({ ok: false, error: 'Parol səhvdir' });
    const sh = sheet_();

    if (d.action === 'add') {
      if (LEVELS.indexOf(d.level) < 0) return json_({ ok: false, error: 'Səviyyə yanlışdır' });
      if (!d.title || !/^https?:\/\//i.test(d.url || '')) return json_({ ok: false, error: 'Ad və ya link yanlışdır' });
      sh.appendRow([Utilities.getUuid(), d.level, safe_(d.title), String(d.url).trim(), new Date()]);
      return json_({ ok: true });
    }
    if (d.action === 'move') {
      const v = sh.getDataRange().getValues();
      const cols = v[0].length;
      let idx = -1;
      for (let i = 1; i < v.length; i++) { if (String(v[i][0]) === String(d.id)) { idx = i; break; } }
      if (idx < 0) return json_({ ok: false, error: 'Link tapılmadı' });
      const level = v[idx][1];
      const step = d.direction === 'up' ? -1 : 1;
      let j = idx + step;
      while (j >= 1 && j < v.length && v[j][1] !== level) j += step;
      if (j < 1 || j >= v.length) return json_({ ok: true }); // artıq kənardadır
      sh.getRange(idx + 1, 1, 1, cols).setValues([v[j]]);
      sh.getRange(j + 1, 1, 1, cols).setValues([v[idx]]);
      return json_({ ok: true });
    }
    if (d.action === 'delete') {
      const v = sh.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][0]) === String(d.id)) { sh.deleteRow(i + 1); return json_({ ok: true }); }
      }
      return json_({ ok: false, error: 'Link tapılmadı' });
    }
    return json_({ ok: false, error: 'Naməlum əməliyyat' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
