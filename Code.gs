/**
 * G1 Motorcycle Tech AI - Penerima Survei -> Google Sheet
 * Pasang dari Google Sheet: Ekstensi > Apps Script (script terikat ke sheet).
 * Sheet "Siswa" dan "Guru" dibuat otomatis. Kolom mengikuti pertanyaan di HTML.
 */
var SHEET_NAMES = { siswa: 'Siswa', guru: 'Guru' };

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    if (!SHEET_NAMES[d.role]) throw new Error('Peran tidak dikenal');
    if (d.hp) return out_({ ok: true });            // jebakan bot
    if (!d.answers || typeof d.answers !== 'object') throw new Error('Data kosong');
    simpan_(d);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet() {
  return out_({ ok: true, status: 'aktif', siswa: hitung_('Siswa'), guru: hitung_('Guru') });
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  // Jika script berdiri sendiri (tidak terikat), isi ID spreadsheet di bawah.
  var ID = '';
  return ID ? SpreadsheetApp.openById(ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function hitung_(nama) {
  var sh = ss_().getSheetByName(nama);
  return sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
}

function bersih_(v) {
  if (typeof v !== 'string') return v;
  v = v.trim().substring(0, 1000);
  if (/^[=+\-@]/.test(v)) v = "'" + v;               // cegah injeksi rumus
  return v;
}

function simpan_(d) {
  var sh = ss_().getSheetByName(SHEET_NAMES[d.role]) || ss_().insertSheet(SHEET_NAMES[d.role]);
  if (sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Waktu', 'Peran', 'Durasi isi (detik)']]).setFontWeight('bold').setBackground('#14232E').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var labels = d.labels || {}, row = {};
  Object.keys(d.answers).forEach(function (k) {
    var h = String(labels[k] || k).substring(0, 200);
    var i = head.indexOf(h);
    if (i < 0) {                                       // kolom baru otomatis
      head.push(h);
      sh.getRange(1, head.length).setValue(h).setFontWeight('bold').setBackground('#14232E').setFontColor('#ffffff');
      i = head.length - 1;
    }
    var val = d.answers[k];
    if (/\[(1-5|menit)\]$/.test(h) && val !== '') val = Number(val);
    row[i] = bersih_(val);
  });
  var line = new Array(head.length).fill('');
  line[0] = new Date(); line[1] = SHEET_NAMES[d.role]; line[2] = Number(d.secs) || '';
  Object.keys(row).forEach(function (i) { line[i] = row[i]; });
  sh.appendRow(line);
}

/** Menu di Google Sheet: G1 Survei > Buat / Perbarui Rekap */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('G1 Survei').addItem('Buat / Perbarui Rekap', 'buatRekap').addToUi();
}

function buatRekap() {
  var ss = ss_(), rk = ss.getSheetByName('Rekap') || ss.insertSheet('Rekap');
  rk.clear();
  var rows = [['Kelompok', 'Pernyataan', 'Jumlah responden', 'Rata-rata', '% setuju (skor 4-5)']];
  ['Siswa', 'Guru'].forEach(function (nama) {
    var sh = ss.getSheetByName(nama);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getValues(), head = data[0];
    head.forEach(function (h, c) {
      var likert = /\[1-5\]$/.test(h), menit = /\[menit\]$/.test(h);
      if (!likert && !menit) return;
      var v = [];
      for (var r = 1; r < data.length; r++) { var x = data[r][c]; if (typeof x === 'number') v.push(x); }
      if (!v.length) return;
      var avg = v.reduce(function (a, b) { return a + b; }, 0) / v.length;
      var pct = likert ? v.filter(function (n) { return n >= 4; }).length / v.length : '';
      rows.push([nama, String(h).replace(/\s*\[(1-5|menit)\]$/, '') + (menit ? ' (menit)' : ''), v.length, Math.round(avg * 100) / 100, pct]);
    });
  });
  rk.getRange(1, 1, rows.length, 5).setValues(rows);
  rk.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#14232E').setFontColor('#ffffff');
  if (rows.length > 1) rk.getRange(2, 5, rows.length - 1, 1).setNumberFormat('0%');
  rk.autoResizeColumns(1, 5);
}
