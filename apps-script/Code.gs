// Rimba Rantai - penyimpanan nilai ke Google Sheets
// GANTI PIN di bawah ini (dipakai guru untuk melihat rekap di web).
var PIN = 'guru123';
var NAMA_SHEET = 'Rekap';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(NAMA_SHEET) || ss.insertSheet(NAMA_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['ID', 'Nama', 'Misi 1', 'Misi 2', 'Misi 3', 'Misi 4', 'Misi 5', 'Rata-rata', 'Terakhir diubah']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Siswa mengirim nilai (dipanggil otomatis oleh web kuis)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // Pengaman: jika dijalankan lewat tombol Run di editor, tidak ada data kiriman
    if (!e || !e.postData) return json_({ ok: false, pesan: 'doPost hanya dipanggil oleh web kuis' });
    var d = JSON.parse(e.postData.contents);
    var id = String(d.id || '').slice(0, 60);
    var name = String(d.name || '').slice(0, 40);
    if (!id || !name) return json_({ ok: false });
    var sc = [];
    for (var i = 0; i < 5; i++) {
      var v = d.scores && d.scores[i];
      sc.push(typeof v === 'number' ? Math.max(0, Math.min(100, v)) : '');
    }
    var sum = 0;
    sc.forEach(function (x) { sum += (x === '' ? 0 : x); });
    var row = [id, name].concat(sc, [Math.round(sum / 5), new Date()]);
    var sh = getSheet_();
    var data = sh.getDataRange().getValues();
    var r = -1;
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][0]) === id) { r = j + 1; break; }
    }
    if (r > 0) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Guru melihat rekap di halaman #guru (wajib PIN)
function doGet(e) {
  if (!e || !e.parameter || e.parameter.pin !== PIN) return json_({ error: 'pin' });
  var v = getSheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    out.push({
      name: v[i][1],
      scores: [2, 3, 4, 5, 6].map(function (k) { return v[i][k] === '' ? null : v[i][k]; })
    });
  }
  return json_(out);
}

// UJI COBA: pilih fungsi "tesSimpan" lalu klik Run untuk memastikan tulis ke Sheet berhasil.
// Setelah itu cek sheet "Rekap": harus ada baris "Siswa Uji".
function tesSimpan() {
  var hasil = doPost({ postData: { contents: JSON.stringify({ id: 'siswa-uji-0001', name: 'Siswa Uji', scores: [100, 75, 50, 100, 75] }) } });
  Logger.log(hasil.getContent());
}
