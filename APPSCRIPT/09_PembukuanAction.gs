function filterPembukuan(e){
  // Jika edit bukan di E2 atau F2, stop
  var range = e.range;

  if (!(range.getA1Notation() === "E2" || range.getA1Notation() === "F2")) return;

  // Ambil tanggal awal dan akhir
  var startDate = new Date(DBPembukuanSheet.getRange("E2").getValue());
  var endDate = new Date(DBPembukuanSheet.getRange("F2").getValue());

  if (!startDate || !endDate) return; // kalau salah satu kosong, stop

  // Ambil semua data di tabel
  var lastRow = DBPembukuanSheet.getLastRow();
  var dates = DBPembukuanSheet.getRange(5, 3, lastRow - 4, 1).getValues(); // kolom C mulai baris 5

  // Reset semua baris agar terlihat
  DBPembukuanSheet.showRows(5, lastRow - 4);

  // Loop dan sembunyikan jika di luar rentang
  for (var i = 0; i < dates.length; i++) {
    var rowDate = new Date(dates[i][0]);
    if (isNaN(rowDate)) continue; // skip kalau bukan tanggal
    if (rowDate < startDate || rowDate > endDate) {
      DBPembukuanSheet.hideRows(i + 5); // baris mulai dari 5
    }
  }
}

function getFormDataPembukuan() {
  return {
    idPembukuan : FormPembukuanSheet.getRange('D4').getValue(),
    tanggal     : FormPembukuanSheet.getRange('D6').getValue(),
    jenis       : FormPembukuanSheet.getRange('D8').getValue(),
    kategori    : FormPembukuanSheet.getRange('D10').getValue(),
    jumlah      : FormPembukuanSheet.getRange('D12').getValue(),
    keterangan  : FormPembukuanSheet.getRange('D14').getValue(),
  };
}

function setFormDataPembukuan(data) {
  FormPembukuanSheet.getRange('D4').setValue(data.idPembukuan || '');
  FormPembukuanSheet.getRange('D6').setValue(data.tanggal || '');
  FormPembukuanSheet.getRange('D8').setValue(data.jenis || '');
  FormPembukuanSheet.getRange('D10').setValue(data.kategori || '');
  FormPembukuanSheet.getRange('D12').setValue(data.jumlah || '');
  FormPembukuanSheet.getRange('D14').setValue(data.keterangan || '');
}

function generateKodePembukuan() {

  const lastRow = DBPembukuanSheet.getLastRow();

  if (lastRow <= 4) {
    DBPembukuanSheet.getRange("D4").setValue('PBKUAN-0001')
  }

  const kodeTerakhir = DBPembukuanSheet.getRange(lastRow, 2).getValue(); // Kolom B = Kode Pembelian

  const nomorTerakhir = parseInt(kodeTerakhir.split('-')[1]) || 0;
  const kodeBaru = 'PBKUAN-' + String(nomorTerakhir + 1).padStart(4, '0');

  FormPembukuanSheet.getRange("D4").setValue(kodeBaru);

  return {
    IdPembukuan: kodeBaru
  }
}

function clearFormPembukuan() {
  FormPembukuanSheet.getRange('D4:D14').clearContent();
  generateKodePembukuan();
  setTanggalPembukuan();
}

function searchPembukuan() {
  const query = FormPembukuanSheet.getRange('D4').getValue().toString().trim();
  if (!query) {
    SpreadsheetApp.getUi().alert("Masukkan ID Pembukuan untuk mencari.");
    return;
  }

  const data = DBPembukuanSheet.getDataRange().getValues();
  const headerRow = 4;

  for (let i = headerRow; i < data.length; i++) {
    if (data[i][1] === query) {
      setFormDataPembukuan({
        idPembukuan : data[i][1],
        tanggal     : data[i][2],
        jenis       : data[i][3],
        kategori    : data[i][4],
        jumlah      : data[i][5],
        keterangan  : data[i][6]
      });
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

function setTanggalPembukuan(){
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone("Asia/Jakarta");

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "Asia/Jakarta", 'dd/MM/yy');

  FormPembukuanSheet.getRange("D6").setValue(formattedDate);
}

function simpanPembukuan() {
  const data = getFormDataPembukuan();

  if (!data.idPembukuan || !data.tanggal || !data.jenis || !data.kategori || !data.jumlah || !data.keterangan) {
    SpreadsheetApp.getUi().alert("Mohon lengkapi semua kolom yang wajib diisi.");
    return;
  }

  const lastRow = DBPembukuanSheet.getLastRow() + 1;

  DBPembukuanSheet.getRange(`B${lastRow}:G${lastRow}`).setValues([[
    data.idPembukuan,
    data.tanggal,
    data.jenis,
    data.kategori,
    data.jumlah,
    data.keterangan
  ]]);

  if(data.jenis === 'Pemasukan'){
    DBPembukuanSheet.getRange('D' + lastRow).setBackground("#42f581");
  }else{
    DBPembukuanSheet.getRange('D' + lastRow).setBackground("#f3f573");
  }

  SpreadsheetApp.getUi().alert("Data pembukuan berhasil disimpan.");
  clearFormPembukuan();
}

function updatePembukuan() {
  const data = getFormDataPembukuan();
  const query = data.idPembukuan.toString().trim();

  if (!query) {
    SpreadsheetApp.getUi().alert("Masukkan ID Pembukuan untuk update.");
    return;
  }

  const allData = DBPembukuanSheet.getDataRange().getValues();
  const headerRow = 4;

  for (let i = headerRow; i < allData.length; i++) {
    if (allData[i][1] === query) {
      DBPembukuanSheet.getRange(`B${i+1}:G${i+1}`).setValues([[
        data.idPembukuan,
        data.tanggal,
        data.jenis,
        data.kategori,
        data.jumlah,
        data.keterangan
      ]]);

      SpreadsheetApp.getUi().alert("Data berhasil diupdate.");
      clearFormPembukuan();
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

function deletePembukuan() {
  const query = FormPembukuanSheet.getRange('D4').getValue().toString().trim();

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert("Konfirmasi", `Yakin ingin menghapus data dengan ID Pembukuan: ${query}?`, ui.ButtonSet.YES_NO);
  if (confirm == ui.Button.NO) return;

  const data = DBPembukuanSheet.getDataRange().getValues();
  const headerRow = 4;

  for (let i = headerRow; i < data.length; i++) {
    if (data[i][1] === query) {
      DBPembukuanSheet.deleteRow(i + 1);
      SpreadsheetApp.getUi().alert("Data berhasil dihapus.");
      clearFormPembukuan();
      return;
    }
  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

