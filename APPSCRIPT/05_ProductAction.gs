function filterRealtimeStock(e){
  // Jika edit bukan di E2 atau F2, stop
  var range = e.range;
  if (!(range.getA1Notation() === "E2" || range.getA1Notation() === "F2")) return;

  // Ambil tanggal awal dan akhir
  var startDate = new Date(DBStockSheet.getRange("E2").getValue());
  var endDate = new Date(DBStockSheet.getRange("F2").getValue());

  SpreadsheetApp.getUi().alert("Start Date "+ startDate);
  SpreadsheetApp.getUi().alert("End Date "+ endDate);

  if (!startDate || !endDate) return; // kalau salah satu kosong, stop

  // Ambil semua data di tabel
  var lastRow = DBStockSheet.getLastRow();
  var dates = DBStockSheet.getRange(5, 3, lastRow - 4, 1).getValues(); // kolom C mulai baris 5

  // Reset semua baris agar terlihat
  DBStockSheet.showRows(5, lastRow - 4);

  // Loop dan sembunyikan jika di luar rentang
  for (var i = 0; i < dates.length; i++) {
    var rowDate = new Date(dates[i][0]);
    if (isNaN(rowDate)) continue; // skip kalau bukan tanggal
    if (rowDate < startDate || rowDate > endDate) {
      DBStockSheet.hideRows(i + 5); // baris mulai dari 5
    }
  }
}

function generateKodeProduk() {
  const ProductInputSheet = Sheet.getSheetByName('Form_Produk');
  const DbProductSheet = Sheet.getSheetByName('DB_Produk');

  let JumlahBaris = DbProductSheet.getRange('I3').getValue();
  JumlahBaris += 1;

  const kodeBaru = "PRD-" + String(JumlahBaris).padStart(3, '0');
  ProductInputSheet.getRange("E5").setValue(kodeBaru);
}

function getFormData() {
  const sheet = Sheet.getSheetByName('Form_Produk');
  return {
    kode: sheet.getRange('E5').getValue(),
    nama: sheet.getRange('E7').getValue(),
    hpp: sheet.getRange('E9').getValue(),
    marginPersen: sheet.getRange('E11').getValue(),
    marginHarga: sheet.getRange('G11').getValue(),
    hargaJual: sheet.getRange('E13').getValue(),
    stock: sheet.getRange('E15').getValue(),
    kategori: sheet.getRange('E17').getValue(),
    paket: sheet.getRange('E19').getValue(),
  };
}

function setFormData(data) {
  const sheet = Sheet.getSheetByName('Form_Produk');
  sheet.getRange('E5').setValue(data.kode || '');
  sheet.getRange('E7').setValue(data.nama || '');
  sheet.getRange('E9').setValue(data.hpp || '');
  sheet.getRange('E11').setValue(data.marginPersen || '');
  sheet.getRange('G11').setValue(data.marginHarga || '');
  sheet.getRange('E13').setValue(data.hargaJual || '');
  sheet.getRange('G11').setFormula("=E11/100*E9");
  sheet.getRange('E13').setFormula("=E9+G11");
  sheet.getRange('E15').setValue(data.stock || 0);
  sheet.getRange('E17').setValue(data.kategori || '');
  sheet.getRange('E19').setValue(data.paket || '');
}

function addProduct() {
  const DbProductSheet = Sheet.getSheetByName('DB_Produk');
  const data = getFormData();

  let JumlahBaris = DbProductSheet.getRange('I3').getValue();
  let HeaderRow = 4;
  JumlahBaris += 1;
  const BarisBaru = JumlahBaris + HeaderRow;

  const allData = DbProductSheet.getDataRange().getValues();

  for (let i = HeaderRow; i < allData.length; i++) {
    if (allData[i][1].toString().trim() === data.kode.toString().trim()) {
      SpreadsheetApp.getUi().alert("Data dengan Kode "+data.kode+" telah ada sebelumnya, klik Update untuk memperbarui");
      return
    }
  }

  DbProductSheet.getRange('B' + BarisBaru + ':J' + BarisBaru).setValues(
    [
      [
        data.kode, 
        data.nama, 
        data.hpp, 
        data.marginPersen, 
        data.marginHarga,
        data.hargaJual, 
        data.stock, 
        data.kategori,
        data.paket
      ]
    ]);

  clearProductForm();
}

function updateProduct() {
  const DbProductSheet = Sheet.getSheetByName('DB_Produk');
  const data = getFormData();

  if (!data.kode) {
    SpreadsheetApp.getUi().alert("Kode Produk harus diisi untuk update.");
    return;
  }

  const allData = DbProductSheet.getDataRange().getValues();
  const HeaderRow = 4;
  let found = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    if (allData[i][1].toString().trim() === data.kode.toString().trim()) {
      DbProductSheet.getRange('B' + (i + 1) + ':I' + (i + 1)).setValues(
        [
          [
            data.kode, 
            data.nama, 
            data.hpp, 
            data.marginPersen,
            data.marginHarga,
            data.hargaJual, 
            data.stock, 
            data.kategori,
            data.paket
          ]
        ]);
      SpreadsheetApp.getUi().alert("Produk berhasil diupdate.");
      found = true;
      break;
    }
  }

  if (!found) {
    SpreadsheetApp.getUi().alert("Produk dengan kode tersebut tidak ditemukan.");
  }

  clearProductForm();
}

function deleteProductForm() {
  const DbProductSheet = Sheet.getSheetByName('DB_Produk');
  const data = getFormData();

  if (!data.kode) {
    SpreadsheetApp.getUi().alert("Kode Produk harus diisi untuk menghapus.");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("Konfirmasi", "Yakin ingin menghapus produk dengan kode: " + data.kode + "?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) return;

  const allData = DbProductSheet.getDataRange().getValues();
  const HeaderRow = 4;
  let found = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    if (allData[i][1].toString().trim() === data.kode.toString().trim()) {
      DbProductSheet.deleteRow(i + 1);
      SpreadsheetApp.getUi().alert("Produk berhasil dihapus.");
      found = true;
      break;
    }
  }

  if (!found) {
    SpreadsheetApp.getUi().alert("Produk tidak ditemukan.");
  }

  clearProductForm();
}

function searchProduct() {
  const ProductInputSheet = Sheet.getSheetByName('Form_Produk');
  const DbProductSheet = Sheet.getSheetByName('DB_Produk');

  const query = ProductInputSheet.getRange('E5').getValue().toString().toLowerCase().trim();
  if (!query) {
    SpreadsheetApp.getUi().alert("Isi kode atau nama produk untuk pencarian.");
    return;
  }

  const allData = DbProductSheet.getDataRange().getValues();
  const HeaderRow = 4;
  let found = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    const kode = allData[i][0].toString().toLowerCase().trim();
    const nama = allData[i][1].toString().toLowerCase().trim();

    // SpreadsheetApp.getUi().alert(kode);

    if (kode === query || nama.includes(query)) {
      setFormData({
        kode: allData[i][1],
        nama: allData[i][2],
        hpp: allData[i][3],
        marginPersen: allData[i][4],
        marginHarga: allData[i][5],
        hargaJual: allData[i][6],
        stock: allData[i][7],
        kategori: allData[i][8],
        paket: allData[i][9]
      });
      found = true;
      break;
    }
  }

  if (!found) {
    SpreadsheetApp.getUi().alert("Produk tidak ditemukan.");
  }
}

function clearProductForm() {
  const ProductInputSheet = Sheet.getSheetByName('Form_Produk');

  const sheet = Sheet.getSheetByName('Form_Produk');

  sheet.getRange('E5').clearContent();
  sheet.getRange('E7').clearContent();
  sheet.getRange('E9').clearContent();
  sheet.getRange('E11').clearContent().setNumberFormat('#,##0');
  sheet.getRange('G11').clearContent().setNumberFormat('#,##0');
  sheet.getRange('E13').clearContent().setNumberFormat('#,##0');
  sheet.getRange('E15').clearContent().setNumberFormat('#,##0');
  sheet.getRange('E17').clearContent();
  sheet.getRange('E19').clearContent();

  ProductInputSheet.getRange('G11').setFormula("=E11/100*E9");
  ProductInputSheet.getRange('E13').setFormula("=E9+G11");

  generateKodeProduk();
}
