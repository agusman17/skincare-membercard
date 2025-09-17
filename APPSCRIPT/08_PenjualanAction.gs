function tambahDetailPenjualanPreview() {
  const tableSheet = FormPenjualanSheet; // tabel ada di sheet yg sama

  const productName = FormPenjualanSheet.getRange("D14").getValue();
  const qty = FormPenjualanSheet.getRange("D16").getValue() || 1;
  const satuan = FormPenjualanSheet.getRange("H16").getValue() || "PCS";
  const discPercent = FormPenjualanSheet.getRange("D20").getValue() || 0;
  let discPrice = FormPenjualanSheet.getRange("F20").getValue() || 0;
  const total = FormPenjualanSheet.getRange("D22").getValue() || 1;

  const result = getHargaSatuan(productName);
  // SpreadsheetApp.getUi().alert("Result type : "+result.type);

  if (!result) {
    SpreadsheetApp.getUi().alert("Produk atau Paket tidak ditemukan!");
    return;
  }

  let startRow = 26; // baris tabel pembelian
  let lastRow = tableSheet.getLastRow();
  if (lastRow < startRow) lastRow = startRow;

  if (result.type === "product") {
    // Tambah 1 baris produk
    tableSheet.getRange(lastRow + 1, 2, 1, 9).setValues([
      [productName, qty, satuan, result.harga, "", discPercent, discPrice, total, "-"]
    ]);

    tableSheet.getRange(lastRow + 1, 11).insertCheckboxes();

  } else if (result.type === "paket") {
    // Tambah semua produk dalam paket
    result.items.forEach(item => {
      lastRow++;
      qtyHarga = qty * item.harga;
      discPrice = discPercent/100*qtyHarga;
      tableSheet.getRange(lastRow, 2, 1, 9).setValues([
        [item.nama, qty, satuan, item.harga, "", discPercent, discPrice, qtyHarga - discPrice, productName]
      ]);

      tableSheet.getRange(lastRow, 11).insertCheckboxes();

    });
  }

  clearFormProductPenjualan();

}

function deletePenjualanProductChecked(){
  const lastRow = FormPenjualanSheet.getLastRow();

  for (let i = lastRow; i >= 27; i--) { // Mulai dari baris data pertama (row 34) sampai atas
    const checkValue = FormPenjualanSheet.getRange(i, 11).getValue(); // Kolom J = 10

    if (checkValue === true) {
      deleteSingleProductPenjualan(i, lastRow)
    }

  }
}

function deleteProductPenjualan() {
  const lastRow = FormPenjualanSheet.getLastRow();

  for (let i = lastRow; i >= 27; i--) {
    deleteSingleProductPenjualan(i, lastRow)
  }
}

function deleteSingleProductPenjualan(i, lastRow){
  FormPenjualanSheet.deleteRow(i);
  var newRow = lastRow+1;
  FormPenjualanSheet.insertRowAfter(newRow);

  var rangeToMerge = FormPenjualanSheet.getRange(newRow+1, 5, 1, 2);

  rangeToMerge.merge();
}

function clearFormProductPenjualan(){
  FormPenjualanSheet.getRange('D14:D22').clearContent(); 
  FormPenjualanSheet.getRange('F18').setValue(0);
  FormPenjualanSheet.getRange('D22').setFormula("=(D16*D18)-F20");
}

function getHargaSatuan(nama) {
  const data = DbProductSheet.getRange("C5:J" + DbProductSheet.getLastRow()).getValues();

  let harga = 0;
  let items = [];

  for (let row of data) {
    let namaProduk = row[0]; // C = Nama Produk
    let hargaProduk = row[4]; // E+F = HPP + Margin (atau langsung pakai G kalau ada Total)
    let paket = row[7]; // J = Paket

    // Jika pilih produk langsung
    if (nama === namaProduk) {
      return { type: "product", harga: hargaProduk, items: [{ nama: namaProduk, harga: hargaProduk }] };
    }

    // Jika pilih paket
    if (nama === paket) {
      items.push({ nama: namaProduk, harga: hargaProduk });
      harga += hargaProduk;
    }
  }

  if (items.length > 0) {
    return { type: "paket", harga: harga, items: items };
  }

  return null;
}


function setProductDropdownPenjualan() {
  // Ambil data Nama Produk dan Paket dari DB_PRODUK
  const lastRow = DbProductSheet.getLastRow();
  const namaProduk = DbProductSheet.getRange("C5:C" + lastRow).getValues().flat().filter(String);
  const paketProduk = DbProductSheet.getRange("J5:J" + lastRow).getValues().flat().filter(String);

  // Gabungkan keduanya tanpa duplikat
  const produkList = [...new Set([...namaProduk, ...paketProduk])];

  // Set dropdown ke cell Product di Formulir (misalnya di D14)
  const cellProduct = FormPenjualanSheet.getRange("D14");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(produkList, true)
    .setAllowInvalid(false)
    .build();

  cellProduct.setDataValidation(rule);
}


function filterPenjualan(e){
  // Jika edit bukan di E2 atau F2, stop
  var range = e.range;

  if (!(range.getA1Notation() === "F2" || range.getA1Notation() === "G2")) return;

  // Ambil tanggal awal dan akhir
  var startDate = new Date(DBPenjualanSheet.getRange("F2").getValue());
  var endDate = new Date(DBPenjualanSheet.getRange("G2").getValue());

  if (!startDate || !endDate) return; // kalau salah satu kosong, stop

  // Ambil semua data di tabel
  var lastRow = DBPenjualanSheet.getLastRow();
  var dates = DBPenjualanSheet.getRange(5, 3, lastRow - 4, 1).getValues(); // kolom C mulai baris 5

  // Reset semua baris agar terlihat
  DBPenjualanSheet.showRows(5, lastRow - 4);

  // Loop dan sembunyikan jika di luar rentang
  for (var i = 0; i < dates.length; i++) {
    var rowDate = new Date(dates[i][0]);
    if (isNaN(rowDate)) continue; // skip kalau bukan tanggal
    if (rowDate < startDate || rowDate > endDate) {
      DBPenjualanSheet.hideRows(i + 5); // baris mulai dari 5
    }
  }
}

function setTanggalPenjualan() {
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yy");
  FormPenjualanSheet.getRange("D8").setValue(formattedDate); // Cell tanggal
}

function generateKodePenjualan() {
  const lastRow = DBPenjualanSheet.getLastRow();
  if (lastRow <= 4) {
    FormPenjualanSheet.getRange("D6").setValue('PJ-0001');
    return;
  }

  const lastCode = DBPenjualanSheet.getRange(lastRow, 2).getValue(); // Kolom B
  const lastNumber = parseInt(lastCode.split('-')[1]) || 0;
  const newCode = 'PJ-' + String(lastNumber + 1).padStart(4, '0');
  FormPenjualanSheet.getRange("D6").setValue(newCode);
}

function setFormPenjualan(data) {
  FormPenjualanSheet.getRange('D6').setValue(data.kodePenjualan || '');
  FormPenjualanSheet.getRange('D8').setValue(data.tanggal || '');
  FormPenjualanSheet.getRange('D10').setValue(data.customer || '');
}

function setFormPenjualanDetail(data){
  const lastRow = FormPenjualanSheet.getLastRow() + 1;

  FormPenjualanSheet.getRange(lastRow, 2, 1, 9).setValues([[
    data.produk,
    data.qty,
    data.satuan,
    data.hargaSatuan, "",
    data.discPercent,
    data.discPrice,
    data.total,
    data.paket
  ]]);

  FormPenjualanSheet.getRange(lastRow, 11).insertCheckboxes();
}

function getFormDataPenjualan() {
  return {
    kodePenjualan: FormPenjualanSheet.getRange('D6').getValue(),
    tanggal: FormPenjualanSheet.getRange('D8').getValue(),
    customer: FormPenjualanSheet.getRange('D10').getValue()
  };
}

function storePenjualan(data, row) {
  DBPenjualanSheet.getRange('B' + row + ':L' + row).setValues([[
    data.kodePenjualan,
    data.tanggal,
    data.customer,
    data.produk,
    data.qty,
    data.satuan,
    data.hargaSatuan,
    data.discPercent, 
    data.discPrice,
    data.total,
    data.paket
  ]]);
}

function storeStockKeluar(data, row) {
  DBStockSheet.getRange('B' + row + ':G' + row).setValues([[
    data.kodePenjualan,
    data.tanggal,
    data.produk,
    data.qty,
    data.satuan,
    'out'
  ]]);

  DBStockSheet.getRange('G' + row).setBackground("#f3f573");
}

function storePembukuanPenjualan(data, row, kodePembukuan){

  DBPembukuanSheet.getRange('B' + row + ':G' + row).setValues(
    [[
      kodePembukuan.IdPembukuan,
      data.tanggal,
      'Pemasukan',
      'Penjualan',
      data.total,
      data.kodePenjualan
    ]]
  );

  DBPembukuanSheet.getRange('D' + row).setBackground("#42f581");

}

function simpanPenjualan() {
  const data = getFormDataPenjualan();
  if (!data.kodePenjualan || !data.tanggal || !data.customer) {
    SpreadsheetApp.getUi().alert('Mohon lengkapi semua informasi.');
    return;
  }

    // Ambil baris produk (mulai dari baris 27)
  const startRow = 27;
  const lastRow = FormPenjualanSheet.getLastRow();
  const productRange = FormPenjualanSheet.getRange(startRow, 2, lastRow - startRow + 1, 9);
  const productData = productRange.getValues();

  let kodePembukuan = generateKodePembukuan();

  productData.forEach((row) => {
      const lastRowPenjualan = DBPenjualanSheet.getLastRow() + 1;
      const lastRowStock = DBStockSheet.getLastRow() + 1;
      const lastRowPembukuan = DBPembukuanSheet.getLastRow()+1;

      const [produk, qty, satuan, hargaSatuan, , discPercent, discPrice, total, paket] = row;

      // SpreadsheetApp.getUi().alert("Kode Penjualan : "+data.kodePenjualan+", Tanggal : "+data.tanggal+", Customer : "+data.customer+", Produk : "+produk+", QTY : "+qty+", Satuan : "+satuan+", Harga Satuan : "+hargaSatuan+", Disc. Percent : "+discPercent+", Disc. Price : "+discPrice+", Total : "+total+", Paket : "+paket);

      let kodePembukuan = generateKodePembukuan();

      let dataStore = [];
      dataStore = {
        kodePenjualan: data.kodePenjualan,
        tanggal: data.tanggal,
        customer: data.customer,
        produk: produk,
        qty: qty,
        satuan: satuan,
        hargaSatuan: hargaSatuan,
        discPercent: discPercent,
        discPrice: discPrice,
        total: total,
        paket: paket
      }

      storePenjualan(dataStore, lastRowPenjualan);
      storeStockKeluar(dataStore, lastRowStock);
      storePembukuanPenjualan(dataStore, lastRowPembukuan, kodePembukuan);

      updateProductStockPenjualan(produk, 0, qty);
      // updatePointReward(dataStore.customer, total);

  });

  SpreadsheetApp.getUi().alert('Data penjualan berhasil disimpan.');

  clearFormPenjualan();

}

function searchPenjualan() {

  const query = FormPenjualanSheet.getRange('D6').getValue().toString().toLowerCase().trim();
  if (!query) {
    SpreadsheetApp.getUi().alert("Isi kode atau nama produk untuk pencarian.");
    return;
  }

  const allData = DBPenjualanSheet.getDataRange().getValues();
  const HeaderRow = 4;
  let found = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    const kode = allData[i][1].toString().toLowerCase().trim();

    if (kode === query) {

      if(!found){
        setFormPenjualan({
          kodePenjualan : allData[i][1],
          tanggal       : allData[i][2],
          customer      : allData[i][3],
        });
      }

      // SpreadsheetApp.getUi().alert("Produk : "+allData[i][4]);

      setFormPenjualanDetail({
        produk        : allData[i][4],
        qty           : allData[i][5],
        satuan        : allData[i][6],
        hargaSatuan   : allData[i][7],
        discPercent   : allData[i][8],
        discPrice     : allData[i][9],
        total         : allData[i][10],
        paket         : allData[i][11]
      })

      found = true;

    }
  }
}

function findDataPenjualan() {
  const kode = FormPenjualanSheet.getRange("D6").getValue().toString().trim();

  const lastRow = DBPenjualanSheet.getLastRow();
  for (let i = 5; i <= lastRow; i++) {

    const kodeDB = DBPenjualanSheet.getRange('B' + i).getValue().toString().trim();
    const produk = DBPenjualanSheet.getRange('E'+i).getValue();
    const qtyLama = DBPenjualanSheet.getRange('F' + i).getValue();

    // SpreadsheetApp.getUi().alert("kodeDB : "+kodeDB+", kode : "+kode);

    if (kodeDB === kode) {
      return { row: i, oldQty: qtyLama, produk:  produk};
    }
  }
  return null;
}

function searchRealtimeStockPenjualan(){
  let data        = getFormDataPenjualan();

  if (!data.kodePenjualan) {
    SpreadsheetApp.getUi().alert("Kode Penjualan harus diisi.");
    return false;
  }

  const HeaderRow = 4;
  
  let lastRowStockRealtime = DBStockSheet.getLastRow()+1;

  for (let j = HeaderRow; j <= lastRowStockRealtime; j++) {

    let kodePenjualanStock = DBStockSheet.getRange('B'+j).getValue();

    if (kodePenjualanStock === data.kodePenjualan.toString().trim()) {
      return {
        row : j
      }
    }

  }

}

function searchPembukuanPenjualan(){
  let data        = getFormDataPenjualan();

  if(!data.kodePenjualan){
    SpreadsheetApp.getUi().alert("Kode Penjualan harus diisi.");
    return false;
  }

  const HeaderRow = 4;
  
  let lastRowPembukuan = DBPembukuanSheet.getLastRow()+1;

  // SpreadsheetApp.getUi().alert("Last Row Pembukuan : "+lastRowPembukuan);

  for (let j = HeaderRow; j <= lastRowPembukuan; j++) {

    let kodePenjualanPembukuan = DBPembukuanSheet.getRange('G'+j).getValue();
    let kodePembukuan = DBPembukuanSheet.getRange('B'+j).getValue();

    if (kodePenjualanPembukuan === data.kodePenjualan.toString().trim()) {
      return {
        row : j,
        kodePembukuan: kodePembukuan
      }
    }

  }
}

function updatePenjualan() {
  deletePenjualan(true);
  simpanPenjualan();
}

function deletePenjualan(isUpdate = false) {
  const data = getFormDataPenjualan();

  if(!isUpdate){
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Konfirmasi", `Hapus data penjualan: ${data.kodePenjualan}?`, ui.ButtonSet.YES_NO);
    if (response === ui.Button.NO) return;
  }

  const HeaderRow = 4;
  let found       = false;

  let lastRowPenjualan = DBPenjualanSheet.getLastRow()+1;

  for (let i = HeaderRow; i <= lastRowPenjualan; i++) {

    const findPenjualan = findDataPenjualan();

    if (findPenjualan) {
      // SpreadsheetApp.getUi().alert("findPenjualan : "+findPenjualan.row);
      DBPenjualanSheet.deleteRow(findPenjualan.row);

      let findRealtimeStock = searchRealtimeStockPenjualan();
      if(findRealtimeStock){
        DBStockSheet.deleteRow(findRealtimeStock.row);
      }

      let findPembukuan = searchPembukuanPenjualan();
      if(findPembukuan){
        DBPembukuanSheet.deleteRow(findPembukuan.row);
      }
      updateProductStockPenjualan(findPenjualan.produk, findPenjualan.oldQty, 0);
    }

  }

  if(!isUpdate){
    SpreadsheetApp.getUi().alert("Data berhasil dihapus.");
    clearFormPenjualan();
  }

}

function searchRealtimeStockPenjualan(){
  let data        = getFormDataPenjualan();

  if (!data.kodePenjualan) {
    SpreadsheetApp.getUi().alert("Kode Penjualan harus diisi.");
    return false;
  }

  const HeaderRow = 4;
  
  let lastRowStockRealtime = DBStockSheet.getLastRow()+1;

  for (let j = HeaderRow; j <= lastRowStockRealtime; j++) {

    let kodePembelianStock = DBStockSheet.getRange('B'+j).getValue();

    if (kodePembelianStock === data.kodePenjualan.toString().trim()) {
      return {
        row : j
      }
    }

  }

}

function clearFormPenjualan() {
  FormPenjualanSheet.getRange('D6:D10').clearContent();
  generateKodePenjualan();
  setTanggalPenjualan();
  clearFormProductPenjualan();
  deleteProductPenjualan();
}

function updateProductStockPenjualan(namaProduk, qtyOld, qtyNew) {
  const nama = namaProduk.toString().toLowerCase().trim();
  const DBPoductSheet = Sheet.getSheetByName('DB_Produk');
  const data = DBPoductSheet.getDataRange().getValues();

  for (let i = 4; i < data.length; i++) {
    const namaDB = data[i][2].toString().toLowerCase().trim();
    if (namaDB === nama) {
      const stokLama = Number(data[i][7]); // kolom H = stok
      const stokBaru = stokLama + qtyOld - qtyNew;
      DBPoductSheet.getRange(i + 1, 8).setValue(stokBaru);
      break;
    }
  }
}

function updatePointReward(namaPembeli, hargaSatuan){
  const nama = namaPembeli.toString().toLowerCase().trim();
  const DBKustomerSheet = Sheet.getSheetByName('DB_Kustomer');
  const data = DBKustomerSheet.getDataRange().getValues();
  const RpToPoint = FormSettingSheet.getRange('D4').getValue();

  for (let i = 4; i < data.length; i++) {
    const namaDB = data[i][2].toString().toLowerCase().trim();
    if (namaDB === nama) {
      const pointLama = Number(data[i][6]); 
      const point = hargaSatuan/RpToPoint;
      const pointBaru = pointLama + point;
      DBKustomerSheet.getRange(i + 1, 7).setValue(pointBaru);
      break;
    }
  }
}
