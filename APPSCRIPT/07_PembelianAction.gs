function tambahDetailPembelianPreview() {
  const data = ambilDataProdukPembelianForm();

  if (!data.namaProduk || !data.qty || !data.satuan || !data.hargaSatuan) {
    SpreadsheetApp.getUi().alert("Mohon lengkapi semua field produk.");
    return;
  }

  const lastRow = FormPembelianSheet.getLastRow() + 1;

  FormPembelianSheet.getRange(lastRow, 2, 1, 7).setValues([[
    data.namaProduk,
    data.qty,
    data.satuan,
    data.hargaSatuan, "",
    data.totalHarga, ""
  ]]);

  FormPembelianSheet.getRange(lastRow, 9).insertCheckboxes();

  clearFormProductPembelian();

}

function ambilDataProdukPembelianForm() {
  return {
    namaProduk: FormPembelianSheet.getRange('D14').getValue(),  // Product
    qty: Number(FormPembelianSheet.getRange('D16').getValue()),
    satuan: FormPembelianSheet.getRange('H16').getValue(),
    hargaSatuan: Number(FormPembelianSheet.getRange('D18').getValue()),
    totalHarga: Number(FormPembelianSheet.getRange('D20').getValue())
  };
}

function clearFormProductPembelian(){
  FormPembelianSheet.getRange('D14:D20').clearContent(); 
  FormPembelianSheet.getRange('D20').setFormula("=D16*D18");
}

function deleteProductChecked() {
  const lastRow = FormPembelianSheet.getLastRow();

  for (let i = lastRow; i >= 27; i--) { // Mulai dari baris data pertama (row 34) sampai atas
    const checkValue = FormPembelianSheet.getRange(i, 9).getValue(); // Kolom J = 10

    if (checkValue === true) {
      deleteSingleProduct(i, lastRow)
    }

  }
}

function deleteProduct() {
  const lastRow = FormPembelianSheet.getLastRow();

  for (let i = lastRow; i >= 27; i--) {
    deleteSingleProduct(i, lastRow)
  }
}

function deleteSingleProduct(i, lastRow){
  FormPembelianSheet.deleteRow(i);
  var newRow = lastRow+1;
  FormPembelianSheet.insertRowAfter(newRow);

  var rangeToMerge = FormPembelianSheet.getRange(newRow+1, 5, 1, 2); // (row, column, numRows, numColumns) - 5 is column E, 2 is for columns E and F
  var rangeToMerge2 = FormPembelianSheet.getRange(newRow+1, 7, 1, 2);
  var rangeToMerge3 = FormPembelianSheet.getRange(newRow+1, 9, 1, 2);

  rangeToMerge.merge();
  rangeToMerge2.merge();
  rangeToMerge3.merge();
}

function filterPembelian(e){
  // Jika edit bukan di E2 atau F2, stop
  var range = e.range;

  if (!(range.getA1Notation() === "F2" || range.getA1Notation() === "G2")) return;

  // Ambil tanggal awal dan akhir
  var startDate = new Date(DBPembelianSheet.getRange("F2").getValue());
  var endDate = new Date(DBPembelianSheet.getRange("G2").getValue());

  if (!startDate || !endDate) return; // kalau salah satu kosong, stop

  // Ambil semua data di tabel
  var lastRow = DBPembelianSheet.getLastRow();
  var dates = DBPembelianSheet.getRange(5, 3, lastRow - 4, 1).getValues(); // kolom C mulai baris 5

  // Reset semua baris agar terlihat
  DBPembelianSheet.showRows(5, lastRow - 4);

  // Loop dan sembunyikan jika di luar rentang
  for (var i = 0; i < dates.length; i++) {
    var rowDate = new Date(dates[i][0]);
    if (isNaN(rowDate)) continue; // skip kalau bukan tanggal
    if (rowDate < startDate || rowDate > endDate) {
      DBPembelianSheet.hideRows(i + 5); // baris mulai dari 5
    }
  }
}

function getFormDataPembelian() {
  return {
    kodePembelian     : FormPembelianSheet.getRange('D6').getValue(),
    tanggal           : FormPembelianSheet.getRange('D8').getValue(),
    supplier          : FormPembelianSheet.getRange('D10').getValue(),
  };
}

function setFormPembelian(data) {
  FormPembelianSheet.getRange('D6').setValue(data.kode || '');
  FormPembelianSheet.getRange('D8').setValue(data.tanggal || '');
  FormPembelianSheet.getRange('D10').setValue(data.supplier || '');
}

function setFormPembelianDetail(data){
  const lastRow = FormPembelianSheet.getLastRow() + 1;

  FormPembelianSheet.getRange(lastRow, 2, 1, 7).setValues([[
    data.namaProduk,
    data.qty,
    data.satuan,
    data.hargaSatuan, "",
    data.total, ""
  ]]);

  FormPembelianSheet.getRange(lastRow, 9).insertCheckboxes();
}

function searchPembelian() {

  const query = FormPembelianSheet.getRange('D6').getValue().toString().toLowerCase().trim();
  if (!query) {
    SpreadsheetApp.getUi().alert("Isi kode pembelian untuk pencarian.");
    return;
  }

  const allData = DBPembelianSheet.getDataRange().getValues();
  const HeaderRow = 4;
  let found = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    const kode = allData[i][1].toString().toLowerCase().trim();
    const nama = allData[i][2].toString().toLowerCase().trim();

    if (kode === query) {

      if(!found){
        setFormPembelian({
          kode        : allData[i][1],
          tanggal     : allData[i][2],
          supplier    : allData[i][3],
        });
      }

      setFormPembelianDetail({
        namaProduk  : allData[i][4],
        qty         : allData[i][5],
        satuan      : allData[i][6],
        hargaSatuan : allData[i][7],
        total       : allData[i][8]
      });

      found = true;
      // break;
    }
  }
}

function updateProductStockPembelian(queryParam, qtyOld, qty){

  const query = queryParam.toString().toLowerCase().trim();

  // SpreadsheetApp.getUi().alert("qty old "+qtyOld+", qty "+qty);

  if (!query) {
    SpreadsheetApp.getUi().alert("Produk Tidak ditemukan.");
    return;
  }

  const allData     = DbProductSheet.getDataRange().getValues();
  const HeaderRow   = 4;
  let found         = false;

  for (let i = HeaderRow; i < allData.length; i++) {
    let nama      = allData[i][2].toString().toLowerCase().trim();
    let stock     = allData[i][7];
    // SpreadsheetApp.getUi().alert("Kolom Stock : "+stock);
    let nextRow   = i+1;
    let newStock  = stock-qtyOld+qty;
    // newStock      = newStock+qty;

    if (nama.includes(query)) {
      DbProductSheet.getRange('H'+nextRow).setValue(newStock);
      return;
    }
  }
}

function storePembelian(data, row){
  DBPembelianSheet.getRange('B' + row + ':I' + row).setValues([data]);
}

function storeRealtimeStock(data, row){
  DBStockSheet.getRange('B' + row + ':G' + row).setValues([data]);
  DBStockSheet.getRange('G' + row).setBackground("#42f581");
}

function storePembukuan(data, row){
  DBPembukuanSheet.getRange('B' + row + ':G' + row).setValues([data]);
  DBPembukuanSheet.getRange('D' + row).setBackground("#f3f573");
}

function simpanPembelian() {

  data = getFormDataPembelian();

  // Validasi minimal
  if (!data.kodePembelian || !data.tanggal || !data.supplier) {
    SpreadsheetApp.getUi().alert('Mohon lengkapi informasi pembelian terlebih dahulu.');
    return;
  }

  // Ambil baris produk (mulai dari baris 27)
  const startRow = 27;
  const lastRow = FormPembelianSheet.getLastRow();
  const productRange = FormPembelianSheet.getRange(startRow, 2, lastRow - startRow + 1, 7);
  const productData = productRange.getValues();

  let kodePembukuan = generateKodePembukuan();

  let queue = 1;

  productData.forEach((row) => {

    const lastRowPembelian      = DBPembelianSheet.getLastRow()+1;
    const lastRowStockRealtime  = DBStockSheet.getLastRow()+1;
    const lastRowPembukuan      = DBPembukuanSheet.getLastRow()+1;

    const [namaProduk, qty, satuan, hargaSatuan, , totalHarga] = row;

    storePembelian(
      [
        data.kodePembelian,
        data.tanggal,
        data.supplier,
        namaProduk,
        qty,
        satuan,
        hargaSatuan,
        totalHarga,
      ],
      lastRowPembelian
    );

    storeRealtimeStock(
      [
        data.kodePembelian,
        data.tanggal,
        namaProduk,
        qty,
        satuan,
        'in'
      ],
      lastRowStockRealtime
    )

    storePembukuan(
      [
        kodePembukuan.IdPembukuan,
        data.tanggal,
        'Pengeluaran',
        'Pembelian',
        -totalHarga,
        data.kodePembelian
      ],
      lastRowPembukuan
    )

    // Cari Product dan Update Stock
    updateProductStockPembelian(namaProduk, 0, qty);

  });

  SpreadsheetApp.getUi().alert(`data produk berhasil disimpan.`);

  clearFormPembelian()
  deleteProduct()

}

function searchDataPembelian(){
  let data        = getFormDataPembelian();

  if (!data.kodePembelian) {
    SpreadsheetApp.getUi().alert("Kode Pembelian harus diisi.");
    return false;
  }

  const HeaderRow = 4;

  let lastRowPembelian = DBPembelianSheet.getLastRow()+1;

  for (let i = HeaderRow; i <= lastRowPembelian; i++) {

    let kodePembelian = DBPembelianSheet.getRange('B'+i).getValue();
    let namaProduk = DBPembelianSheet.getRange('E'+i).getValue();
    let oldQtyParam = DBPembelianSheet.getRange('F'+i).getValue();

    if (kodePembelian === data.kodePembelian.toString().trim()) {
      return {
        row : i,
        oldQty: oldQtyParam,
        namaProduk: namaProduk
      };
    }
  }
}

function searchRealtimeStock(){
  let data        = getFormDataPembelian();

  if (!data.kodePembelian) {
    SpreadsheetApp.getUi().alert("Kode Pembelian harus diisi.");
    return false;
  }

  const HeaderRow = 4;
  
  let lastRowStockRealtime = DBStockSheet.getLastRow()+1;

  for (let j = HeaderRow; j <= lastRowStockRealtime; j++) {

    let kodePembelianStock = DBStockSheet.getRange('B'+j).getValue();

    if (kodePembelianStock === data.kodePembelian.toString().trim()) {
      return {
        row : j
      }
    }

  }
}

function searchPembukuanPembelian(){
  let data        = getFormDataPembelian();

  if(!data.kodePembelian){
    SpreadsheetApp.getUi().alert("Kode Pembelian harus diisi.");
    return false;
  }

  const HeaderRow = 4;
  
  let lastRowPembukuan = DBPembukuanSheet.getLastRow()+1;

  // SpreadsheetApp.getUi().alert("Last Row Pembukuan : "+lastRowPembukuan);

  for (let j = HeaderRow; j <= lastRowPembukuan; j++) {

    let kodePembelianPembukuan = DBPembukuanSheet.getRange('G'+j).getValue();
    let kodePembukuan = DBPembukuanSheet.getRange('B'+j).getValue();

    if (kodePembelianPembukuan === data.kodePembelian.toString().trim()) {
      return {
        row : j,
        kodePembukuan: kodePembukuan
      }
    }

  }
}

function updatePembelian() {
  deletePembelian(true);
  simpanPembelian();
}

function deletePembelian(isUpdate = false){

  // DELETE DB PEMBELIAN
  let data      = getFormDataPembelian();

  if(!isUpdate){
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("Konfirmasi", "Yakin ingin menghapus data pembelian dengan kode: " + data.kodePembelian + "?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) return;

    if (!data.kodePembelian) {
      SpreadsheetApp.getUi().alert("Kode Pembelian harus diisi untuk update.");
      return;
    }
  }

  const HeaderRow = 4;
  let found       = false;

  let lastRowPembelian = DBPembelianSheet.getLastRow()+1;

  for (let i = HeaderRow; i <= lastRowPembelian; i++) {

    let findPembelian = searchDataPembelian();

    if(findPembelian){
      
      DBPembelianSheet.deleteRow(findPembelian.row)

      let findRealtimeStock = searchRealtimeStock();

      if(findRealtimeStock){

        DBStockSheet.deleteRow(findRealtimeStock.row);

      }
      
      let findPembukuan = searchPembukuanPembelian();

      if(findPembukuan){
        DBPembukuanSheet.deleteRow(findPembukuan.row);
      }

      // UPDATE STOCK.
      updateProductStockPembelian(findPembelian.namaProduk, findPembelian.oldQty, 0)

    }

  }

  if(!isUpdate){
    SpreadsheetApp.getUi().alert("Data berhasil dihapus.");
    clearFormPembelian();
  }
  
}

function clearFormPembelian() {
  FormPembelianSheet.getRange('D6:D18').clearContent(); // Info Pembelian
  setTanggal();
  generateKodePembelian();
  FormPembelianSheet.getRange('D18').setFormula("=D14*D16");
  deleteProduct();
}

function generateKodePembelian() {
  const lastRow = DBPembelianSheet.getLastRow();

  if (lastRow <= 4) {
    FormPembelianSheet.getRange("D6").setValue('PB-0001')
  }

  const kodeTerakhir = DBPembelianSheet.getRange(lastRow, 2).getValue(); // Kolom B = Kode Pembelian

  const nomorTerakhir = parseInt(kodeTerakhir.split('-')[1]) || 0;
  const kodeBaru = 'PB-' + String(nomorTerakhir + 1).padStart(4, '0');

  FormPembelianSheet.getRange("D6").setValue(kodeBaru);
}

function setTanggal(){
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone("Asia/Jakarta");

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, "Asia/Jakarta", 'dd/MM/yy');

  FormPembelianSheet.getRange("D8").setValue(formattedDate);
}

