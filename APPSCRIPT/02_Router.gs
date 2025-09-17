function changeSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("Sheet '" + sheetName + "' not found.");
    return;
  }

  targetSheet.showSheet();
  ss.setActiveSheet(targetSheet);

  activeSheet.hideSheet();

}

// Specific shortcuts
function changeSheetDashboard()   { changeSheet("DASHBOARD"); }
function changeSheetProduk()      { changeSheet("Form_Produk"); }
function changeDBProduk()         { changeSheet("DB_Produk"); }
function changeSheetPartner()     { changeSheet("Form_Partner"); }
function changeDBSupplier()       { changeSheet("DB_Supplier"); }
function changeDBKustomer()       { changeSheet("DB_Kustomer"); }
function changeDBRealtimeStock()  { changeSheet("DB_Realtime_Stock"); }
function changeSheetPembelian()   { changeSheet("Form_Pembelian"); }
function changeDBPembelian()      { changeSheet("DB_Pembelian"); }
function changeSheetPenjualan()   { changeSheet("Form_Penjualan"); }
function changeDBPenjualan()      { changeSheet("DB_Penjualan"); }
function changeSheetPembukuan()   { changeSheet("Form_Pembukuan"); }
function changeDBPembukuan()      { changeSheet("DB_Pembukuan"); }
function changeSheetSetting()     { changeSheet("Form_Setting"); }
