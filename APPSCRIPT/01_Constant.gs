const Sheet = SpreadsheetApp.getActiveSpreadsheet();

const FormPembelianSheet = Sheet.getSheetByName('Form_Pembelian');
const FormPenjualanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form_Penjualan');
const FormPartnerSheet = Sheet.getSheetByName('Form_Partner');
const FormSettingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form_Setting');
const FormPembukuanSheet = Sheet.getSheetByName('Form_Pembukuan');

const DashboardSheet = Sheet.getSheetByName('DASHBOARD');
const DbProductSheet = Sheet.getSheetByName('DB_Produk');
const DBPembelianSheet = Sheet.getSheetByName('DB_Pembelian');
const DBStockSheet = Sheet.getSheetByName('DB_Realtime_Stock');
const DBKustomerSheet = Sheet.getSheetByName('DB_Kustomer');
const DBSupplierSheet = Sheet.getSheetByName('DB_Supplier');
const DBPenjualanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_Penjualan');
const DBPembukuanSheet = Sheet.getSheetByName('DB_Pembukuan');

const urlMember = 'https://agusman17.github.io/skincare-membercard/?id=';