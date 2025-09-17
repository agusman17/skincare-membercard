function onOpen(e) {
  // SpreadsheetApp.getUi().alert("OPEN SCRIPT");
  setMonthDates();
  clearProductForm();
  clearFormPenjualan();
  setTanggal();
  generateKodePembelian();
  generateKodePembukuan();
  setProductDropdownPenjualan();
}