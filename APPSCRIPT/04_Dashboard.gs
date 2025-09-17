function showInputModal() {
  var html = HtmlService.createHtmlOutputFromFile("InputPassword")
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Masukkan Password");
}

/* --- main: called from HTML modal --- */
function checkPassword(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // CHANGE THESE NAMES if your sheet names are different:
  var formSheet = ss.getSheetByName("Form_Setting"); // adjust to your real name
  var dashboardSheet = ss.getSheetByName("DASHBOARD");

  if (!formSheet) throw new Error("FormSetting sheet not found. Update the sheet name in the script.");
  if (!dashboardSheet) throw new Error("Dashboard sheet not found. Update the sheet name in the script.");

  var passwordKu = formSheet.getRange("D8").getValue();

  if (formData.password === passwordKu) {

    dashboardSheet.getRange('C12:C13').setFormula(
      '=SUMIFS(DB_Pembelian[Total];DB_Pembelian[Tanggal];(">"&I3);DB_Pembelian[Tanggal];"<"&L3)'
    );

    dashboardSheet.getRange('F12:F13').setFormula(
      '=SUMIFS(DB_Penjualan[Total];DB_Penjualan[Tanggal];(">"&I3);DB_Penjualan[Tanggal];"<"&L3)'
    );

    dashboardSheet.getRange('I12:I13').setFormula(
      '=ABS(SUMIFS(DB_Pembukuan[JUMLAH];DB_Pembukuan[TANGGAL];">"&I3;DB_Pembukuan[TANGGAL];"<"&L3;DB_Pembukuan[JENIS];"Pengeluaran"))'
    );

    dashboardSheet.getRange('L12:L13').setFormula(
      '=SUMIFS(DB_Pembukuan[JUMLAH];DB_Pembukuan[TANGGAL];">"&I3;DB_Pembukuan[TANGGAL];"<"&L3)'
    );

    dashboardSheet.getRange('C12:C13').setNumberFormat('0,000')
    dashboardSheet.getRange('F12:F13').setNumberFormat('0,000')
    dashboardSheet.getRange('I12:I13').setNumberFormat('0,000')
    dashboardSheet.getRange('L12:L13').setNumberFormat('0,000')

    return { success: true, msg: "Hasil ditampilkan" };

  } else {

    return { success: false, msg: "Wrong password" };
    
  }
}

function formatIDR(n) {
  n = Math.round(Number(n) || 0);
  var s = n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return "Rp. " + s;
}

function showInputModalSetting() {
  var html = HtmlService.createHtmlOutputFromFile("inputPasswordSetting")
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, "Masukkan Password");
}

function checkPasswordSetting(formData){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  // CHANGE THESE NAMES if your sheet names are different:
  var formSheet = ss.getSheetByName("Form_Setting"); // adjust to your real name
  var dashboardSheet = ss.getSheetByName("DASHBOARD");

  if (!formSheet) throw new Error("FormSetting sheet not found. Update the sheet name in the script.");
  if (!dashboardSheet) throw new Error("Dashboard sheet not found. Update the sheet name in the script.");

  var passwordKu = formSheet.getRange("D8").getValue();

  if (formData.password === passwordKu) {
    changeSheetSetting();
    return { success: true, msg: "Hasil ditampilkan" };
  }
}

/* --- helpers for debugging (run from editor or call via google.script.run) --- */
function listTriggers() {
  var ts = ScriptApp.getProjectTriggers();
  var out = ts.map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      eventType: t.getEventType ? String(t.getEventType()) : "unknown",
      source: t.getTriggerSource ? String(t.getTriggerSource()) : "unknown"
    };
  });
  Logger.log(JSON.stringify(out));
  return out;
}

function deleteAllFreezeTriggers() {
  var ts = ScriptApp.getProjectTriggers();
  ts.forEach(function(t) {
    if (t.getHandlerFunction && t.getHandlerFunction() === "freezeFormula") {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { deleted: true };
}


function setMonthDates() {

  DashboardSheet.getRange('C12:C13').setValue('xxx')
  DashboardSheet.getRange('F12:F13').setValue('xxx')
  DashboardSheet.getRange('I12:I13').setValue('xxx')
  DashboardSheet.getRange('L12:L13').setValue('xxx')

  // Get today's date
  var today = new Date();
  
  // First day of current month
  var firstDay = new Date(today.getFullYear(), today.getMonth(), 1);
  
  // Last day of current month
  var lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  
  // Set the values dasboard
  DashboardSheet.getRange("I3").setValue(firstDay);
  DashboardSheet.getRange("L3").setValue(lastDay);

  // Set the values db_pembelian
  DBPembelianSheet.getRange("F2").setValue(firstDay);
  DBPembelianSheet.getRange("G2").setValue(lastDay);

  // Set the values db_penjualan
  DBPenjualanSheet.getRange("F2").setValue(firstDay);
  DBPenjualanSheet.getRange("G2").setValue(lastDay);

  // Set the values db_realtimestock
  DBStockSheet.getRange("E2").setValue(firstDay);
  DBStockSheet.getRange("F2").setValue(lastDay);

  // Set the values db_pembukuan
  DBPembukuanSheet.getRange("E2").setValue(firstDay);
  DBPembukuanSheet.getRange("F2").setValue(lastDay);
  
  // Format as dd/mm/yyyy
  DashboardSheet.getRange("I3:L3").setNumberFormat("dd/MM/yyyy");
  DBPembelianSheet.getRange("F2:G2").setNumberFormat("dd/MM/yyyy");
  DBPenjualanSheet.getRange("F2:G2").setNumberFormat("dd/MM/yyyy");
  DBStockSheet.getRange("E2:F2").setNumberFormat("dd/MM/yyyy");
  DBPembukuanSheet.getRange("E2:F2").setNumberFormat("dd/MM/yyyy");

}
