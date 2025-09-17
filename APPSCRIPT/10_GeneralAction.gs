function onEdit(e){
  var activeSheet = e.source.getActiveSheet();
  const editedRange = e.range;

  if(activeSheet.getName() === "DB_Pembukuan"){
    filterPembukuan(e);
  }else if(activeSheet.getName() === "DB_Realtime_Stock"){
    filterRealtimeStock(e);
  }else if(activeSheet.getName() === "DB_Penjualan"){
    filterPenjualan(e);
  }else if(activeSheet.getName() === "DB_Pembelian"){
    filterPembelian(e);
  }
  
  // Hanya untuk sheet FormPembelian dan kolom Nama Produk (E)
  if (activeSheet.getName() === "Form_Pembelian" && editedRange.getColumn() === 4 && editedRange.getRow() === 14){

    const namaProduk = editedRange.getValue();
    const masterSheet = e.source.getSheetByName("DB_Produk");
    const masterData = masterSheet.getRange(5, 1, masterSheet.getLastRow() - 4, 8).getValues();
    const harga = masterData.find(row => row[2] === namaProduk)?.[3] || 0;

    FormPembelianSheet.getRange('D18').setValue(harga);
  }

  // Hanya untuk sheet FormPenjualan dan kolom Nama Produk (E)
  if (activeSheet.getName() === "Form_Penjualan" && editedRange.getColumn() === 4 && editedRange.getRow() === 14){

    const range = e.range;
    const masterSheet = e.source.getSheetByName("DB_Produk");
    const productName = range.getValue();

    const data = masterSheet.getRange("C5:J" + masterSheet.getLastRow()).getValues();
    
    let harga = 0;
    let found = false;

    for (let row of data) {

      let paket = row[7]; // J = Paket
      let namaProduk = row[0];
      // SpreadsheetApp.getUi().alert("namaProduk : "+namaProduk+"Paket : "+paket);

      // If match Product
      if (productName === namaProduk) {
        activeSheet.getRange("D18").setValue(row[4]); // Use Total column instead of manual calc
        found = true;
        break;
      }

      // If match Paket
      if (productName === paket) {
        harga += row[4]; // Sum of Total
        found = true;
      }
    }

    // SpreadsheetApp.getUi().alert("Harga : "+harga);

    if (found && harga > 0) {
      // Paket case
      activeSheet.getRange("D18").setValue(harga);
    }

    if (!found) {
      // Clear harga if no match
      activeSheet.getRange("D18").clearContent();
    }

    // FormPenjualanSheet.getRange('D16').setValue(harga);
  }

    // Hanya untuk sheet FormPenjualan dan kolom Nama Produk (E)
  if (activeSheet.getName() === "Form_Partner" && editedRange.getColumn() === 4 && editedRange.getRow() === 4){

    const kategoriPartner = editedRange.getValue();
    generateKodePartner(kategoriPartner);
  }

}