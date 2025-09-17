function getFormDataPartner() {
  return {
    kategori : FormPartnerSheet.getRange('D4').getValue(),
    idPartner: FormPartnerSheet.getRange('D6').getValue(),
    nama     : FormPartnerSheet.getRange('D8').getValue(),
    area     : FormPartnerSheet.getRange('D10').getValue(),
    alamat   : FormPartnerSheet.getRange('D12').getValue(),
    point    : 0,
    telp     : FormPartnerSheet.getRange('D14').getValue()
  };
}

function setFormDataPartner(data) {
  FormPartnerSheet.getRange('D4').setValue(data.kategori || '');
  FormPartnerSheet.getRange('D6').setValue(data.idPartner || '');
  FormPartnerSheet.getRange('D8').setValue(data.nama || '');
  FormPartnerSheet.getRange('D10').setValue(data.area || '');
  FormPartnerSheet.getRange('D12').setValue(data.alamat || '');
  FormPartnerSheet.getRange('D14').setValue(data.telp || '');
}

function clearFormPartner() {
  FormPartnerSheet.getRange('D4:D14').clearContent();
}

function generateKodePartner(kategoriPartner) {

  // SpreadsheetApp.getUi().alert("Generate Kode Partner : "+kategoriPartner);
  let lastRow;

  if(kategoriPartner === 'Suppliers'){

    lastRow = DBSupplierSheet.getLastRow();

    if (lastRow <= 4) {
      FormPartnerSheet.getRange("D6").setValue('SUPPLIER-0001');
      return;
    }

    const kodeTerakhir = DBSupplierSheet.getRange(lastRow, 2).getValue();
    const nomorTerakhir = parseInt(kodeTerakhir.split('-')[1]) || 0;
    const kodeBaru = 'SUPPLIER-' + String(nomorTerakhir + 1).padStart(4, '0');

    FormPartnerSheet.getRange("D6").setValue(kodeBaru);

  }else{
      lastRow = DBKustomerSheet.getLastRow();

      if (lastRow <= 4) {
        FormPartnerSheet.getRange("D6").setValue('220707001-0001');
        return;
      }

      const kodeTerakhir = DBKustomerSheet.getRange(lastRow, 2).getValue(); // Kolom B = Kode Pembelian

      const nomorTerakhir = parseInt(kodeTerakhir.split('-')[1]) || 0;
      const kodeBaru = '220707001-' + String(nomorTerakhir + 1).padStart(4, '0');

      FormPartnerSheet.getRange("D6").setValue(kodeBaru);
  }
}

// function generate

function generateQRCode(idPartner, row) {
  urlBarcode = urlMember+idPartner;
  const imageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(urlBarcode)}`;
  DBKustomerSheet.getRange('I'+row).setFormula(`=IMAGE("${imageUrl}")`); // put QR image in column B
}

function generateBarcode(idPartner, row) {
  const imageUrl = `https://bwipjs-api.metafloor.com/?bcid=code128&text=${encodeURIComponent(idPartner)}&scale=2&includetext`;
  // Masukkan sebagai formula IMAGE ke sel H16
  DBKustomerSheet.getRange('I'+row).setFormula(`=IMAGE("${imageUrl}")`);
}

function simpanPartner() {
  const data = getFormDataPartner();
  let lastRowPartner;

  if (!data.idPartner || !data.nama || !data.kategori || !data.area) {
    SpreadsheetApp.getUi().alert("Mohon lengkapi semua kolom yang wajib diisi.");
    return;
  }

  if(data.kategori === 'Suppliers'){
    lastRowPartner = DBSupplierSheet.getLastRow()+1;

    DBSupplierSheet.getRange(`B${lastRowPartner}:G${lastRowPartner}`).setValues([[
      data.idPartner,
      data.nama,
      data.kategori,
      data.area,
      data.alamat,
      data.telp,
    ]]);

  }else{
    const lastRowPartner = DBKustomerSheet.getLastRow() + 1;

    DBKustomerSheet.getRange(`B${lastRowPartner}:H${lastRowPartner}`).setValues([[
      data.idPartner,
      data.nama,
      data.kategori,
      data.area,
      data.alamat,
      0,
      data.telp,
    ]]);

    generateQRCode(data.idPartner, lastRowPartner);
    let url = urlMember+data.idPartner;

    const formula = `=HYPERLINK("${url}"; "${data.idPartner}")`;
    DBKustomerSheet.getRange("B"+lastRowPartner).setFormula(formula);
  }

  SpreadsheetApp.getUi().alert("Data partner berhasil disimpan.");
  clearFormPartner();
}

function searchPartner() {
  const query = FormPartnerSheet.getRange('D6').getValue().toString().trim();
  const kategori = FormPartnerSheet.getRange('D4').getValue();

  if (!query || !kategori) {
    SpreadsheetApp.getUi().alert("Pilih Kategori dan Masukkan ID Member untuk mencari.");
    return;
  }

  let headerRow = 4;

  if(kategori === 'Suppliers'){

    const data = DBSupplierSheet.getDataRange().getValues();

    for (let i = headerRow; i < data.length; i++) {
      if (data[i][1] === query) {
        setFormDataPartner({
          idPartner: data[i][1],
          nama: data[i][2],
          kategori: data[i][3],
          area: data[i][4],
          alamat: data[i][5],
          telp: data[i][6]
        });
        return;
      }
    }

  }else{

    const data = DBKustomerSheet.getDataRange().getValues();

    for (let i = headerRow; i < data.length; i++) {
      if (data[i][1] === query) {
        setFormDataPartner({
          idPartner: data[i][1],
          nama: data[i][2],
          kategori: data[i][3],
          area: data[i][4],
          alamat: data[i][5],
          telp: data[i][7]
        });
        return;
      }
    }

  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

function updatePartner() {
  const data = getFormDataPartner();
  const query = data.idPartner.toString().trim();
  const headerRow = 4;

  if (!query || !data.kategori) {
    SpreadsheetApp.getUi().alert("Pilih Kategori dan Masukkan ID Member untuk mencari.");
    return;
  }

  if(data.kategori === 'Suppliers'){

    const allData = DBSupplierSheet.getDataRange().getValues();

    for (let i = headerRow; i < allData.length; i++) {
      if (allData[i][1] === query) {

        rowUpdate = i+1;

        DBSupplierSheet.getRange(`B${rowUpdate}:G${rowUpdate}`).setValues([[
          data.idPartner,
          data.nama,
          data.kategori,
          data.area,
          data.alamat,
          data.telp,
        ]]);

      }
    }

    SpreadsheetApp.getUi().alert("Data berhasil diupdate.");
    clearFormPartner();
    return;

  }else{

    const allData = DBKustomerSheet.getDataRange().getValues();

    for (let i = headerRow; i < allData.length; i++) {
      if (allData[i][1] === query) {

        rowUpdate = i+1;

        const lastPoint = allData[i][6];

        DBKustomerSheet.getRange(`B${rowUpdate}:H${rowUpdate}`).setValues([[
          data.idPartner,
          data.nama,
          data.kategori,
          data.area,
          data.alamat,
          lastPoint,
          data.telp
        ]]);

        generateQRCode(data.idPartner, rowUpdate);
        let url = urlMember+data.idPartner;

        const formula = `=HYPERLINK("${url}"; "${data.idPartner}")`;
        DBKustomerSheet.getRange("B"+rowUpdate).setFormula(formula);

      }
    }

    SpreadsheetApp.getUi().alert("Data berhasil diupdate.");
    clearFormPartner();
    return;

  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

function deletePartner() {
  const kategori = FormPartnerSheet.getRange('D4').getValue().toString().trim();
  const query = FormPartnerSheet.getRange('D6').getValue().toString().trim();

  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert("Konfirmasi", `Yakin ingin menghapus data dengan ID Member: ${query}?`, ui.ButtonSet.YES_NO);
  if (confirm == ui.Button.NO) return;

  let data;
  const headerRow = 4;

  if(kategori === 'Suppliers'){
    data = DBSupplierSheet.getDataRange().getValues();
    for (let i = headerRow; i < data.length; i++) {
      if (data[i][1] === query) {
        DBSupplierSheet.deleteRow(i + 1);
        SpreadsheetApp.getUi().alert("Data berhasil dihapus.");
        clearFormPartner();
        return;
      }
    }
  }else{
    data = DBKustomerSheet.getDataRange().getValues();
    for (let i = headerRow; i < data.length; i++) {
      if (data[i][1] === query) {
        DBKustomerSheet.deleteRow(i + 1);
        SpreadsheetApp.getUi().alert("Data berhasil dihapus.");
        clearFormPartner();
        return;
      }
    }
  }

  SpreadsheetApp.getUi().alert("Data tidak ditemukan.");
}

function doGet(e) {
  const id = e.parameter.id;

  let member = getMemberById(id);

  if (member) {
    return ContentService.createTextOutput(JSON.stringify(member || {}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify(json))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMemberById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_Kustomer");
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const row = data.find(r => r[1] === id);
  if (!row) return null;


  const imageBarcode = urlMember+row[1];

  const sheetProduk = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_Produk");
  const dataProduk = sheetProduk.getDataRange().getValues();
  const produkList = dataProduk.slice(4).map(row => row[2]);
  const isPrint = FormSettingSheet.getRange('D6').getValue();
  const point = row[6];
  var formattedPoint = point.toFixed(0) // remove decimals
                      .replace(/\B(?=(\d{3})+(?!\d))/g, ".");

  return {
    id: row[1],
    nama: row[2],
    kategori: row[3],
    area: row[4],
    alamat: row[5],
    point: formattedPoint,
    telp: row[7],
    barcodeUrl: `https://api.qrserver.com/v1/create-qr-code/?data=${encodeURIComponent(imageBarcode)}&size=100x100`,
    produkList: produkList,
    isPrint: isPrint
  };
}

// Untuk load file HTML terpisah
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


