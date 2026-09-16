function doGet(e) {
  let page = 'Index';
  if (e && e.parameter && e.parameter.p === 'admin') {
    page = 'Admin';
  }
  
  let template = HtmlService.createTemplateFromFile(page);
  
  if (page === 'Index') {
    let perdaData = getPerdaDataArray();
    template.perdaListJson = JSON.stringify(perdaData);
  }
  
  return template.evaluate()
    .setTitle(page === 'Admin' ? 'Admin - Tambah Perda' : 'Matriks Perda - Satpol PP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet() || ss.getSheetByName("DataPerda") || ss.getSheets()[0];
  return sheet;
}

// Fungsi konversi link Google Drive
function convertDriveLink(url) {
  if (!url) return '';
  url = url.toString().trim();
  
  let match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return "https://drive.google.com/file/d/" + match[1] + "/preview";
  }
  
  let cleanId = url.replace(/^[\/]+/, '');
  if (cleanId.length > 15 && !cleanId.includes(' ')) {
    return "https://drive.google.com/file/d/" + cleanId + "/preview";
  }
  
  return url;
}

// MENGAMBIL data dengan pemetaan kolom baru (Kolom A s.d. H)
function getPerdaDataArray() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol <= 0) return [];
    
    const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 8)).getDisplayValues();
    let perdaList = [];
    
    for (let i = 0; i < rows.length; i++) {
      let row = rows[i];
      let id = String(row[0] || "").trim();
      let nomor = String(row[1] || "").trim();
      let tahun = String(row[2] || "").trim();
      let jenis = String(row[3] || "").trim();
      let namaPerda = String(row[4] || "").trim();
      let tentang = String(row[5] || "").trim();
      let sanksi = String(row[6] || "").trim();
      let link = String(row[7] || "").trim();
      
      if (nomor === "" && tahun === "" && tentang === "" && jenis === "") {
        continue;
      }
      
      perdaList.push({
        ID: id || ("PRD-" + (i + 1)),
        Nomor_Perda: nomor || "-",
        Tahun: tahun || "-",
        Jenis_Perda: jenis || "Peraturan Umum",
        Nama_Perda: namaPerda || "-",
        Tentang: tentang || "-",
        Sanksi_Tindakan: sanksi || "Sesuai ketentuan berlaku",
        Link_Dokumen: convertDriveLink(link)
      });
    }
    
    return perdaList;
  } catch (error) {
    return [];
  }
}

// MENAMBAH data baru dari Panel Admin ke Spreadsheet
function addPerdaData(formData) {
  try {
    const sheet = getSheet();
    const newId = "PRD-" + new Date().getTime();
    
    const rowData = [
      newId,                          // Kolom A: ID
      formData.nomor || "",           // Kolom B: Nomor Perda
      formData.tahun || "",           // Kolom C: Tahun
      formData.jenis || "",           // Kolom D: Kategori Perda
      formData.namaPerda || "",       // Kolom E: Nama Perda
      formData.tentang || "",         // Kolom F: Tentang
      formData.sanksi || "",          // Kolom G: Sanksi
      formData.link || ""             // Kolom H: Link Dokumen
    ];
    
    sheet.appendRow(rowData);
    return JSON.stringify({ status: "success", message: "Data Perda berhasil ditambahkan!" });
  } catch (error) {
    return JSON.stringify({ status: "error", message: error.toString() });
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}