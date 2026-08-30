const FOLDER_ID = "1HcGM5JYvQ-aXUy8pVcbwkOMM8AzPpW0H";
const FOLDER_NOTA_ID = "1xbG-HsX0gbz-TRnOyzFLb3GWdnyE3g59"; 
const KATA_SANDI_ADMIN = "MUZAKIR@15";

// MASUKKAN ID SPREADSHEET BARU KAMU DI SINI
const SPREADSHEET_ID = "103D-ycz2SnFjkWzu_LY18V_O_CaB6GY_6_lxJsR32As";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Data request kosong"})).setMimeType(ContentService.MimeType.JSON);
    }
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if (action === "cek_login") return cekLogin(requestData);
    if (action === "start_session") return startSession(ss, requestData);
    if (action === "end_session") return endSession(ss, requestData);
    if (action === "add_expense") return addExpense(ss, requestData);
    if (action === "get_dashboard") return getDashboardData(ss);
    if (action === "buat_nota") return buatNotaPDF(ss, requestData);
    if (action === "update_jenis_klien_proyek") return updateJenisKlienProyek(ss, requestData);
    if (action === "simpan_aset") return simpanAset(ss, requestData); 
    if (action === "get_payroll_data") return getPayrollData(ss);
    if (action === "simpan_kasbon") return simpanKasbon(ss, requestData);
    if (action === "proses_gaji") return prosesGaji(ss, requestData);
    if (action === "bayar_invoice") return bayarInvoice(ss, requestData);

    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Aksi tidak dikenali"})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// FUNGSI AUTO-BUILDER SHEET (ANTI-HILANG)
// ==========================================
function cekDanBuatSheetAset(ss) {
  let sheet = ss.getSheetByName("Data Aset");
  if (!sheet) {
    sheet = ss.insertSheet("Data Aset");
    sheet.appendRow(["ID ASET", "NAMA PANGGUNG", "JENIS ARMADA", "HARGA BELI", "MASA MANFAAT (TAHUN)", "HARGA SEWA / JAM", "STATUS", "TERAKHIR OPERASI"]);
    sheet.getRange("D:D").setNumberFormat('Rp #,##0');
    sheet.getRange("F:F").setNumberFormat('Rp #,##0');
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#facc15");
  }
  return sheet;
}

function cekDanBuatSheetSesi(ss) {
  let sheet = ss.getSheetByName("Data Petugas Lapangan");
  if (!sheet) {
    sheet = ss.insertSheet("Data Petugas Lapangan");
    sheet.appendRow(["ID ESK", "TGL", "NAMA PETUGAS", "LOKASI", "JAM MULAI", "JAM SELESAI", "TOTAL JAM", "HARGA", "SUBTOTAL", "FOTO MULAI", "FOTO SELESAI", "STATUS", "ARMADA", "JENIS KLIEN"]);
    sheet.getRange("A1:N1").setFontWeight("bold").setBackground("#facc15");
  }
  return sheet;
}

function cekDanBuatSheetPengeluaran(ss) {
  let sheet = ss.getSheetByName("Pengeluaran");
  if (!sheet) {
    sheet = ss.insertSheet("Pengeluaran");
    sheet.appendRow(["ID PENGELUARAN", "ID PETUGAS", "TGL", "NAMA PETUGAS", "URAIAN OP", "ANGGARAN", "LINK BUKTI", "KETERANGAN"]);
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#facc15");
  }
  return sheet;
}

function cekDanBuatSheetINV(ss) {
  let sheet = ss.getSheetByName("INV");
  if (!sheet) {
    sheet = ss.insertSheet("INV");
    sheet.appendRow(["NO INVOICE", "TANGGAL TERBIT", "KLIEN", "GRAND TOTAL", "BANK", "REKENING", "ATAS NAMA", "LINK PDF", "KODE UNIK", "STATUS BAYAR", "LOKASI PROYEK", "TOTAL TERBAYAR", "SISA TAGIHAN", "RIWAYAT BAYAR"]);
    sheet.getRange("A1:N1").setFontWeight("bold").setBackground("#facc15");
  }
  return sheet;
}

// ==========================================
// FUNGSI UTAMA TRACKER & DASHBOARD
// ==========================================
function simpanAset(ss, data) {
  const sheet = cekDanBuatSheetAset(ss);
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  if (data.idAset) {
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.idAset) { rowIndex = i + 1; break; }
    }
  }
  const idAsetFix = data.idAset || ("AST-" + new Date().getTime());
  
  if (rowIndex > -1) { 
    sheet.getRange(rowIndex, 2).setValue(data.namaPanggung);
    sheet.getRange(rowIndex, 3).setValue(data.jenis);
    sheet.getRange(rowIndex, 4).setValue(data.hargaBeli || 0);
    sheet.getRange(rowIndex, 5).setValue(data.masaManfaat || 0);
    sheet.getRange(rowIndex, 6).setValue(data.hargaSewa || 0);
  } else { 
    sheet.appendRow([idAsetFix, data.namaPanggung, data.jenis, data.hargaBeli || 0, data.masaManfaat || 0, data.hargaSewa || 0, "Nganggur", "-"]);
  }
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Data Aset Berhasil Disimpan!"})).setMimeType(ContentService.MimeType.JSON);
}

function cekLogin(data) {
  if (data.password === KATA_SANDI_ADMIN) {
    return ContentService.createTextOutput(JSON.stringify({status: "success"})).setMimeType(ContentService.MimeType.JSON);
  } else {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Kunci rahasia salah!"})).setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper to safely format time values
function formatTimeWithPeriod(val, nextRow, colMulai, colSelesai, isStart) {
  let formula = '=IF(F' + nextRow + '=""; ""; (IF(INT(F' + nextRow + ')=0; B' + nextRow + '+F' + nextRow + '+IF(MOD(F' + nextRow + ';1)<MOD(E' + nextRow + ';1);1;0); F' + nextRow + ') - IF(INT(E' + nextRow + ')=0; B' + nextRow + '+E' + nextRow + '; E' + nextRow + ')) * 24)';
  return formula;
}

function startSession(ss, data) {
  const sheet = cekDanBuatSheetSesi(ss); 
  let tarif = data.armada === "Tronton" ? 0 : 600000; 
  let namaPanggungAset = data.armada;
  if (data.idAset) {
    const sheetAset = cekDanBuatSheetAset(ss);
    const dataAsetRaw = sheetAset.getDataRange().getValues();
    for (let i = 1; i < dataAsetRaw.length; i++) {
      if (dataAsetRaw[i][0] === data.idAset) {
        namaPanggungAset = dataAsetRaw[i][1]; tarif = dataAsetRaw[i][5] || 0; sheetAset.getRange(i + 1, 7).setValue("Sedang Beroperasi"); break;
      }
    }
  }
  let imageUrl = data.image ? uploadToDrive(data.image, "mulai_" + data.sessionId + ".jpg", FOLDER_ID) : "";
  let nextRow = sheet.getLastRow() + 1;
  let jamMulaiFix = data.jamMulai ? String(data.jamMulai).replace(/\./g, ':') : "";
  let rumusSubtotal = data.armada === "Tronton" ? 0 : '=IF(F' + nextRow + '=""; ""; G' + nextRow + '*H' + nextRow + ')';
  
  sheet.appendRow([
    data.sessionId, data.tanggal, data.petugas, data.lokasi || "", jamMulaiFix, "", 
    '=IF(F' + nextRow + '=""; ""; (IF(INT(F' + nextRow + ')=0; B' + nextRow + '+F' + nextRow + '+IF(MOD(F' + nextRow + ';1)<MOD(E' + nextRow + ';1);1;0); F' + nextRow + ') - IF(INT(E' + nextRow + ')=0; B' + nextRow + '+E' + nextRow + '; E' + nextRow + ')) * 24)', 
    tarif, rumusSubtotal, imageUrl, "", "Sedang Beroperasi", namaPanggungAset, data.jenisKlien || "Perorangan"
  ]);

  // PAKSA FORMAT ANGKA & RUPIAH UNTUK MENCEGAH BUG LINTAS-HARI
  sheet.getRange(nextRow, 7).setNumberFormat("0.00");
  sheet.getRange(nextRow, 8, 1, 2).setNumberFormat('Rp #,##0');

  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Sesi kerja berhasil dimulai!"})).setMimeType(ContentService.MimeType.JSON);
}

function updateJenisKlienProyek(ss, data) {
  const sheet = cekDanBuatSheetSesi(ss);
  const values = sheet.getDataRange().getValues();
  let changed = false;
  
  // Update massal kolom JENIS KLIEN (index 13) berdasarkan LOKASI (index 3)
  for (let i = 1; i < values.length; i++) {
    let lokasiSheet = (values[i][3] || "").toString().trim();
    if (lokasiSheet === data.lokasi) {
      sheet.getRange(i + 1, 14).setValue(data.jenisKlien);
      changed = true;
    }
  }
  
  if (!changed) throw new Error("Proyek tidak ditemukan!");
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Status proyek " + data.lokasi + " diubah ke " + data.jenisKlien + "!"})).setMimeType(ContentService.MimeType.JSON);
}

function endSession(ss, data) {
  const sheet = cekDanBuatSheetSesi(ss);
  const values = sheet.getDataRange().getValues();
  let rowTarget = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.sessionId) { rowTarget = i + 1; break; }
  }
  if (rowTarget === -1) throw new Error("Sesi aktif tidak ditemukan!");
  
  let imageUrl = data.image ? uploadToDrive(data.image, "selesai_" + data.sessionId + ".jpg", FOLDER_ID) : "";
  let jamSelesaiFix = data.jamSelesai ? String(data.jamSelesai).replace(/\./g, ':') : "";
  
  sheet.getRange(rowTarget, 6).setValue(jamSelesaiFix); 
  sheet.getRange(rowTarget, 7).setFormula('=IF(F' + rowTarget + '=""; ""; (IF(INT(F' + rowTarget + ')=0; B' + rowTarget + '+F' + rowTarget + '+IF(MOD(F' + rowTarget + ';1)<MOD(E' + rowTarget + ';1);1;0); F' + rowTarget + ') - IF(INT(E' + rowTarget + ')=0; B' + rowTarget + '+E' + rowTarget + '; E' + rowTarget + ')) * 24)');
  sheet.getRange(rowTarget, 11).setValue(imageUrl); 
  sheet.getRange(rowTarget, 12).setValue("Selesai"); 

  // PAKSA ULANG FORMAT ANGKA & RUPIAH SAAT SELESAI
  sheet.getRange(rowTarget, 7).setNumberFormat("0.00");
  sheet.getRange(rowTarget, 8, 1, 2).setNumberFormat('Rp #,##0');

  if (data.idAset) {
    const sheetAset = cekDanBuatSheetAset(ss);
    const dataAsetRaw = sheetAset.getDataRange().getValues();
    const curDate = new Date().toLocaleDateString('id-ID', {day: 'numeric', month: 'long', year: 'numeric'});
    for (let i = 1; i < dataAsetRaw.length; i++) {
      if (dataAsetRaw[i][0] === data.idAset) {
        sheetAset.getRange(i + 1, 7).setValue("Nganggur"); sheetAset.getRange(i + 1, 8).setValue(curDate); break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Sesi kerja diselesaikan!"})).setMimeType(ContentService.MimeType.JSON);
}

function addExpense(ss, data) {
  const sheet = cekDanBuatSheetPengeluaran(ss);
  let imageUrl = data.image ? uploadToDrive(data.image, "nota_" + data.idPengeluaran + ".jpg", FOLDER_ID) : "";
  sheet.appendRow([data.idPengeluaran, data.sessionId, data.tanggal, data.petugas, data.jenis, data.nominal, imageUrl, data.keterangan || ""]);
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Biaya operasional dicatat!"})).setMimeType(ContentService.MimeType.JSON);
}

function getDashboardData(ss) {
  const tz = ss.getSpreadsheetTimeZone(); 
  
  const parseSheet = (sheetName, createFunc) => {
    let sheet = ss.getSheetByName(sheetName) || createFunc(ss);
    let data = sheet.getDataRange().getDisplayValues();
    if(data.length < 2) return [];
    let headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          if (h.toString().toUpperCase().includes("JAM") || h.toString().toUpperCase().includes("WAKTU")) val = Utilities.formatDate(val, tz, "HH:mm");
          else val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        }
        obj[h] = val;
      });
      return obj;
    });
  };

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", 
    data: { 
      sesi: parseSheet("Data Petugas Lapangan", cekDanBuatSheetSesi), 
      pengeluaran: parseSheet("Pengeluaran", cekDanBuatSheetPengeluaran), 
      aset: parseSheet("Data Aset", cekDanBuatSheetAset),
      inv: parseSheet("INV", cekDanBuatSheetINV),
      gaji: parseSheet("Riwayat Gaji", cekDanBuatSheetGaji),
      pegawai: parseSheet("Database Pegawai", cekDanBuatSheetPegawai),
      kasbon: parseSheet("Kasbon", cekDanBuatSheetKasbon)
    } 
  })).setMimeType(ContentService.MimeType.JSON);
}

function getPayrollData(ss) {
  const tz = ss.getSpreadsheetTimeZone(); 
  const parseSheet = (sheetName, createFunc) => {
    let sheet = ss.getSheetByName(sheetName) || createFunc(ss);
    let data = sheet.getDataRange().getDisplayValues();
    if(data.length < 2) return [];
    let headers = data[0];
    return data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (val instanceof Date) {
          if (h.toString().toUpperCase().includes("JAM") || h.toString().toUpperCase().includes("WAKTU")) val = Utilities.formatDate(val, tz, "HH:mm");
          else val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        }
        obj[h] = val;
      });
      return obj;
    });
  };

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    pegawai: parseSheet("Database Pegawai", cekDanBuatSheetPegawai),
    gaji: parseSheet("Riwayat Gaji", cekDanBuatSheetGaji),
    kasbon: parseSheet("Kasbon", cekDanBuatSheetKasbon),
    sesi: parseSheet("Data Petugas Lapangan", cekDanBuatSheetSesi),
    pengeluaran: parseSheet("Pengeluaran", cekDanBuatSheetPengeluaran),
    aset: parseSheet("Data Aset", cekDanBuatSheetAset) 
  })).setMimeType(ContentService.MimeType.JSON);
}

function uploadToDrive(base64Data, filename, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data.split(",")[1]), "image/jpeg", filename);
  return folder.createFile(blob).getUrl();
}

function angkaKeTerbilang(angka) {
  angka = Math.round(Math.abs(angka)); 
  if (angka === 0) return "Nol";
  function proses(n) {
    const huruf = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"];
    if (n === 0) return "";
    if (n < 12) return " " + huruf[n];
    if (n < 20) return proses(n - 10) + " Belas";
    if (n < 100) return proses(Math.floor(n / 10)) + " Puluh" + proses(n % 10);
    if (n < 200) return " Seratus" + proses(n - 100);
    if (n < 1000) return proses(Math.floor(n / 100)) + " Ratus" + proses(n % 100);
    if (n < 2000) return " Seribu" + proses(n - 1000);
    if (n < 1000000) return proses(Math.floor(n / 1000)) + " Ribu" + proses(n % 1000);
    if (n < 1000000000) return proses(Math.floor(n / 1000000)) + " Juta" + proses(n % 1000000);
    return proses(Math.floor(n / 1000000000)) + " Miliar" + proses(n % 1000000000);
  }
  return proses(angka).trim();
}

// ==========================================
// INVOICE PDF MAKER (SUPPORT INFO DP & SISA)
// ==========================================
function buatNotaPDF(ss, data) {
  const folder = DriveApp.getFolderById(FOLDER_NOTA_ID);
  let tabelHtml = "";
  function escHTML(str) { return String(str || "").replace(/[&<>'"]/g, match => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[match])); }

  data.sesi.forEach((item, index) => {
    tabelHtml += `<tr><td style="text-align:center;">${index + 1}</td><td>${escHTML(item.tanggal)}</td><td>${escHTML(item.lokasi)}</td><td style="text-align:right;">${escHTML(item.hargaSatuan)}</td><td style="text-align:center;"><strong>${escHTML(item.durasi)}</strong></td><td style="text-align:right; font-weight: bold;">${escHTML(item.subtotal)}</td></tr>`;
  });

  let imgTag = "";
  if (data.kopUrl) {
    if (data.kopUrl.indexOf("data:") === 0) {
      imgTag = `<img src="${data.kopUrl}" alt="Kop Surat" style="width: 100%; max-height: 140px; object-fit: contain; margin-bottom: 5px;">`;
    } else {
      try {
        let blob;
        if (data.kopUrl.includes("drive.google.com")) {
          let match = data.kopUrl.match(/[-\w]{25,}/);
          blob = DriveApp.getFileById(match ? match[0] : "").getBlob();
        } else { blob = UrlFetchApp.fetch(data.kopUrl).getBlob(); }
        imgTag = `<img src="data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}" alt="Kop Surat" style="width: 100%; max-height: 140px; object-fit: contain; margin-bottom: 5px;">`;
      } catch(e) { imgTag = `<div style="color:red; text-align:center; border:1px solid red;">Gagal memuat Kop Surat</div>`; }
    }
  }

  const formatRp = (num) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(num);
  const dpp = parseInt(data.totalHarga.replace(/[^0-9]/g, ''), 10) || 0;
  
  const persenPPN = parseFloat(data.persenPPN) || 11;
  const persenPPh = parseFloat(data.persenPPh) || 2;
  const nilaiPPN = data.applyPPN ? dpp * (persenPPN / 100) : 0;
  const nilaiPPh = data.applyPPh ? dpp * (persenPPh / 100) : 0;
  
  const grandTotalRaw = Math.round(dpp + nilaiPPN - nilaiPPh);
  const teksTerbilang = angkaKeTerbilang(grandTotalRaw).trim() + " Rupiah";
  const invoiceNo = "INV/" + new Date().getFullYear() + "/" + ("0" + (new Date().getMonth() + 1)).slice(-2) + "/" + new Date().getTime().toString().slice(-4);
  const curDate = new Date().toLocaleDateString('id-ID', {day: 'numeric', month: 'long', year: 'numeric'});
  const kodeUnik = Math.random().toString(36).substring(2, 8).toUpperCase();
  const kodeDokumen = kodeUnik + " - " + invoiceNo;

  let dpRaw = parseFloat(data.bayarAwal) || 0;
  let sisaRaw = grandTotalRaw - dpRaw;
  if (sisaRaw < 0) sisaRaw = 0;

  let pajakHtml = "";
  if (data.applyPPN) pajakHtml += `<tr class="total-row"><td colspan="5" class="total-label" style="font-weight:normal;">PPN (${persenPPN}%)</td><td class="total-amount" style="font-weight:normal;">${formatRp(nilaiPPN)}</td></tr>`;
  if (data.applyPPh) pajakHtml += `<tr class="total-row"><td colspan="5" class="total-label" style="font-weight:normal; color:#d32f2f;">PPH 23 (${persenPPh}%)</td><td class="total-amount" style="color:#d32f2f; font-weight:normal;">- ${formatRp(nilaiPPh)}</td></tr>`;

  let dpHtml = "";
  if (dpRaw > 0) {
      dpHtml = `
        <tr class="total-row"><td colspan="5" class="total-label" style="font-size: 13px; color: #15803d;">TELAH DIBAYAR (DP / CICILAN AWAL)</td><td class="total-amount" style="font-size: 13px; color: #15803d;">${formatRp(dpRaw)}</td></tr>
        <tr class="total-row"><td colspan="5" class="total-label" style="font-size: 14px; color: #d32f2f;">SISA TAGIHAN / BELUM DIBAYAR</td><td class="total-amount" style="font-size: 16px; color: #d32f2f;">${formatRp(sisaRaw)}</td></tr>
      `;
  }

  const htmlTemplate = `
  <html><head><style>
    @page { size: A4 portrait; margin: 0.5in; background-color: #ffffff; }
    *, *::before, *::after { box-sizing: border-box; }
    body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #111111; line-height: 1.5; font-size: 12px; margin: 0; padding: 0; }
    .top-accent-bar { height: 6px; background: #d32f2f; margin-bottom: 15px; }
    .header { display: table; width: 100%; margin-bottom: 15px; }
    .header-left { display: table-cell; vertical-align: top; width: 60%; }
    .header-right { display: table-cell; vertical-align: top; width: 40%; text-align: right; }
    .title-text { font-size: 26px; font-weight: bold; color: #d32f2f; margin: 0 0 5px 0; }
    .info-grid { display: table; width: 100%; margin-bottom: 15px; border-top: 1px solid #111111; border-bottom: 1px solid #111111; padding: 8px 0; }
    .info-col { display: table-cell; width: 50%; vertical-align: top; }
    .info-col.right { padding-left: 20px; border-left: 1px solid #111111; }
    table.items { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
    table.items th { padding: 10px; font-size: 11px; color: #ffffff; background-color: #111111; border: 1px solid #111111; text-align: center; }
    table.items td { padding: 10px; border: 1px solid #111111; font-size: 12px; }
    table.items tr.total-row td { background-color: #ffffff; padding: 12px 10px; }
    .total-label { text-align: right !important; font-weight: bold; }
    .total-amount { text-align: right !important; font-weight: bold; }
    .terbilang-box { background: #f9f9f9; padding: 10px; border: 1px dashed #111111; margin-bottom: 15px; font-weight: bold; font-style: italic; }
    .ttd-section { display: table; width: 100%; margin-top: 30px; page-break-inside: avoid; }
    .ttd-box { display: table-cell; width: 33%; text-align: center; vertical-align: bottom; }
    .ttd-space { height: 60px; }
  </style></head>
  <body>
    <div class="top-accent-bar"></div>
    ${imgTag}
    <div class="header">
        <div class="header-left"><h1 class="title-text">INVOICE</h1><p style="font-weight:bold;">NO: ${invoiceNo}</p></div>
        <div class="header-right"><p style="font-weight:bold;">TANGGAL TERBIT</p><p>${curDate}</p></div>
    </div>
    <div class="info-grid">
        <div class="info-col"><p style="font-weight:bold;">DITAGIHKAN KEPADA:</p><p>${escHTML(data.klienTujuan)}</p><p>${escHTML(data.alamatProyek)}</p></div>
        <div class="info-col right"><p style="font-weight:bold;">INFORMASI PROYEK:</p><p><strong>Periode:</strong> ${escHTML(data.periodeInfo)}</p><p><strong>Tim/Alat:</strong> ${escHTML(data.operatorInfo)}</p></div>
    </div>
    <table class="items">
        <thead><tr><th>NO</th><th>TANGGAL KERJA</th><th>LOKASI / URAIAN</th><th>TARIF DASAR</th><th>DURASI / VOL</th><th>SUBTOTAL</th></tr></thead>
        <tbody>
            ${tabelHtml}
            <tr class="total-row"><td colspan="5" class="total-label">DASAR PENGENAAN PAJAK (DPP)</td><td class="total-amount">${formatRp(dpp)}</td></tr>
            ${pajakHtml}
            <tr class="total-row"><td colspan="5" class="total-label" style="font-size: 14px;">GRAND TOTAL INVOICE</td><td class="total-amount" style="font-size: 16px;">${formatRp(grandTotalRaw)}</td></tr>
            ${dpHtml}
        </tbody>
    </table>
    <div class="terbilang-box">Terbilang: "${teksTerbilang}"</div>
    <div style="margin-bottom: 20px;">
        <p style="font-weight:bold;">INSTRUKSI PEMBAYARAN:</p>
        <p><strong>Bank:</strong> ${escHTML(data.namaBank)}<br><strong>No. Rekening:</strong> ${escHTML(data.noRekening)}<br><strong>Atas Nama:</strong> ${escHTML(data.atasNamaRek)}</p>
    </div>
    <div class="ttd-section">
        <div class="ttd-box"></div><div class="ttd-box"></div>
        <div class="ttd-box">
            <p style="margin: 0 0 10px 0;">${escHTML(data.lokasiTtd)}, ${curDate}</p><p style="margin: 0;">Hormat Kami,</p>
            <div class="ttd-space"></div>
            <p style="font-weight: bold; text-decoration: underline; margin: 0;">${escHTML(data.namaTtd)}</p>
            <p style="margin: 5px 0 0 0; font-size: 11px;">${escHTML(data.jabatanTtd)}</p>
        </div>
    </div>
  </body></html>`;

  const blobPDF = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF).setName(kodeDokumen + ".pdf");
  const linkPDF = folder.createFile(blobPDF).getUrl();
  const sheetINV = cekDanBuatSheetINV(ss);
  
  // LOGIKA PENCATATAN PEMBAYARAN AWAL DI INVOICE
  let riwayat = [];
  let imageUrl = "";
  if (data.buktiBayarBase64) {
      imageUrl = uploadToDrive(data.buktiBayarBase64, "dp_" + kodeUnik + ".jpg", FOLDER_NOTA_ID);
  }
  if (dpRaw > 0) {
      riwayat.push({ tanggal: data.tglBayarAwal || curDate, nominal: dpRaw, bukti: imageUrl });
  }
  let terbayar = dpRaw;
  let status = sisaRaw <= 0 ? "Lunas" : (terbayar > 0 ? "Cicilan / Sebagian" : "Belum Lunas");

  // Format Array: [NO, TGL, KLIEN, GRAND TOTAL, BANK, REK, ATASNAMA, PDF, KODE, STATUS, LOKASI PROYEK, TERBAYAR, SISA, JSON RIWAYAT]
  sheetINV.appendRow([invoiceNo, curDate, data.klienTujuan, grandTotalRaw, data.namaBank, data.noRekening, data.atasNamaRek, linkPDF, kodeUnik, status, data.lokasiProyekReal, terbayar, sisaRaw, JSON.stringify(riwayat)]);
  
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Invoice berhasil dibuat!", url: linkPDF})).setMimeType(ContentService.MimeType.JSON);
}

function bayarInvoice(ss, data) {
  const sheet = cekDanBuatSheetINV(ss);
  const values = sheet.getDataRange().getValues();
  let rowTarget = -1;
  for(let i=1; i<values.length; i++) {
      if(values[i][0] === data.noInvoice) { rowTarget = i+1; break; }
  }
  if(rowTarget === -1) return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invoice tidak ditemukan di database"})).setMimeType(ContentService.MimeType.JSON);

  let grandTotal = parseFloat(values[rowTarget-1][3]) || 0;
  let terbayarLama = parseFloat(sheet.getRange(rowTarget, 12).getValue()) || 0;
  let sisaLama = sheet.getRange(rowTarget, 13).getValue();
  if (sisaLama === "" || isNaN(parseFloat(sisaLama))) sisaLama = grandTotal; else sisaLama = parseFloat(sisaLama);
  
  let rawRiwayat = sheet.getRange(rowTarget, 14).getValue();
  let riwayat = [];
  try { riwayat = rawRiwayat ? JSON.parse(rawRiwayat) : []; } catch(e) {}

  let nominalBayar = parseFloat(data.nominal);
  let imageUrl = data.image ? uploadToDrive(data.image, "bayar_" + data.noInvoice.replace(/\//g,'_') + "_" + Date.now() + ".jpg", FOLDER_NOTA_ID) : "";

  let terbayarBaru = terbayarLama + nominalBayar;
  let sisaBaru = sisaLama - nominalBayar;
  if(sisaBaru <= 0) sisaBaru = 0;
  let statusBayar = sisaBaru <= 0 ? "Lunas" : "Cicilan / Sebagian";

  riwayat.push({ tanggal: data.tanggal, nominal: nominalBayar, bukti: imageUrl });

  sheet.getRange(rowTarget, 10).setValue(statusBayar); 
  sheet.getRange(rowTarget, 12).setValue(terbayarBaru);
  sheet.getRange(rowTarget, 13).setValue(sisaBaru);
  sheet.getRange(rowTarget, 14).setValue(JSON.stringify(riwayat));

  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Pembayaran berhasil dicatat!"})).setMimeType(ContentService.MimeType.JSON);
}

function simpanKasbon(ss, data) {
  const sheetKasbon = cekDanBuatSheetKasbon(ss);
  const sheetPegawai = cekDanBuatSheetPegawai(ss);
  const idKasbon = "KSB-" + Date.now();
  const tgl = new Date().toLocaleDateString('id-ID');
  sheetKasbon.appendRow([idKasbon, tgl, data.nama, data.jabatan, data.nominal, data.sumberDana, data.keterangan]);
  
  const dataPegawai = sheetPegawai.getDataRange().getValues();
  let ketemu = false;
  for(let i = 1; i < dataPegawai.length; i++) {
    if(dataPegawai[i][1].toString().toLowerCase() === data.nama.toLowerCase()) {
      let sisaLama = parseFloat(String(dataPegawai[i][4]).replace(/[^0-9.-]/g, '')) || 0;
      sheetPegawai.getRange(i+1, 5).setValue(sisaLama + parseFloat(data.nominal));
      ketemu = true; break;
    }
  }
  if(!ketemu) sheetPegawai.appendRow(["PEG-"+Date.now(), data.nama, data.jabatan, "-", data.nominal]);
  return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Kasbon berhasil disimpan!"})).setMimeType(ContentService.MimeType.JSON);
}

// SLIP GAJI PDF MAKER (FULL VERSION)
function prosesGaji(ss, data) {
  const sheetGaji = cekDanBuatSheetGaji(ss);
  const sheetPegawai = cekDanBuatSheetPegawai(ss);
  
  const idGaji = "GAJI-" + Date.now();
  const tgl = new Date().toLocaleDateString('id-ID');
  const terimaBersih = (parseFloat(data.gajiPokok) + parseFloat(data.bonus)) - parseFloat(data.potongan);
  
  let buktiImg = data.buktiImage ? uploadToDrive(data.buktiImage, idGaji + "_tf.jpg", FOLDER_ID) : "-";
  
  // LOGIKA MENGHITUNG SISA KASBON TERBARU
  let sisaBaru = 0;
  const dataPegawai = sheetPegawai.getDataRange().getValues();
  for(let i = 1; i < dataPegawai.length; i++) {
    if(dataPegawai[i][1].toString().toLowerCase() === data.nama.toLowerCase()) {
      let sisaLama = parseFloat(String(dataPegawai[i][4]).replace(/[^0-9.-]/g, '')) || 0;
      sisaBaru = sisaLama - parseFloat(data.potongan || 0);
      if(sisaBaru < 0) sisaBaru = 0;
      sheetPegawai.getRange(i+1, 5).setValue(sisaBaru);
      break;
    }
  }

  // DESAIN PDF KWITANSI HORIZONTAL (LANDSCAPE)
  const formatRp = (num) => new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR' }).format(num);
  const htmlSlip = `
    <html>
    <head>
    <style>
      @page { size: A5 landscape; margin: 30px; }
      body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #111; font-size: 14px; }
      .kwitansi-box { border: 3px double #000; padding: 20px; border-radius: 8px; position: relative; }
      .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 15px; }
      .header h2 { margin: 0; color: #16a34a; font-size: 24px; font-weight: bold; }
      .header h3 { margin: 5px 0 0 0; letter-spacing: 2px; font-size: 16px; }
      
      table.info { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
      table.info td { padding: 5px 0; vertical-align: top; }
      table.info td.label { width: 25%; font-weight: bold; }
      table.info td.colon { width: 2%; text-align: center; }
      table.info td.value { width: 73%; border-bottom: 1px dotted #000; font-weight: bold; font-style: italic; }
      
      table.rincian { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      table.rincian td { padding: 6px 0; }
      table.rincian td.uang { text-align: right; font-weight: bold; }
      
      .total-box { background: #facc15; padding: 8px 15px; border: 2px dashed #000; font-size: 16px; font-weight: bold; display: inline-block; }
      .sisa-kasbon { color: #dc2626; font-weight: bold; font-size: 13px; margin-top: 10px; display: inline-block; border: 1px dashed #dc2626; padding: 6px 12px; background: #fee2e2; }
      
      table.footer { width: 100%; text-align: center; }
      table.footer td { width: 50%; vertical-align: bottom; }
    </style>
    </head>
    <body>
      <div class="kwitansi-box">
         <div class="header">
             <h2>CV KEMBAR GROUP LSM</h2>
             <h3>KWITANSI & SLIP GAJI</h3>
         </div>
         
         <table class="info">
             <tr><td class="label">No. Referensi</td><td class="colon">:</td><td class="value">${idGaji}</td></tr>
             <tr><td class="label">Sudah Terima Dari</td><td class="colon">:</td><td class="value">CV Kembar Group</td></tr>
             <tr><td class="label">Nama Pegawai</td><td class="colon">:</td><td class="value">${data.nama} (${data.jabatan})</td></tr>
             <tr><td class="label">Periode Kerja</td><td class="colon">:</td><td class="value">${data.periodeTeks}</td></tr>
         </table>
         
         <table class="rincian">
            <tr><td>Gaji Pokok / Upah Lapangan</td><td class="uang">${formatRp(data.gajiPokok)}</td></tr>
            <tr><td>Bonus / Tunjangan Tambahan</td><td class="uang">${formatRp(data.bonus)}</td></tr>
            <tr style="border-bottom: 2px solid #000;"><td style="color: red; padding-bottom: 10px;">Potongan Kasbon Bulan Ini</td><td class="uang" style="color: red; padding-bottom: 10px;">- ${formatRp(data.potongan)}</td></tr>
         </table>
         
         <table style="width: 100%;">
           <tr>
             <td style="width: 50%; vertical-align: top;">
                 <div class="total-box">TERIMA BERSIH: ${formatRp(terimaBersih)}</div><br>
                 <div class="sisa-kasbon">Sisa Hutang Kasbon: ${formatRp(sisaBaru)}</div>
             </td>
             <td style="width: 50%;">
                 <table class="footer">
                     <tr>
                         <td>
                             <p style="margin:0 0 60px 0;">Penerima,</p>
                             <p style="margin:0; font-weight:bold; text-decoration:underline;">${data.nama}</p>
                         </td>
                         <td>
                             <p style="margin:0 0 60px 0;">Mengetahui, ${tgl}</p>
                             <p style="margin:0; font-weight:bold; text-decoration:underline;">Direktur Utama</p>
                         </td>
                     </tr>
                 </table>
             </td>
           </tr>
         </table>
      </div>
    </body>
    </html>
  `;
  
  const folderPDF = DriveApp.getFolderById(FOLDER_NOTA_ID); 
  const blobPDF = Utilities.newBlob(htmlSlip, MimeType.HTML).getAs(MimeType.PDF).setName("SLIP_" + data.nama + "_" + idGaji + ".pdf");
  const filePDF = folderPDF.createFile(blobPDF);
  const linkPDF = filePDF.getUrl();

  sheetGaji.appendRow([idGaji, tgl, data.nama, data.jabatan, data.periodeTeks, data.gajiPokok, data.bonus, data.potongan, terimaBersih, data.sumberDana, buktiImg, linkPDF]);

  return ContentService.createTextOutput(JSON.stringify({
    status: "success", 
    message: "Gaji berhasil diproses!",
    urlSlip: linkPDF
  })).setMimeType(ContentService.MimeType.JSON);
}
