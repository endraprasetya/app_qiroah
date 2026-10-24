// ====================================================================
// SYSTEM DEFAULT CORNER
// ====================================================================
function doGet(e) {
  // 1. Penanganan Cetak Laporan (Harian & Rekap Bulanan)
  if (e.parameter.page === 'cetak') {
    var jenis = e.parameter.jenis || "HARIAN";
    var templateFile = (jenis === "REKAP") ? 'CetakRekapLaporan' : 'CetakLaporan';
    
    var template = HtmlService.createTemplateFromFile(templateFile);
    template.tglMulai = e.parameter.tglMulai || "";
    template.tglSelesai = e.parameter.tglSelesai || "";
    template.ustadz = e.parameter.ustadz || "";
    template.unit = e.parameter.unit || "";
    template.kategori = e.parameter.kategori || "SEMUA";
    template.filterVal = e.parameter.filterVal || "";

    var judulWindow = (jenis === "REKAP") ? 'Cetak Rekap Laporan Qiroah' : 'Cetak Laporan Qiroah';

    return template.evaluate()
      .setTitle(judulWindow)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  // 2. Penanganan Cetak Dokumen Ujian Kenaikan Jilid & SKL
  if (e.parameter.page === 'cetakUjian') {
    var template = HtmlService.createTemplateFromFile('CetakSKLulus');
    template.idSiswa = e.parameter.idSiswa || "";
    template.unit = e.parameter.unit || "";
    template.tipe = e.parameter.tipe || "SKL";
    template.jilidAsal = e.parameter.jilidAsal || "";
    template.jilidTujuan = e.parameter.jilidTujuan || "";

    return template.evaluate()
      .setTitle('Cetak Dokumen Ujian Qiroah')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  // 3. Penanganan Portal Wali Murid
  if (e.parameter.page === 'wali') {
    return HtmlService.createTemplateFromFile('LoginWali')
      .evaluate()
      .setTitle('Portal Wali Murid - Qiroah')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  
  // 4. Halaman Aplikasi Utama (Index)
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Aplikasi Qiroatul Quran')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ====================================================================
// 1. FUNGSI REGISTRASI GURU BARU (SIGN UP)
// ====================================================================
function prosesDaftarServer(nama, username, password, role, unit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === username.toLowerCase()) {
        return { success: false, message: "Username '" + username + "' sudah digunakan! Coba yang lain." };
      }
    }
    var nextNum = data.length;
    var targetID = "U" + String(nextNum).padStart(3, '0');
    sheet.appendRow([username, password, role, targetID, nama, unit, "BELUM"]);
    return { success: true, message: "Pendaftaran berhasil! Silakan login menggunakan username Anda." };
  } catch(e) {
    return { success: false, message: "Error Server: " + e.toString() };
  }
}

// ====================================================================
// 2. FUNGSI PROSES LOGIN (DENGAN SABUK PENGAMAN ANTI-NULL)
// ====================================================================
function prosesLoginServer(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Error: Sheet 'Users' tidak ditemukan!" };
    
    var data = sheet.getDataRange().getValues();
    var safeUsername = (username || "").toString().trim();
    var safePassword = (password || "").toString().trim();

    for (var i = 1; i < data.length; i++) {
      var dbUser = (data[i][0] || "").toString().trim();
      var dbPass = (data[i][1] || "").toString().trim();
      
      if (dbUser === "") continue; 
      
      if (dbUser === safeUsername && dbPass === safePassword) {
        return {
          success: true,
          username: dbUser,
          role: data[i][2],
          targetID: data[i][3],
          nama: data[i][4],
          unit: data[i][5],
          statusSetup: data[i][6]
        };
      }
    }
    return { success: false, message: "Username atau Password salah!" };
  } catch(e) {
    return { success: false, message: "Error Sistem: " + e.message };
  }
}

// ====================================================================
// 3. FUNGSI AMBIL DAFTAR SISWA UNTUK CHECKLIST SETTING AWAL (YANG HILANG)
// ====================================================================
function getSiswaUntukSetup(unit, kelas, targetID) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Siswa");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var hasil = [];
    
    // Normalisasi input agar pencarian lebih akurat
    var filterUnit = unit.toString().trim().toLowerCase();
    var filterKelas = kelas.toString().trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      // Pastikan kolom D (Index 3) dan G (Index 6) memiliki data
      var sKelas = (data[i][3] || "").toString().trim().toLowerCase();
      var sUnit = (data[i][6] || "").toString().trim().toLowerCase();
      var sPembimbing = (data[i][4] || "").toString().trim();
      
      if (sUnit === filterUnit && sKelas === filterKelas) {
        hasil.push({
          idSiswa: data[i][0].toString(),
          namaSiswa: data[i][1].toString(), // Pastikan kolom B ada datanya
          isChecked: (sPembimbing === targetID)
        });
      }
    }
    return hasil;
  } catch(e) {
    return [];
  }
}

// ====================================================================
// 4. FUNGSI AMBIL DAFTAR SISWA BINAAN AKTIF (VERSI MULTI-USTADZ)
// ====================================================================
function getSiswaBinaanServer(targetID, role) {
  try {
    if (!targetID) {
      Logger.log("getSiswaBinaanServer: TargetID kosong atau tidak valid.");
      return [];
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Siswa");
    
    if (!sheet) {
      Logger.log("getSiswaBinaanServer: Sheet 'Siswa' tidak ditemukan.");
      return [];
    }

    var data = sheet.getDataRange().getValues();
    var target = String(targetID).trim().toUpperCase(); 
    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var sPembimbing = row[4] ? String(row[4]).trim().toUpperCase() : "";
      
      // Pecah ID pembimbing ke dalam array
      var listPembimbing = sPembimbing ? sPembimbing.split(',').map(function(id){ return id.trim(); }) : [];
      
      // 4. Pencocokan Data dalam Array
      if (listPembimbing.indexOf(target) !== -1) {
        hasil.push({
          idSiswa:   String(row[0] || ""), // Kolom A
          namaSiswa: String(row[1] || ""), // Kolom B
          kelas:     String(row[3] || "")  // Kolom D
        });
      }
    }

    return hasil;
  } catch (e) {
    Logger.log("Error pada getSiswaBinaanServer: " + e.toString());
    return [];
  }
}

// ====================================================================
// 5. FUNGSI SIMPAN SETTING BINAAN KELOMPOK SISWA (MULTI-GURU)
// ====================================================================
function simpanSettingSiswaServer(targetID, listIDSiswa, unit, kelas) {
  var lock = LockService.getScriptLock();
  
  if (!lock.tryLock(10000)) {
    return { success: false, message: "Gagal disimpan. Silakan coba beberapa detik lagi." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    
    var pembimbingUpdate = []; 
    var isUpdated = false;
    var targetUpper = targetID.toString().trim().toUpperCase();

    for (var i = 1; i < dataSiswa.length; i++) {
      var sID = (dataSiswa[i][0] || "").toString().trim(); 
      var sKelas = (dataSiswa[i][3] || "").toString().trim(); 
      var sPembimbing = (dataSiswa[i][4] || "").toString().trim().toUpperCase(); 
      var sUnit = (dataSiswa[i][6] || "").toString().trim(); 
      
      // Ubah data pembimbing lama menjadi Array
      var listPembimbingAktif = sPembimbing ? sPembimbing.split(',').map(function(id){ return id.trim(); }) : [];
      var indexTarget = listPembimbingAktif.indexOf(targetUpper);
      var adaDiDatabase = (indexTarget !== -1);
      var diCeklisUser = (listIDSiswa.indexOf(sID) !== -1);
      
      // Proses HANYA JIKA unit dan kelas cocok dengan filter yang sedang dibuka ustadz
      if (sUnit.toLowerCase() === unit.toLowerCase() && sKelas.toLowerCase() === kelas.toLowerCase()) {
        
        if (diCeklisUser && !adaDiDatabase) {
          // Langkah A: Jika siswa di-ceklis tapi ID guru belum ada, maka TAMBAHKAN
          listPembimbingAktif.push(targetUpper);
        } else if (!diCeklisUser && adaDiDatabase) {
          // Langkah B: Jika siswa TIDAK di-ceklis tapi ID guru ada, maka HAPUS ID guru ini saja
          listPembimbingAktif.splice(indexTarget, 1);
        }
      }
      
      // Gabungkan array kembali menjadi string dipisahkan koma dan spasi
      var newPembimbingStr = listPembimbingAktif.join(', ');

      pembimbingUpdate.push([newPembimbingStr]);
      
      if (newPembimbingStr !== sPembimbing) {
        isUpdated = true;
      }
    }
    
    if (isUpdated && pembimbingUpdate.length > 0) {
      sheetSiswa.getRange(2, 5, pembimbingUpdate.length, 1).setValues(pembimbingUpdate);
    }
    
    // Update status di sheet Users 
    var sheetUsers = ss.getSheetByName("Users");
    var dataUsers = sheetUsers.getDataRange().getValues();
    for (var i = 1; i < dataUsers.length; i++) {
      if ((dataUsers[i][3] || "").toString() === targetID) {
        sheetUsers.getRange(i + 1, 7).setValue("SUDAH"); 
        break;
      }
    }
    
    return { success: true, message: "Kelompok binaan berhasil disimpan!" };
  } catch(e) {
    return { success: false, message: "Gagal menyimpan ke server: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ====================================================================
// FUNGSI HAPUS DATA (DUAL DELETE) DI .GS - FIXED & ROBUST
// ====================================================================
function hapusSetoranData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Setoran");
    var db = getFirestoreDB();
    
    var idQiroah = payload.idQiroah ? String(payload.idQiroah).trim() : "";
    var row = parseInt(payload.rowIndex);
    
    // 1. CARI BARIS DI SPREADSHEET JIKA ROWINDEX KOSONG
    if ((isNaN(row) || row < 2) && idQiroah !== "") {
      var dataA = sheet.getRange("A:A").getValues();
      for (var i = dataA.length - 1; i >= 1; i--) {
        if (String(dataA[i][0]).trim() === idQiroah) {
          row = i + 1; // Ketemu barisnya!
          break;
        }
      }
    }
    
    // 2. HAPUS BARIS DI SPREADSHEET
    if (!isNaN(row) && row >= 2) {
      sheet.deleteRow(row);
    }
    
    // 3. HAPUS DOKUMEN DI FIRESTORE
    if (idQiroah !== "") {
      db.deleteDocument("Setoran/" + idQiroah);
    }
    
    return { status: 'success', message: 'Data riwayat berhasil dihapus dari sistem.' };
  } catch(e) {
    return { status: 'error', message: "Gagal menghapus: " + String(e) };
  }
}
// ====================================================================
// FUNGSI SIMPAN SETORAN BARU (DUAL WRITE)
// ====================================================================
function simpanSetoranBaru(payload) {
  try {
    var db = getFirestoreDB(); // Panggil koneksi database
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setoran"); 
    if (!sheet) throw new Error("Sheet 'Setoran' tidak ditemukan!");

    if (!payload.idSiswa || String(payload.idSiswa).trim() === "") {
      throw new Error("Siswa belum dipilih!");
    }
    if (!payload.halaman || String(payload.halaman).trim() === "" || String(payload.halaman).trim() === "-") {
      throw new Error("Halaman / Surat tidak boleh kosong!");
    }

    var idQiroah = payload.idQiroah || "QR" + Math.floor(1000000000 + Math.random() * 9000000000);
    var tanggalSekarang = new Date();

    // 1. SIMPAN KE SPREADSHEET
    sheet.appendRow([
      idQiroah, tanggalSekarang, payload.unit || "-", payload.idSiswa, payload.namaSiswa,
      payload.namaUstadz || "-", payload.jilid, payload.halaman, payload.status,
      payload.catatan || "-", payload.tugasRumah || "-", payload.idAudio || "",
      payload.statusSetoran || "", payload.nilai || "", payload.catatanAudio || "",
      payload.murojaah || "Belum", payload.membacaSekolah || "Belum", payload.murojaahTeman || "Belum",
      payload.nilaiAudio || "" // Kolom S
    ]);

    // 2. SIMPAN KE FIRESTORE (Struktur disamakan)
    var payloadFirestore = {
      tanggal: tanggalSekarang,             
      unit: payload.unit || "-",                
      id_siswa: payload.idSiswa.toString(), 
      nama_siswa: payload.namaSiswa,          
      ustadz: payload.namaUstadz || "-",              
      jilid: payload.jilid,               
      halaman: payload.halaman,             
      status_lancar: payload.status,       
      catatan: payload.catatan || "-",             
      tugas_rumah: payload.tugasRumah || "-",        
      id_audio: payload.idAudio || "",           
      status_setoran: payload.statusSetoran || "",     
      nilai: payload.nilai || "",              
      catatan_audio: payload.catatanAudio || "",      
      murojaah_di_rumah: payload.murojaah || "Belum",  
      membaca_di_sekolah: payload.membacaSekolah || "Belum", 
      murojaah_bersama_teman: payload.murojaahTeman || "Belum", 
      nilai_audio: payload.nilaiAudio || "" 
    };

    db.createDocument("Setoran/" + idQiroah, payloadFirestore);

    return { status: "success", message: "Alhamdulillah, data setoran " + payload.namaSiswa + " berhasil disimpan!" };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

// ====================================================================
// FUNGSI UPDATE SETORAN (OPTIMASI SUPER CEPAT - BATCH UPDATE)
// ====================================================================
function updateSetoran(payload) {
  try {
    var db = getFirestoreDB();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setoran");
    var row = parseInt(payload.rowIndex);
    
    // --- PENCARIAN BARIS ---
    if (isNaN(row) || row < 1) {
       if (!payload.idQiroah) throw new Error("ID Qiroah tidak ditemukan dari form!");
       
       var targetId = String(payload.idQiroah).trim();
       var dataA = sheet.getRange("A:A").getValues();
       var foundRow = -1;
       
       for (var i = dataA.length - 1; i >= 0; i--) {
          if (String(dataA[i][0]).trim() === targetId) {
             foundRow = i + 1;
             break;
          }
       }
       
       if (foundRow === -1) {
          throw new Error("Data ID '" + targetId + "' tidak ditemukan di Kolom A!");
       }
       row = foundRow; 
    }

    // --- 🚀 OPTIMASI 1: BATCH UPDATE SPREADSHEET ---
    // Menggabungkan 8x setValue menjadi hanya 2x setValues untuk memangkas waktu loading
    
    // 1. Update Kolom G (Jilid) sampai K (Tugas Rumah) sekaligus (5 Kolom)
    sheet.getRange(row, 7, 1, 5).setValues([[
      payload.jilid, 
      payload.halaman, 
      payload.status, 
      payload.catatan || "-", 
      payload.tugasRumah || "-"
    ]]);
    
    // 2. Update Kolom P (Murojaah) sampai R (Murojaah Teman) sekaligus (3 Kolom)
    sheet.getRange(row, 16, 1, 3).setValues([[
      payload.murojaah, 
      payload.membacaSekolah, 
      payload.murojaahTeman
    ]]);

    // --- 🚀 OPTIMASI 2: BYPASS GET DOCUMENT FIRESTORE ---
    // Ambil seluruh data baris yang baru diupdate (Kolom A s/d S)
    var rowData = sheet.getRange(row, 1, 1, 19).getValues()[0];
    
    // Langsung tembak ke Firestore tanpa perlu db.getDocument() terlebih dahulu
    if (payload.idQiroah) {
        var docId = String(payload.idQiroah).trim();
        
        var dataUpdate = {
           tanggal: (rowData[1] instanceof Date) ? rowData[1] : new Date(),
           unit: String(rowData[2] || "").trim(),
           id_siswa: String(rowData[3] || "").trim(),
           nama_siswa: String(rowData[4] || "").trim(),
           ustadz: String(rowData[5] || "").trim(),
           jilid: payload.jilid,
           halaman: payload.halaman,
           status_lancar: payload.status,
           catatan: payload.catatan || "-",
           tugas_rumah: payload.tugasRumah || "-",
           id_audio: String(rowData[11] || "").trim(),
           status_setoran: String(rowData[12] || "").trim(),
           nilai: String(rowData[13] || "").trim(),
           catatan_audio: String(rowData[14] || "").trim(),
           murojaah_di_rumah: payload.murojaah,
           membaca_di_sekolah: payload.membacaSekolah,
           murojaah_bersama_teman: payload.murojaahTeman,
           nilai_audio: String(rowData[18] || "").trim()
        };
        
        db.updateDocument("Setoran/" + docId, dataUpdate);
    }

    return { status: "success", message: "Data setoran berhasil diupdate!" };
  } catch (error) {
    return { status: "error", message: "Gagal update data: " + error.toString() };
  }
}

// ====================================================================
// FUNGSI ROUTER: MENENTUKAN INSERT BARU ATAU UPDATE
// ====================================================================
function prosesSetoranData(payload) {
  // 🚀 PERBAIKAN: Gunakan idQiroah sebagai penentu mode Edit, bukan rowIndex
  if (payload.idQiroah && String(payload.idQiroah).trim() !== "") {
    return updateSetoran(payload);
  } else {
    return simpanSetoranBaru(payload);
  }
}

// ====================================================================
// 7. FUNGSI AMBIL RIWAYAT TERAKHIR SISWA (SUPER AMAN ANTI-CRASH)
// ====================================================================
function getRiwayatTerakhirSiswa(idSiswa) {
  try {
    // 1. Sabuk Pengaman Input: Pastikan ID Siswa dari aplikasi tidak null
    if (idSiswa === undefined || idSiswa === null) {
       return { ditemukan: false };
    }
    
    // Gunakan String() alih-alih .toString() agar 100% kebal dari error null
    var idTarget = String(idSiswa).trim();
    if (idTarget === "") return { ditemukan: false };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setoran");
    if (!sheet) return { ditemukan: false, error: "Sheet Setoran tidak ditemukan!" };
    
    var data = sheet.getDataRange().getValues();
    
    // Looping dari bawah ke atas mencari riwayat terbaru
    for (var i = data.length - 1; i > 0; i--) {
      var idCell = data[i][3]; // Kolom D (Index 3)
      
      // 2. Sabuk Pengaman Sel: Jika sel kosong, lewati saja
      if (idCell === undefined || idCell === null || String(idCell).trim() === "") continue;

      // 3. Pencocokan ID dengan aman
      if (String(idCell).trim() === idTarget) { 
         return {
           ditemukan: true,
           rowIndex: i + 1,              // Baris asli untuk target GSheets
           idQiroah: data[i][0] ? String(data[i][0]).trim() : "", // 🌟 TAMBAHKAN INI! (Kolom A Index 0 untuk Firestore)
           tanggal: data[i][1] ? String(data[i][1]) : "", 
           jilid: data[i][6] ? String(data[i][6]) : "-",      // Kolom G
           halaman: data[i][7] ? String(data[i][7]) : "-",    // Kolom H
           status: data[i][8] ? String(data[i][8]) : "-",     // Kolom I
           catatan: data[i][9] ? String(data[i][9]) : "-",    // Kolom J
           tugasRumah: data[i][10] ? String(data[i][10]) : "-", // Kolom K
           murojaah: data[i][15] ? String(data[i][15]) : "-",   // Kolom P
           membacaSekolah: data[i][16] ? String(data[i][16]) : "-", // Kolom Q
           murojaahTeman: data[i][17] ? String(data[i][17]) : "-"   // Kolom R
         };
      }
    }
    
    return { ditemukan: false };
    
  } catch(e) {
    // Tangkap error apapun dan kembalikan dengan aman
    return { ditemukan: false, error: String(e) };
  }
}


// ====================================================================
// 8. FUNGSI AMBIL DATA LAPORAN (GLOBAL & PERSONAL - SINKRON INDEKS KOLOM)
// ====================================================================
function getLaporanData(mode, targetId, dataSiswaList) {
  try {
    var sheetSetoran = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setoran");
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    var hasil = [];

    if (mode === "GLOBAL") {
      for (var j = 0; j < dataSiswaList.length; j++) {
        var idSiswa = dataSiswaList[j].id;
        var namaSiswa = dataSiswaList[j].nama;
        var ditemukan = false;

        for (var i = dataSetoran.length - 1; i > 0; i--) {
          if (dataSetoran[i][3] === idSiswa) { // Kolom D (Index 3)
            hasil.push({
              nama: namaSiswa,
              tanggal: dataSetoran[i][1] ? dataSetoran[i][1].toString() : "-",
              jilid: dataSetoran[i][6],   // Kolom G
              halaman: dataSetoran[i][7], // Kolom H
              status: dataSetoran[i][8]   // Kolom I
            });
            ditemukan = true;
            break; 
          }
        }
        if (!ditemukan) {
          hasil.push({ nama: namaSiswa, tanggal: "-", jilid: "Belum ada data", halaman: "-", status: "-" });
        }
      }
    } else {
      // Mode PERSONAL
      for (var i = dataSetoran.length - 1; i > 0; i--) {
        if (dataSetoran[i][3] === targetId) { // Kolom D
          hasil.push({
            tanggal: dataSetoran[i][1] ? dataSetoran[i][1].toString() : "-",
            jilid: dataSetoran[i][6],
            halaman: dataSetoran[i][7],
            status: dataSetoran[i][8],
            catatan: dataSetoran[i][9] || "-" // Kolom J adalah Catatan
          });
        }
      }
    }
    return { sukses: true, data: hasil, mode: mode };
  } catch(e) {
    return { sukses: false, error: e.toString() };
  }
}

function getRekapSetoranAdmin() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var mapKelas = {}; 

    for (var i = 1; i < dataSiswa.length; i++) {
      mapKelas[dataSiswa[i][0]] = dataSiswa[i][3]; // Index 3: Kelas
    }

    var sheetSetoran = ss.getSheetByName("Setoran");
    var data = sheetSetoran.getDataRange().getValues();
    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var idSiswaSetoran = data[i][3]; // Kolom D
      hasil.push({
        tanggal: data[i][1] ? data[i][1].toString() : "-", 
        namaSiswa: data[i][4],     
        kelas: mapKelas[idSiswaSetoran] || "-", 
        ustadz: data[i][5],        
        capaian: (data[i][6] || "") + " - " + (data[i][7] || ""), 
        status: data[i][8],        
        catatan: data[i][9] || "-" 
      });
    }
    return { sukses: true, data: hasil, mode: "ADMIN_REKAP" };
  } catch(e) {
    return { sukses: false, error: e.toString() };
  }
}

function getFilterOptions(userRole, userName) {
  var result = {
    kelas: [],
    siswa: []
  };
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var sheetSetoran = ss.getSheetByName("Setoran");
    
    if (!sheetSiswa || !sheetSetoran) return result;
    
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    
    var roleClean = userRole ? userRole.toString().toLowerCase().trim() : "ustadz";
    var nameClean = userName ? userName.toString().toLowerCase().trim() : "";
    
    // 1. Cari ID Siswa binaan Ustadz dari sheet Setoran
    var idSiswaBinaan = [];
    if (roleClean !== "admin") {
      for (var i = 1; i < dataSetoran.length; i++) {
        var rowUstadz = dataSetoran[i][5] ? dataSetoran[i][5].toString().toLowerCase().trim() : ""; // Kolom F
        var idSiswaRow = dataSetoran[i][3] ? dataSetoran[i][3].toString().trim() : ""; // Kolom D
        if (rowUstadz === nameClean && idSiswaRow) {
          if (idSiswaBinaan.indexOf(idSiswaRow) === -1) {
            idSiswaBinaan.push(idSiswaRow);
          }
        }
      }
    }
    
    // 2. Ambil data Kelas dan Siswa dari master sheet "Siswa"
    var setKelas = [];
    var listSiswa = [];
    
    for (var j = 1; j < dataSiswa.length; j++) {
      var row = dataSiswa[j];
      var idSiswaMaster = row[0] ? row[0].toString().trim() : ""; // Kolom A (ID)
      var namaSiswaMaster = row[1] ? row[1].toString().trim() : ""; // Kolom B (Nama)
      var kelasMaster = row[3] ? row[3].toString().trim() : ""; // Kolom D (Kelas)
      
      if (roleClean === "admin" || idSiswaBinaan.indexOf(idSiswaMaster) !== -1) {
        if (kelasMaster && setKelas.indexOf(kelasMaster) === -1) {
          setKelas.push(kelasMaster);
        }
        if (idSiswaMaster && namaSiswaMaster) {
          listSiswa.push({
            id: idSiswaMaster,
            nama: namaSiswaMaster
          });
        }
      }
    }
    
    result.kelas = setKelas.sort();
    result.siswa = listSiswa.sort(function(a, b) {
      return a.nama.localeCompare(b.nama);
    });
    
    return result;
    
  } catch (error) {
    Logger.log("Error getFilterOptions: " + error.toString());
    return result;
  }
}

function getDataLaporan(tglMulai, tglSelesai, kategori, filterVal, roleUser, namaUser) {
  if (!tglMulai || !tglSelesai) return [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSetoran = ss.getSheetByName("Setoran");
  var dataSetoran = sheetSetoran.getDataRange().getValues();

  var start = new Date(tglMulai);
  var end = new Date(tglSelesai);
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);

  var sheetSiswa = ss.getSheetByName("Siswa");
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var mapKelas = {};
  for(var i = 1; i < dataSiswa.length; i++) {
    mapKelas[dataSiswa[i][0]] = dataSiswa[i][3]; // Index 3 (Kelas)
  }

  var hasil = [];
  for (var i = 1; i < dataSetoran.length; i++) {
    var row = dataSetoran[i];
    if (!row[1]) continue; 
    var tglRow = new Date(row[1]);

    if (tglRow >= start && tglRow <= end) {
      var idSiswa = row[3]; 
      var kelasSiswa = mapKelas[idSiswa] || "-";
      var namaUstadzData = row[5]; 

      var isHakAksesLolos = true;
      if (roleUser && roleUser.toString().toLowerCase() !== "admin") {
        if (namaUstadzData !== namaUser) isHakAksesLolos = false; 
      }

      if (isHakAksesLolos) {
        var cocok = false;
        if (kategori === "SEMUA") cocok = true;
        else if (kategori === "PER_KELAS" && kelasSiswa == filterVal) cocok = true;
        else if (kategori === "PER_SISWA" && idSiswa == filterVal) cocok = true;

        if (cocok) {
          hasil.push({
            tanggal: Utilities.formatDate(tglRow, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            namaSiswa: row[4], 
            kelas: kelasSiswa,
            ustadz: row[5],    
            capaian: row[6] + " - " + row[7], 
            status: row[8],    
            catatan: row[9]    
          });
        }
      }
    }
  }
  return hasil;
}

function getLaporanProgressFormal(tglMulai, tglSelesai, namaUstadz, unit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSetoran = ss.getSheetByName("Setoran");
  var sheetSetting = ss.getSheetByName("Setting");
  var dataSetoran = sheetSetoran.getDataRange().getValues();
  var dataSetting = sheetSetting.getDataRange().getValues();
  
  var namaKepsek = "Kepala Sekolah Belum Diatur";
  var varDicari = "Kepsek_" + unit;
  
  for (var i = 1; i < dataSetting.length; i++) {
    if (dataSetting[i][0] == varDicari) {
      namaKepsek = dataSetting[i][1];
      break;
    }
  }
  
  var dMulai = new Date(tglMulai);
  var dSelesai = new Date(tglSelesai);
  dSelesai.setDate(dSelesai.getDate() + 1); 
  var progressSiswa = {};
  
  for (var j = 1; j < dataSetoran.length; j++) {
    var row = dataSetoran[j];
    var tglRow = new Date(row[1]); 
    var namaSiswaRow = row[4];     
    var ustadzRow = row[5];        
    var jilid = row[6];            
    var halaman = row[7];          
    var capaian = jilid + " - " + halaman;
    
    if (tglRow >= dMulai && tglRow < dSelesai && ustadzRow === namaUstadz) {
      if (!progressSiswa[namaSiswaRow]) {
        progressSiswa[namaSiswaRow] = { nama: namaSiswaRow, capaianAwal: capaian, capaianAkhir: capaian, tglAwal: tglRow, tglAkhir: tglRow };
      } else {
        progressSiswa[namaSiswaRow].capaianAkhir = capaian;
        progressSiswa[namaSiswaRow].tglAkhir = tglRow;
      }
    }
  }
  var hasilAkhir = [];
  for (var key in progressSiswa) { hasilAkhir.push(progressSiswa[key]); }
  
  return {
    data: hasilAkhir, kepsek: namaKepsek, ustadz: namaUstadz, unit: unit,
    periode: Utilities.formatDate(dMulai, "GMT+7", "dd/MM/yyyy") + " s/d " + Utilities.formatDate(new Date(tglSelesai), "GMT+7", "dd/MM/yyyy")
  };
}

function getLaporanDataForPrint(tglMulai, tglSelesai, ustadz, unit, kategori, filterVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- 1. Ambil Nama Kepsek ---
    var sheetSetting = ss.getSheetByName("Setting");
    var dataSetting = sheetSetting.getDataRange().getValues();
    var namaKepsek = "Kepala Sekolah"; 
    var searchKey = "Kepsek_" + unit; 
    for (var s = 1; s < dataSetting.length; s++) {
      if (dataSetting[s][0] && dataSetting[s][0].toString().toLowerCase() === searchKey.toLowerCase()) {
        namaKepsek = dataSetting[s][1]; 
        break;
      }
    }

    // --- 2. Ambil Data Setoran & Siswa ---
    var sheetSetoran = ss.getSheetByName("Setoran");
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var mapKelas = {};
    for(var i = 1; i < dataSiswa.length; i++) { mapKelas[dataSiswa[i][0]] = dataSiswa[i][3]; }
    
    var tglAwal = new Date(tglMulai); tglAwal.setHours(0,0,0,0);
    var tglAkhir = new Date(tglSelesai); tglAkhir.setHours(23,59,59,999);
    
    var listData = []; // Sekarang pakai array biasa, mengambil SEMUA riwayat
    
    // --- 3. Proses Looping Filter ---
    for (var i = 1; i < dataSetoran.length; i++) {
      var row = dataSetoran[i];
      if (!row[1]) continue; 
      
      var rowDate = new Date(row[1]);
      if (rowDate >= tglAwal && rowDate <= tglAkhir) {
        var rowUnit = row[2] ? row[2].toString().toLowerCase() : "";
        var idSiswa = row[3];
        var rowNamaSiswa = row[4];  // Kolom E
        var rowUstadz = row[5];
        
        // Capaian: Jilid (G / 6) - Halaman (H / 7) - Status Kelancaran (I / 8)
        var jilid = row[6] ? row[6] : "";
        var hal = row[7] ? row[7] : "";
        var kelancaran = row[8] ? " (" + row[8] + ")" : "";
        var rowCapaian = jilid + " - " + hal + kelancaran; 
        
        // Status Tugas (Status Setoran M / 12)
        var rowStatusTugas = row[12] ? row[12] : "-";
        
        var kelasSiswa = mapKelas[idSiswa] || "-";
        
        // Filter Ustadz & Unit
        if (rowUstadz !== ustadz) continue;
        if (unit && rowUnit) {
          var cleanUnit = unit.toString().toLowerCase();
          if (!cleanUnit.includes(rowUnit) && !rowUnit.includes(cleanUnit)) continue;
        }
        
        // Filter Kategori (Semua/Kelas/Siswa)
        var cocok = false;
        if (kategori === "SEMUA" || !kategori) cocok = true;
        else if (kategori === "PER_KELAS" && kelasSiswa == filterVal) cocok = true;
        else if (kategori === "PER_SISWA" && idSiswa == filterVal) cocok = true;
        
        if (!cocok) continue;
        
        // Format Tanggal jadi DD/MM/YYYY
        var tglFormat = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");

        // Masukkan data baris ini ke dalam list
        listData.push({
          nama: rowNamaSiswa,
          tgl: tglFormat,
          capaian: rowCapaian,
          statusTugas: rowStatusTugas
        });
      }
    }
    
    return { status: "success", data: listData, namaKepsek: namaKepsek };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getLaporanRekapQiroahDataForPrint(tglMulai, tglSelesai, ustadz, unit, kategori, filterVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ambil Nama Kepsek
    var sheetSetting = ss.getSheetByName("Setting");
    var dataSetting = sheetSetting ? sheetSetting.getDataRange().getValues() : [];
    var namaKepsek = "Kepala Sekolah"; 
    var searchKey = "Kepsek_" + (unit || ""); 
    for (var s = 1; s < dataSetting.length; s++) {
      if (dataSetting[s][0] && dataSetting[s][0].toString().toLowerCase() === searchKey.toLowerCase()) {
        namaKepsek = dataSetting[s][1]; 
        break;
      }
    }

    // 2. Ambil Data Setoran & Siswa
    var sheetSetoran = ss.getSheetByName("Setoran");
    if (!sheetSetoran) return { status: "success", data: [], namaKepsek: namaKepsek };
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa ? sheetSiswa.getDataRange().getValues() : [];
    var mapKelas = {};
    for (var i = 1; i < dataSiswa.length; i++) { 
      var idMaster = dataSiswa[i][0] ? dataSiswa[i][0].toString().trim() : "";
      var kelasMaster = dataSiswa[i][3] ? dataSiswa[i][3].toString().trim() : "";
      if (idMaster) mapKelas[idMaster] = kelasMaster;
    }
    
    var tglAwal = new Date(tglMulai); tglAwal.setHours(0,0,0,0);
    var tglAkhir = new Date(tglSelesai); tglAkhir.setHours(23,59,59,999);
    
    var rekapSiswaMap = {};
    var unitClean = unit ? unit.toString().toLowerCase().trim() : "";
    var ustadzClean = ustadz ? ustadz.toString().toLowerCase().trim() : "";
    
    // 3. Looping dan Rekapitulasi per Siswa
    for (var i = 1; i < dataSetoran.length; i++) {
      var row = dataSetoran[i];
      if (!row[1]) continue; 
      
      var rowDate = new Date(row[1]);
      if (isNaN(rowDate.getTime())) continue;
      
      if (rowDate >= tglAwal && rowDate <= tglAkhir) {
        var rowUnit = row[2] ? row[2].toString().toLowerCase().trim() : "";
        var idSiswa = row[3] ? row[3].toString().trim() : "";
        var rowNamaSiswa = row[4] ? row[4].toString().trim() : ""; 
        var rowUstadz = row[5] ? row[5].toString().toLowerCase().trim() : "";
        
        if (!rowNamaSiswa) continue;
        
        var jilid = row[6] ? row[6] : "";
        var hal = row[7] ? row[7] : "";
        var kelancaran = row[8] ? " (" + row[8] + ")" : "";
        var rowCapaian = jilid + " - " + hal + kelancaran; 
        
        var kelasSiswa = mapKelas[idSiswa] || "-";
        
        // A. FILTER UNIT: Pastikan hanya unit yang dipilih yang masuk (Mencegah unit lain ikut tampil)
        if (unitClean && rowUnit && rowUnit !== unitClean) {
          continue;
        }
        
        // B. FILTER KATEGORI (SEMUA / PER_KELAS / PER_SISWA)
        var cocok = false;
        if (kategori === "SEMUA" || !kategori) {
          cocok = true;
        } else if (kategori === "PER_KELAS") {
          if (!filterVal || kelasSiswa.toLowerCase().trim() === filterVal.toString().toLowerCase().trim()) cocok = true;
        } else if (kategori === "PER_SISWA") {
          if (!filterVal || idSiswa === filterVal.toString().trim() || rowNamaSiswa.toLowerCase().trim() === filterVal.toString().toLowerCase().trim()) cocok = true;
        }
        
        if (!cocok) continue;
        
        // C. FILTER USTADZ: Hanya diterapkan jika memilih kategori PER_SISWA (pribadi).
        // Untuk PER_KELAS, dibebaskan agar seluruh siswa di kelas tersebut (24 siswa) tampil lengkap.
        if (kategori === "PER_SISWA" && ustadzClean && rowUstadz && rowUstadz !== ustadzClean) {
          continue;
        }
        
        var tglFormat = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        var mapKey = idSiswa || rowNamaSiswa.toLowerCase();

        if (!rekapSiswaMap[mapKey]) {
          rekapSiswaMap[mapKey] = {
            nama: rowNamaSiswa,
            tglPertama: rowDate,
            tglAwalStr: tglFormat,      
            tglTerakhirObj: rowDate,
            tglTerakhirStr: tglFormat, 
            capaianAwal: rowCapaian, 
            capaianAkhir: rowCapaian  
          };
        } else {
          var item = rekapSiswaMap[mapKey];
          // Update Capaian Akhir (Tanggal yang paling baru)
          if (rowDate >= item.tglTerakhirObj) {
            item.tglTerakhirObj = rowDate;
            item.tglTerakhirStr = tglFormat;
            item.capaianAkhir = rowCapaian;
          }
          // Update Capaian Awal (Tanggal yang paling awal dalam rentang)
          if (rowDate < item.tglPertama) {
            item.tglPertama = rowDate;
            item.tglAwalStr = tglFormat;
            item.capaianAwal = rowCapaian;
          }
        }
      }
    }
    
    var listData = [];
    for (var key in rekapSiswaMap) {
      listData.push({
        nama: rekapSiswaMap[key].nama,
        tglAwal: rekapSiswaMap[key].tglAwalStr,          
        capaianAwal: rekapSiswaMap[key].capaianAwal,
        tglTerakhirStr: rekapSiswaMap[key].tglTerakhirStr, 
        capaianAkhir: rekapSiswaMap[key].capaianAkhir
      });
    }
    
    // Urutkan berdasarkan abjad nama murid
    listData.sort(function(a, b) {
      return a.nama.localeCompare(b.nama);
    });
    
    return { status: "success", data: listData, namaKepsek: namaKepsek };
  } catch (err) {
    return { status: "error", message: err.toString(), data: [] };
  }
}

// Fungsi wajib untuk mengembalikan URL Web Apps yang aktif ke frontend
function getScriptUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    return "";
  }
}

// ====================================================================
// FUNGSI PROSES AMBIL LAPORAN (VERSI TAMPIL SEMUA RIWAYAT)
// ====================================================================
function prosesAmbilLaporan(tglMulai, tglSelesai, kategori, filterVal, roleUser, namaUser) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetoran = ss.getSheetByName("Setoran");
    var sheetSiswa = ss.getSheetByName("Siswa");
    
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    
    var start = new Date(tglMulai);
    var end = new Date(tglSelesai);
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);
    
    // 1. Buat peta data siswa (ID -> {Nama, Kelas})
    var mapSiswa = {};
    for (var i = 1; i < dataSiswa.length; i++) {
      var idS = dataSiswa[i][0] ? dataSiswa[i][0].toString().trim() : "";
      if (idS) {
        mapSiswa[idS] = {
          nama: dataSiswa[i][1] ? dataSiswa[i][1].toString().trim() : "",
          kelas: dataSiswa[i][3] ? dataSiswa[i][3].toString().trim() : ""
        };
      }
    }
    
    var logSetoranSemua = []; // Menyimpan SEMUA baris transaksi untuk Tabel
    var logPerSiswa = {};     // Menyimpan rekap terakhir per siswa untuk Grafik Target
    
    var roleClean = roleUser ? roleUser.toString().toLowerCase().trim() : "";
    var nameClean = namaUser ? namaUser.toString().toLowerCase().trim() : "";
    
    // 2. Kumpulkan semua log setoran
    for (var j = 1; j < dataSetoran.length; j++) {
      var row = dataSetoran[j];
      if (!row[1]) continue; 
      
      var tglRow = new Date(row[1]);
      if (tglRow >= start && tglRow <= end) {
        var idSiswaRow = row[3] ? row[3].toString().trim() : "";
        var ustadzRow = row[5] ? row[5].toString().trim() : "-"; 
        var ustadzLower = row[5] ? row[5].toString().toLowerCase().trim() : "";
        
        // Cek hak akses
        if (roleClean !== "admin" && ustadzLower !== nameClean) continue;
        
        var infoSiswa = mapSiswa[idSiswaRow] || { nama: row[4] || "Tanpa Nama", kelas: "-" };
        
        // Cek kategori filter
        if (kategori === "PER_KELAS" && infoSiswa.kelas !== filterVal) continue;
        if (kategori === "PER_SISWA" && idSiswaRow !== filterVal) continue;
        
        var capaianStr = (row[6] || "-") + " (Hal " + (row[7] || "-") + ")"; 
        // Tambahkan jam dan menit agar jika input 3x sehari terlihat beda waktunya
        var tglFormat = Utilities.formatDate(tglRow, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        
        var dataBaris = {
          waktu: tglRow.getTime(),
          namaSiswa: infoSiswa.nama,
          kelas: infoSiswa.kelas,
          tgl: tglFormat,
          ustadz: ustadzRow,
          capaianAwal: "-", // Dikosongkan karena ini log per kejadian, bukan summary progres
          capaianAkhir: capaianStr
        };

        // Push ke array tabel
        logSetoranSemua.push(dataBaris);
        
        // Simpan/Timpa ke logPerSiswa untuk data grafik (ambil yang paling baru)
        if (!logPerSiswa[idSiswaRow]) {
          logPerSiswa[idSiswaRow] = dataBaris;
        } else {
          if (tglRow.getTime() > logPerSiswa[idSiswaRow].waktu) {
            logPerSiswa[idSiswaRow] = dataBaris;
          }
        }
      }
    }
    
    // 3. Urutkan Tabel: Berdasarkan Nama Siswa, lalu Waktu Transaksi (Terbaru di atas)
    logSetoranSemua.sort(function(a, b) { 
      var cmp = a.namaSiswa.localeCompare(b.namaSiswa);
      if (cmp === 0) return b.waktu - a.waktu; 
      return cmp;
    });

    // 4. Ekstrak objek unik untuk grafik
    var rekapSiswaArray = [];
    for (var key in logPerSiswa) {
      rekapSiswaArray.push(logPerSiswa[key]);
    }
    
    // Kembalikan 2 set data
    return {
      tabelRiwayat: logSetoranSemua,
      grafikRekap: rekapSiswaArray
    };
    
  } catch (e) {
    throw new Error("Gagal mengolah laporan backend: " + e.toString());
  }
}

function cekLoginWali(idSiswa, passwordOrtu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var data = sheetSiswa.getDataRange().getValues();

    idSiswa = idSiswa.toString().trim();
    passwordOrtu = passwordOrtu.toString().trim();

    for (var i = 1; i < data.length; i++) {

      var dbId = data[i][0].toString().trim();
      var dbPassword = (data[i][7] || "123456").toString().trim();

      if (dbId == idSiswa) {

        if (passwordOrtu == dbPassword) {

          return {
            status: "success",
            data: {
              idSiswa: dbId,
              namaSiswa: data[i][1],
              kelas: data[i][3],
              unit: data[i][6]
            }
          };

        } else {

          return {
            status: "error",
            message: "Password / PIN salah."
          };

        }

      }

    }

    return {
      status: "error",
      message: "ID Siswa tidak ditemukan."
    };

  } catch (e) {

    return {
      status: "error",
      message: e.toString()
    };

  }
}

function getRiwayatSiswa(idSiswa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Setoran");
    var data = sheet.getDataRange().getValues();
    
    var riwayat = [];
    var daftarTugas = [];

    // PENGAMANAN 1: Pastikan ID Siswa yang dicari aman dari null
    var idSiswaAman = String(idSiswa || "").trim();

    for (var i = 1; i < data.length; i++) {
      // PENGAMANAN 2: Hindari .toString(), gunakan String() untuk kolom ID
      var idSiswaDiSheet = String(data[i][3] || "").trim();

      // Hanya proses jika ID di baris tersebut tidak kosong dan cocok dengan ID target
      if (idSiswaDiSheet !== "" && idSiswaDiSheet === idSiswaAman) {
        
        // PENGAMANAN 3: Bungkus seluruh kolom dengan String() agar 100% bebas dari error null
        var barisData = {
          tanggal: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], "GMT+7", "dd/MM/yyyy") : String(data[i][1] || ""),
          jilid: String(data[i][6] || ""),
          halaman: String(data[i][7] || ""),
          ustadz: String(data[i][5] || ""),
          statusLancar: String(data[i][8] || ""),
          catatan: String(data[i][9] || ""),
          tugas: String(data[i][10] || "").trim(),
          statusSetoran: String(data[i][12] || ""),
          nilai: String(data[i][13] || "").trim(),
          catatanGuru: String(data[i][14] || "").trim(),
          barisAsli: i + 1
        };
        
        riwayat.push(barisData);

        // Jika kolom tugas ada isinya (bukan kosong dan bukan cuma tanda strip), masukkan ke daftarTugas
        if (barisData.tugas !== "" && barisData.tugas !== "-") {
          daftarTugas.push(barisData);
        }
      }
    }
    
    // Membalik urutan agar data riwayat dan tugas yang terbaru berada di urutan paling atas
    riwayat.reverse();
    daftarTugas.reverse();

    return { 
      status: "success", 
      riwayat: riwayat, 
      daftarTugas: daftarTugas 
    };
    
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// ====================================================================
// FUNGSI UPDATE STATUS TUGAS
// SINKRONISASI SHEET & FIRESTORE
// UPDATE HANYA FIELD status_setoran
// ====================================================================
function updateStatusCentang(idAtauRow, isSelesai) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Setoran");
    var db = getFirestoreDB();

    var targetRow = -1;
    var parameterInput = String(idAtauRow).trim();
    var docId = "";

    // ================================================================
    // 1. CARI BARIS & IDENTIFIKASI ID DOKUMEN
    // ================================================================
    if (parameterInput.indexOf("QR") === 0) {

      // Dipanggil dari Portal Wali
      docId = parameterInput;

      var dataA = sheet.getRange("A:A").getValues();

      for (var i = dataA.length - 1; i >= 0; i--) {
        if (String(dataA[i][0]).trim() === parameterInput) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow === -1) {
        throw new Error("ID Data tidak ditemukan di Spreadsheet.");
      }

    } else {

      // Jika yang masuk berupa nomor baris
      targetRow = parseInt(parameterInput);

      if (isNaN(targetRow) || targetRow < 1) {
        throw new Error("Parameter baris tidak valid.");
      }

      docId = String(
        sheet.getRange(targetRow, 1).getValue()
      ).trim();
    }

    // ================================================================
    // 2. STATUS YANG AKAN DISIMPAN
    // ================================================================
    var statusTeks = isSelesai
      ? "Sudah Dikerjakan"
      : "";

    // ================================================================
    // 3. UPDATE GOOGLE SHEET
    // ================================================================
    sheet.getRange(targetRow, 13).setValue(statusTeks);

    // ================================================================
    // 4. UPDATE FIRESTORE
    //    PENTING:
    //    parameter TRUE = HANYA update field status_setoran
    // ================================================================
    if (docId !== "") {

      try {

        db.updateDocument(
          "Setoran/" + docId,
          {
            status_setoran: statusTeks
          },
          true
        );

      } catch (fsError) {

        throw new Error(
          "Update Sheet berhasil, tetapi Firestore gagal: " +
          String(fsError)
        );
      }

    } else {

      throw new Error(
        "ID Dokumen kosong, tidak bisa update Firestore."
      );
    }

    return {
      status: "success",
      message: "Status tugas berhasil diperbarui di Sheet dan Firestore!"
    };

  } catch (e) {

    throw new Error(
      "Gagal memproses konfirmasi: " + String(e)
    );
  }
}

function ujiCobaMonitoring() {
  // Ganti tanggal di bawah ini dengan bulan yang PASTI ada data setorannya
  var tglMulai = "2026-07-01"; 
  var tglAkhir = "2026-07-31"; 
  var role = "Admin";
  var unit = "Semua Unit";
  
  var hasil = getMonitoringUstadz(tglMulai, tglAkhir, role, unit);
  
  // Menampilkan hasil mentah ke log eksekusi
  Logger.log("Total Ustadz: " + hasil.total_ustadz_aktif);
  Logger.log("Total Setoran: " + hasil.total_setoran);
  Logger.log(JSON.stringify(hasil.data, null, 2));
}

// ==========================================
// FITUR PENILAIAN TUGAS RUMAH
// ==========================================

// 1. Mengambil daftar Siswa sesuai Unit Ustadz yang login
function getSiswaUntukPenilaian(userUnit, userRole) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var listSiswa = [];
    
    var roleAkses = userRole ? userRole.toString().toLowerCase() : "";
    var isAdminPusat = (roleAkses.includes("admin") || roleAkses.includes("kepala"));
    
    for (var i = 1; i < dataSiswa.length; i++) {
      var idSiswa = dataSiswa[i][0]; // Kolom A
      var namaSiswa = dataSiswa[i][1]; // Kolom B
      var unitSiswa = dataSiswa[i][6]; // Kolom G
      
      if (idSiswa !== "") {
        if (isAdminPusat || unitSiswa === userUnit) {
          listSiswa.push({ id: idSiswa, nama: namaSiswa });
        }
      }
    }
    
    // Urutkan abjad
    listSiswa.sort(function(a, b){
      if(a.nama < b.nama) return -1;
      if(a.nama > b.nama) return 1;
      return 0;
    });
    
    return { sukses: true, data: listSiswa };
  } catch (e) {
    return { sukses: false, pesan: e.message };
  }
}

// 2. Mengambil riwayat Tugas Rumah berdasarkan ID Siswa (Untuk Halaman Wali)
function getTugasSiswa(idSiswa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetoran = ss.getSheetByName("Setoran");
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    var listTugas = [];
    
    for (var i = 1; i < dataSetoran.length; i++) {
      var idQiroahDb = dataSetoran[i][0]; // Kolom A
      var idSiswaDb = dataSetoran[i][3];  // Kolom D
      var tugasRumah = dataSetoran[i][10]; // Kolom K
      
      // Ambil baris jika ID Siswa cocok dan ada Tugas Rumah
      if (String(idSiswaDb).trim() === String(idSiswa).trim() && tugasRumah && String(tugasRumah).trim() !== "") {
        var tgl = dataSetoran[i][1]; // Kolom B
        var tglFormat = (Object.prototype.toString.call(tgl) === '[object Date]') 
                        ? Utilities.formatDate(tgl, "Asia/Jakarta", "dd/MM/yyyy") 
                        : String(tgl).split(' ')[0];
        
        // Ambil Nilai (Kolom N / index 13) dan Catatan (Kolom O / index 14)
        var nilaiGuru = dataSetoran[i][13] ? String(dataSetoran[i][13]).trim() : "";
        var catatanGuru = dataSetoran[i][14] ? String(dataSetoran[i][14]).trim() : "";
        
        listTugas.push({
          id_qiroah: idQiroahDb,
          tanggal: tglFormat,
          tugas: tugasRumah,
          status_tugas: dataSetoran[i][12] || "Belum Dikerjakan", // Kolom M
          nilai: nilaiGuru,     // Kolom N
          catatan: catatanGuru  // Kolom O
        });
      }
    }
    // Urutkan dari yang terbaru
    return { sukses: true, data: listTugas.reverse() }; 
  } catch (e) {
    return { sukses: false, pesan: e.message };
  }
}

// ====================================================================
// 3. Menyimpan Nilai & Catatan ke Sheet Setoran DAN Firestore (Full Sync)
// ====================================================================
function simpanNilaiTugas(idQiroah, nilai, catatan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetoran = ss.getSheetByName("Setoran");
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    var db = getFirestoreDB(); // Pastikan koneksi Firestore dipanggil
    
    for (var i = 1; i < dataSetoran.length; i++) {
      if (dataSetoran[i][0] == idQiroah) { // Cocokkan ID Qiroah (Kolom A)
        var row = i + 1;
        
        // 1. Update ke Google Sheet terlebih dahulu
        sheetSetoran.getRange(row, 14).setValue(nilai); // Kolom N (Nilai)
        sheetSetoran.getRange(row, 15).setValue(catatan); // Kolom O (Catatan Penilaian)
        
        // 2. Ambil SEMUA data di baris tersebut setelah diupdate (Kolom A sampai R / 1 sampai 18)
        var rowValues = sheetSetoran.getRange(row, 1, 1, 18).getValues()[0];
        
        var docId = String(idQiroah).trim();
        var tanggalSheet = rowValues[1]; // Kolom B
        
        // Memastikan format tanggal valid untuk Firestore
        var tanggalValid = (tanggalSheet instanceof Date) ? tanggalSheet : new Date();
        
        // 3. Timpa Firestore dengan SELURUH data lengkap dari Sheet agar field lain tidak hilang
        db.updateDocument("Setoran/" + docId, {
          tanggal: tanggalValid,
          unit: String(rowValues[2] || "").trim(),
          id_siswa: String(rowValues[3] || "").trim(),
          nama_siswa: String(rowValues[4] || "").trim(),
          ustadz: String(rowValues[5] || "").trim(),
          jilid: String(rowValues[6] || "").trim(),
          halaman: String(rowValues[7] || "").trim(),
          status_lancar: String(rowValues[8] || "").trim(),
          catatan: String(rowValues[9] || "").trim(),
          tugas_rumah: String(rowValues[10] || "").trim(),
          id_audio: String(rowValues[11] || "").trim(),
          status_setoran: String(rowValues[12] || "").trim(),
          nilai: String(nilai || "").trim(),       // Field Nilai yang baru
          catatan_audio: String(catatan || "").trim(), // Field Catatan yang baru
          murojaah_di_rumah: String(rowValues[15] || "").trim(),
          membaca_di_sekolah: String(rowValues[16] || "").trim(),
          murojaah_bersama_teman: String(rowValues[17] || "").trim()
        });
        
        return { sukses: true, pesan: "Penilaian berhasil disimpan di Sheet dan Firestore!" };
      }
    }
    return { sukses: false, pesan: "ID Data tidak ditemukan di Sheet." };
  } catch (e) {
    return { sukses: false, pesan: "Gagal menyimpan nilai: " + String(e) };
  }
}

// ====================================================================
// FUNGSI AMBIL RIWAYAT TUGAS RUMAH UNTUK TABEL PENILAIAN
// ====================================================================
function getTugasUntukPenilaian(idSiswa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Setoran");
    var data = sheet.getDataRange().getDisplayValues(); 
    var hasil = [];

    // Mulai dari baris 2 (index 1)
    for (var i = 1; i < data.length; i++) {
      var rowIdSiswa = (data[i][3] || "").toString().trim(); // Kolom D: ID Siswa
      var tugas = (data[i][10] || "").toString().trim();     // Kolom K: Tugas Rumah

      if (rowIdSiswa === idSiswa && tugas !== "") { 
        hasil.push({
          idQiroah: data[i][0],            // <=== TAMBAHAN BARU: Mengambil ID Qiroah dari Kolom A
          baris: i + 1, 
          tanggal: data[i][1],             
          catatanGuru: data[i][9],         
          tugasRumah: tugas,               
          statusSetoran: data[i][12],      
          nilai: data[i][13],              
          catatanPenilaian: data[i][14]    
        });
      }
    }
    return hasil.reverse(); 
  } catch(e) {
    return [];
  }
}

// ====================================================================
// FUNGSI BARU: AMBIL DATA TUGAS DI RUMAH DARI SHEET SETORAN
// ====================================================================
function getMonitoringTugasGS(tglMulaiStr, tglAkhirStr, kelasTarget, userUnit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetoran = ss.getSheetByName("Setoran");
    var sheetSiswa = ss.getSheetByName("Siswa");
    
    if (!sheetSetoran) return { sukses: false, pesan: "Sheet Setoran tidak ditemukan!" };

    // 1. Ambil relasi Kelas DAN Unit (di sheet Siswa)
    // Pastikan ada kolom "Kelas" dan "Unit" di sheet Siswa
    var mapSiswa = {}; 
    var colUnitName = "Unit"; // <--- SESUAIKAN DENGAN NAMA HEADER DI SHEET SISWA ANDA
    
    if (sheetSiswa) {
       var dataSiswa = sheetSiswa.getDataRange().getValues();
       var headers = dataSiswa[0] || [];
       var idxId = -1, idxKelas = -1, idxUnit = -1;
       
       for(var c=0; c<headers.length; c++) {
          var h = headers[c].toString().toLowerCase().trim();
          if(h === "id siswa" || h === "id_siswa") idxId = c;
          if(h === "kelas") idxKelas = c;
          if(h === colUnitName.toLowerCase()) idxUnit = c;
       }
       
       if (idxId !== -1) {
          for(var r=1; r<dataSiswa.length; r++) {
             var id = dataSiswa[r][idxId] ? dataSiswa[r][idxId].toString().trim() : "";
             if (id) {
                mapSiswa[id] = {
                   kelas: idxKelas !== -1 ? dataSiswa[r][idxKelas].toString().trim() : "",
                   unit: idxUnit !== -1 ? dataSiswa[r][idxUnit].toString().trim() : ""
                };
             }
          }
       }
    }

    // 2. Loop Data Setoran dan filter berdasarkan Tanggal, Kelas, DAN Unit
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    var hasil = [];
    var targetMulai = tglMulaiStr.trim(); 
    var targetAkhir = tglAkhirStr.trim();

    for (var i = 1; i < dataSetoran.length; i++) {
       var tglRaw = dataSetoran[i][1]; // Kolom B (Tanggal dan Waktu)
       if (!tglRaw) continue;

       // Variabel untuk kebutuhan filter rentang tanggal (Format yyyy-MM-dd)
       var tglSetoranStr = "";
       if (Object.prototype.toString.call(tglRaw) === '[object Date]') {
           tglSetoranStr = Utilities.formatDate(tglRaw, "Asia/Jakarta", "yyyy-MM-dd");
       } else {
           tglSetoranStr = tglRaw.toString().split(' ')[0];
       }

       if (tglSetoranStr >= targetMulai && tglSetoranStr <= targetAkhir) {
           
           var idSiswa = (dataSetoran[i][3] || "").toString().trim(); // Kolom D (ID Siswa)
           var infoSiswa = mapSiswa[idSiswa] || { kelas: "", unit: "" };
           
           // FILTER UNIT: Jika user memiliki unit (bukan super admin), wajib cocokkan unitnya
           if (userUnit && userUnit !== "" && infoSiswa.unit.toLowerCase() !== userUnit.toLowerCase()) {
               continue; 
           }

           // FILTER KELAS: Jika ada filter kelas, cocokkan
           if (kelasTarget && kelasTarget !== "" && infoSiswa.kelas.toLowerCase() !== kelasTarget.toLowerCase()) {
               continue;
           }

           // === BAGIAN BARU: FORMAT TANGGAL UNTUK DITAMPILKAN KE TABEL ===
           // Kita potong bagian jamnya agar hanya muncul tanggal (Format dd/MM/yyyy)
           var tglTampil = "";
           if (Object.prototype.toString.call(tglRaw) === '[object Date]') {
               tglTampil = Utilities.formatDate(tglRaw, "Asia/Jakarta", "dd/MM/yyyy");
           } else {
               // Jika data berbentuk teks seperti "29/07/2026 13:33:35", kita pisahkan berdasarkan spasi lalu ambil depannya saja
               tglTampil = tglRaw.toString().split(' ')[0]; 
           }

           var namaSiswa = dataSetoran[i][4] ? dataSetoran[i][4].toString().trim() : "-";
           var namaUstadz = dataSetoran[i][5] ? dataSetoran[i][5].toString().trim() : "-";
           var tugasRumah = dataSetoran[i][10] ? dataSetoran[i][10].toString().trim() : "";
           var statusTugas = dataSetoran[i][12] ? dataSetoran[i][12].toString().trim() : "-";

           if (tugasRumah === "" || tugasRumah === "-") continue;

           // Memasukkan data ke dalam array termasuk variabel 'tanggal'
           hasil.push({
               nama_siswa: namaSiswa,
               tugas: tugasRumah,
               tanggal: tglTampil, // <--- Data tanggal disisipkan ke sini
               ustadz: namaUstadz,
               status: statusTugas
           });
       }
    }
    
    hasil.reverse();
    return { sukses: true, data: hasil };
    
  } catch (e) {
    return { sukses: false, pesan: "Error Server: " + e.message };
  }
}

// ====================================================================
// FUNGSI QUERY TERARAH FIRESTORE (SUPER CEPAT & HEMAT KUOTA)
// ====================================================================
function getRiwayatWaliDariFirestore(idSiswa) {
  try {
    var idSiswaCari = String(idSiswa || "").trim();
    if (!idSiswaCari) {
      return { status: "error", message: "ID Siswa tidak ditemukan atau kosong." };
    }

    var db = getFirestoreDB();
    var documents = [];

    // 🚀 MENJALANKAN QUERY TERARAH (Hanya mengambil data ID Siswa terkait)
    if (typeof db.query === 'function') {
      try {
        documents = db.query("Setoran").Where("id_siswa", "==", idSiswaCari).Execute();
      } catch(e1) {
        try {
          documents = db.query("Setoran").where("id_siswa", "==", idSiswaCari).get();
        } catch(e2) {
          documents = db.getDocuments("Setoran");
        }
      }
    } else {
      documents = db.getDocuments("Setoran");
    }

    var riwayatList = [];
    var daftarTugas = [];

    // Helper ekstraksi nilai field Firestore
    function ambilNilai(field) {
      if (field === undefined || field === null) return "";
      if (typeof field === "object") {
        if (field.stringValue !== undefined) return field.stringValue;
        if (field.integerValue !== undefined) return field.integerValue;
        if (field.timestampValue !== undefined) return field.timestampValue;
        if (field.booleanValue !== undefined) return field.booleanValue;
      }
      return String(field);
    }

    for (var i = 0; i < documents.length; i++) {
      var doc = documents[i];
      var fields = doc.fields || doc;
      
      var idSiswaDoc = ambilNilai(fields.id_siswa || fields.idSiswa).trim();

      // Filter presisi berdasarkan ID Siswa
      if (idSiswaDoc === idSiswaCari) {
        var docId = doc.name ? doc.name.split("/").pop() : (doc.id || doc.id_dokumen || "");

        var tglRaw = ambilNilai(fields.tanggal);
        var dateObj = tglRaw ? new Date(tglRaw) : new Date();
        var tglFormatted = Utilities.formatDate(dateObj, "Asia/Jakarta", "dd/MM/yyyy");

        var jilidVal = ambilNilai(fields.jilid);
        var halVal = ambilNilai(fields.halaman);
        var catatanVal = ambilNilai(fields.catatan);
        var lancarVal = ambilNilai(fields.status_lancar);
        var nilaiVal = ambilNilai(fields.nilai);
        var catGuruVal = ambilNilai(fields.catatan_audio);
        var tugasVal = ambilNilai(fields.tugas_rumah);
        var statusSetoranVal = ambilNilai(fields.status_setoran);

        // Data Riwayat
        riwayatList.push({
          idDokumen: docId,
          tanggal: tglFormatted,
          tanggalAsli: dateObj.getTime(),
          jilid: jilidVal,
          halaman: halVal,
          catatan: catatanVal || "-",
          statusLancar: lancarVal || "-",
          nilai: nilaiVal,
          catatanGuru: catGuruVal
        });

        // Data Tugas
        if (tugasVal && tugasVal.trim() !== "" && tugasVal.trim() !== "-") {
          daftarTugas.push({
            idDokumen: docId,
            tanggal: tglFormatted,
            tanggalAsli: dateObj.getTime(),
            tugas: tugasVal,
            statusSetoran: statusSetoranVal,
            nilai: nilaiVal,
            catatanGuru: catGuruVal
          });
        }
      }
    }

    // Urutkan dari data terbaru
    riwayatList.sort(function(a, b) { return b.tanggalAsli - a.tanggalAsli; });
    daftarTugas.sort(function(a, b) { return b.tanggalAsli - a.tanggalAsli; });

    return {
      status: "success",
      riwayat: riwayatList,
      daftarTugas: daftarTugas
    };

  } catch (e) {
    return { status: "error", message: "Gagal membaca Firestore: " + e.toString() };
  }
}

// Fungsi Backend Simpan Data Ujian
function simpanUjianKenaikan(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("ujian_kenaikan");
    if (!sheet) {
      sheet = ss.insertSheet("ujian_kenaikan");
      sheet.appendRow(["ID_Ujian", "Tanggal", "Unit", "ID_Siswa", "Nama_Siswa", "Kelas", "Jilid_Asal", "Jilid_Tujuan", "Nilai_Kelancaran", "Nilai_Makhraj", "Nilai_Harakat", "Nilai_Tajwid", "Nilai_Adab", "Total_Nilai", "Predikat", "Status_Kelulusan", "Catatan", "Penguji"]);
    }

    // Ambil data siswa untuk mendapatkan nama & kelas
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var namaSiswa = "";
    var kelasSiswa = "";
    for (var i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0].toString() === data.idSiswa.toString()) {
        namaSiswa = dataSiswa[i][1]; // Kolom B (Nama)
        kelasSiswa = dataSiswa[i][3]; // Kolom D (Kelas)
        break;
      }
    }

    var idUjian = "UJ-" + new Date().getTime();
    var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([
      idUjian, tglHariIni, data.unit, data.idSiswa, namaSiswa, kelasSiswa,
      data.jilidAsal, data.jilidTujuan, data.kelancaran, data.makhraj,
      data.harakat, data.tajwid, data.adab, data.totalNilai,
      data.predikat, data.status, data.catatan, data.ustadz
    ]);

    return { status: "success" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// Fungsi Ambil Data Siswa untuk Cetak Dokumen
function getDataSiswaForCetakUjian(idSiswa) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    var nama = "";
    var kelas = "";

    for (var i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0].toString() === idSiswa.toString()) {
        nama = dataSiswa[i][1];
        kelas = dataSiswa[i][3];
        break;
      }
    }

    // Ambil juga data ujian terakhir siswa tersebut jika ada di sheet ujian_kenaikan
    var sheetUjian = ss.getSheetByName("ujian_kenaikan");
    var kelancaran = 0, makhraj = 0, harakat = 0, tajwid = 0, adab = 0, totalNilai = 0, predikat = "-", status = "LULUS", catatan = "", ustadz = "";
    
    if (sheetUjian) {
      var dataUjian = sheetUjian.getDataRange().getValues();
      for (var j = dataUjian.length - 1; j >= 1; j--) {
        if (dataUjian[j][3].toString() === idSiswa.toString()) {
          kelancaran = dataUjian[j][8];
          makhraj = dataUjian[j][9];
          harakat = dataUjian[j][10];
          tajwid = dataUjian[j][11];
          adab = dataUjian[j][12];
          totalNilai = dataUjian[j][13];
          predikat = dataUjian[j][14];
          status = dataUjian[j][15];
          catatan = dataUjian[j][16];
          ustadz = dataUjian[j][17];
          break;
        }
      }
    }

    // Ambil Nama Kepsek
    var sheetSetting = ss.getSheetByName("Setting");
    var dataSetting = sheetSetting.getDataRange().getValues();
    var namaKepsek = "Kepala Sekolah";
    for (var s = 1; s < dataSetting.length; s++) {
      if (dataSetting[s][0] && dataSetting[s][0].toString().toLowerCase().includes("kepsek")) {
        namaKepsek = dataSetting[s][1];
        break;
      }
    }

    return {
      status: "success",
      data: {
        nama: nama,
        kelas: kelas,
        kelancaran: kelancaran,
        makhraj: makhraj,
        harakat: harakat,
        tajwid: tajwid,
        adab: adab,
        totalNilai: totalNilai,
        predikat: predikat,
        status: status,
        catatan: catatan,
        ustadz: ustadz
      },
      namaKepsek: namaKepsek
    };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getListSemuaSiswa() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Siswa");
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    // Berdasarkan struktur sheet Anda: Kolom A = ID Siswa, Kolom B = Nama, Kolom D = Kelas
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        list.push({
          id: String(data[i][0]),
          nama: String(data[i][1] || ""),
          kelas: String(data[i][3] || "-")
        });
      }
    }
    
    // Urutkan alfabetis
    list.sort(function(a, b) {
      return a.nama.localeCompare(b.nama);
    });

    return { status: "success", data: list };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function pancingIzinDrive() { DriveApp.getFiles(); }

// 1. Fungsi Menyimpan Pengajuan & Auto-Lookup Kelas/Unit Siswa yang BENAR
function serverSimpanPengajuanKenaikan(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPengajuan = ss.getSheetByName("Pengajuan");
  
  if (!sheetPengajuan) {
    sheetPengajuan = ss.insertSheet("Pengajuan");
    sheetPengajuan.appendRow(["Tanggal Diajukan", "Diajukan Oleh", "Nama Siswa", "Kelas", "Unit", "Jilid", "Tanggal Pelaksanaan", "Nama Penguji", "Status"]);
  }

  var teksDropdown = data.namaSiswa || "";
  var kodeAtauNis = teksDropdown.split("-")[0].trim();
  var namaBersih = teksDropdown;
  var kelasSiswa = "-";
  var unitSiswa = "-";

  // Ambil Kelas & Unit dari sheet Siswa berdasarkan foto struktur Anda
  var sheetSiswa = ss.getSheetByName("Siswa");
  if (sheetSiswa) {
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    for (var i = 1; i < dataSiswa.length; i++) {
      var row = dataSiswa[i];
      // Kolom 0(NIS), 1(Nama), 2(ID Ortu), 3(Kelas), 4(Pembimbing), 5(WA), 6(Unit)
      if (String(row[0]).trim() === kodeAtauNis || String(row[1]).trim() === teksDropdown) {
        namaBersih = row[1];
        kelasSiswa = row[3] || "-"; // Kolom D (Kelas)
        unitSiswa = row[6] || "-";  // Kolom G (Unit)
        break;
      }
    }
  }

  // Fallback pembersih nama jika di sheet tidak ketemu
  if (namaBersih === teksDropdown) {
    var match = teksDropdown.match(/(?:\d+\s*-\s*)?([^(]+)/);
    if (match) namaBersih = match[1].trim();
  }

  sheetPengajuan.appendRow([
    data.tanggalDiajukan || new Date().toLocaleDateString('id-ID'),
    data.diajukanOleh || "Guru",
    namaBersih,
    kelasSiswa,
    unitSiswa,
    data.jilidDiuji || data.jilid || "",
    "",
    "",
    "Menunggu Ujian"
  ]);
  
  return { sukses: true };
}

// 2. Fungsi Membaca Antrean (Disesuaikan dengan Struktur Sheet)
function getDaftarAntrean() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pengajuan");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var antrean = [];
  
  // Baris ke-0 adalah Header, jadi mulai dari i = 1
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Lewati jika Nama Siswa (Kolom C / Index 2) kosong
    if (!row[2] || String(row[2]).trim() === "") continue; 
    
    // Cegat: Jika statusnya Selesai, JANGAN masukkan ke antrean UI
    var statusVal = String(row[8] || "").trim().toLowerCase();
    if (statusVal === "selesai") continue; 
    
    // Format Tanggal Diajukan (Kolom A / Index 0)
    let tglDiajukan = "";
    if (row[0] instanceof Date) {
      tglDiajukan = `${String(row[0].getDate()).padStart(2, '0')}/${String(row[0].getMonth() + 1).padStart(2, '0')}/${row[0].getFullYear()}`;
    } else {
      tglDiajukan = String(row[0] || "-");
    }

    // Format Tanggal Pelaksanaan (Kolom G / Index 6)
    let tglPelaksanaanVal = "";
    if (row[6] instanceof Date) {
      tglPelaksanaanVal = `${row[6].getFullYear()}-${String(row[6].getMonth() + 1).padStart(2, '0')}-${String(row[6].getDate()).padStart(2, '0')}`;
    } else {
      tglPelaksanaanVal = String(row[6] || "");
    }
    
    antrean.push({
      tanggalDiajukan: tglDiajukan,
      diajukanOleh: String(row[1] || "-"),
      namaSiswa: String(row[2] || "-"),
      kelas: String(row[3] || "-"),
      unit: String(row[4] || "-"),
      jilid: String(row[5] || "-"),
      tanggalPelaksanaan: tglPelaksanaanVal,
      namaPenguji: String(row[7] || ""),
      status: String(row[8] || "Menunggu Ujian"),
      rowIndex: i + 1  // Menandai nomor baris asli di Sheet
    });
  }
  
  return antrean;
}

// 3. Fungsi Menyimpan Rekapan & Menghapus dari Antrean UI
// Fungsi Menyimpan Rekapan & Menghitung "Ujian Ke" Otomatis
// 1. Fungsi Menyimpan Rekapan (Dengan Tambahan ID Ujian Otomatis)
function serverSimpanRekapanUjian(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRekap = ss.getSheetByName("Rekapan");
  var sheetSiswa = ss.getSheetByName("Siswa");

  var namaSiswaDicari = String(data.namaSiswa || "").trim().toLowerCase();

  // A. Generate ID Ujian (Format: UJ-TahunBulanTanggal-JamMenitDetik)
  // Contoh: UJ-261015-083022
  var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyMMdd-HHmmss");
  var idUjian = "UJ-" + timestamp;

  // B. Logika Penentu Jilid Tujuan Otomatis
  var jilidAsal = String(data.jilid || "").trim();
  var jilidTujuan = "-";
  var jStr = jilidAsal.toLowerCase();
  
  if (jStr.includes("jilid 1")) jilidTujuan = "At-Tartil Jilid 2";
  else if (jStr.includes("jilid 2")) jilidTujuan = "At-Tartil Jilid 3";
  else if (jStr.includes("jilid 3")) jilidTujuan = "At-Tartil Jilid 4";
  else if (jStr.includes("jilid 4")) jilidTujuan = "At-Tartil Jilid 5";
  else if (jStr.includes("jilid 5")) jilidTujuan = "Al-Qur'an";
  else if (jStr.includes("qur'an") || jStr.includes("quran") || jStr.includes("talaqqi")) jilidTujuan = "Hafalan Mandiri";
  else jilidTujuan = "Lanjut";

  // C. Cari ID Siswa otomatis dari Sheet Siswa
  var idSiswa = "-";
  if (sheetSiswa) {
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    for (var i = 1; i < dataSiswa.length; i++) {
      if (String(dataSiswa[i][1] || "").trim().toLowerCase() === namaSiswaDicari) {
        idSiswa = dataSiswa[i][0]; // Kolom A
        break;
      }
    }
  }

  // D. Menghitung "Ujian Ke" berdasarkan nama anak di sheet Rekapan
  var countUjian = 0;
  if (sheetRekap) {
    var dataRekapan = sheetRekap.getDataRange().getValues();
    for (var i = 1; i < dataRekapan.length; i++) {
      // PERHATIAN: Nama anak sekarang bergeser ke Kolom D (index 3)
      var namaDiSheet = String(dataRekapan[i][3] || "").trim().toLowerCase();
      if (namaDiSheet === namaSiswaDicari) {
        countUjian++;
      }
    }
  }
  var ujianKe = countUjian + 1;

  // E. Simpan data ke Sheet Rekapan (Total 21 Kolom dari A sampai U)
  sheetRekap.appendRow([
    idUjian,                     // 1A : ID Ujian (BARU - GENERATE OTOMATIS)
    data.tanggalPelaksanaan,     // 1B : Tanggal Ujian
    idSiswa,                     // 1C : ID Siswa
    data.namaSiswa,              // 1D : Nama
    data.kelas,                  // 1E : Kelas
    data.diajukanOleh,           // 1F : Guru Pengampu
    jilidAsal,                   // 1G : Jilid Asal
    jilidTujuan,                 // 1H : Jilid Tujuan
    ujianKe,                     // 1I : Ujian ke
    data.namaPenguji,            // 1J : Penguji
    data.kelancaran,             // 1K : Kelancaran
    data.makhraj,                // 1L : Makhraj
    data.harakat,                // 1M : Harakat
    data.tajwid,                 // 1N : Tajwid
    data.adab,                   // 1O : Adab
    data.kelebihan,              // 1P : Kelebihan Siswa
    data.perbaikan,              // 1Q : Materi yang Perlu Diperbaiki
    data.rekomendasi,            // 1R : Rekomendasi Tindak Lanjut
    data.predikat,               // 1S : Predikat
    data.keputusan,              // 1T : Keputusan
    data.totalNilai              // 1U : Nilai Akhir
  ]);

  // F. Ubah status di Sheet Pengajuan menjadi "Selesai" agar hilang dari antrean UI
  var sheetPengajuan = ss.getSheetByName("Pengajuan");
  if (sheetPengajuan && data.rowIndex) {
    sheetPengajuan.getRange(data.rowIndex, 9).setValue("Selesai"); 
  }

  return { sukses: true };
}


// 2. Fungsi Mengambil Data Rekapan untuk UI (Semua Index Bergeser +1)
function getDataRekapan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRekap = ss.getSheetByName("Rekapan");
  var sheetSiswa = ss.getSheetByName("Siswa");
  if (!sheetRekap) return [];
  
  var unitMap = {};
  if (sheetSiswa) {
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    for (var i = 1; i < dataSiswa.length; i++) {
       var nm = String(dataSiswa[i][1]).trim().toLowerCase();
       var un = String(dataSiswa[i][6]).trim(); // Kolom G (Unit)
       if (nm) unitMap[nm] = un;
    }
  }

  var data = sheetRekap.getDataRange().getValues();
  var rekapan = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[3]) continue; // Lewati jika Kolom D (Nama) kosong
    
    // row[0] sekarang adalah ID Ujian
    let tgl = row[1] instanceof Date ? `${String(row[1].getDate()).padStart(2, '0')}/${String(row[1].getMonth() + 1).padStart(2, '0')}/${row[1].getFullYear()}` : String(row[1] || "");
    let namaSiswa = String(row[3] || "-");
    let unitSiswa = unitMap[namaSiswa.toLowerCase()] || "-"; 
    
    rekapan.push({
      idUjian: String(row[0] || "-"),
      tanggal: tgl,
      idSiswa: String(row[2] || "-"),
      namaSiswa: namaSiswa,
      kelas: String(row[4] || "-"),
      unit: unitSiswa, 
      guruPengampu: String(row[5] || "-"),
      jilidAsal: String(row[6] || "-"),
      jilidTujuan: String(row[7] || "-"),
      ujianKe: String(row[8] || "-"),
      penguji: String(row[9] || "-"),
      kelancaran: String(row[10] || "-"),
      makhraj: String(row[11] || "-"),
      harakat: String(row[12] || "-"),
      tajwid: String(row[13] || "-"),
      adab: String(row[14] || "-"),
      keputusan: String(row[19] || "-"), // Kolom T
      nilaiAkhir: String(row[20] || "-") // Kolom U
    });
  }
  return rekapan.reverse(); 
}

// 3. Fungsi Menyimpan Jadwal (Disesuaikan letak kolomnya)
function serverSimpanJadwalPenguji(index, tanggalBaru, namaPenguji) {
  if (!index) throw new Error("Nomor baris (index) kosong!");
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pengajuan");
  if (!sheet) return;
  
  // Karena kolom bergeser, sekarang kita update di kolom 7 (G) dan 8 (H)
  sheet.getRange(index, 7).setValue(tanggalBaru); // Kolom G (Tgl Pelaksanaan)
  sheet.getRange(index, 8).setValue(namaPenguji); // Kolom H (Nama Penguji)
}
