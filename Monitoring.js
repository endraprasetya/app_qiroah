function getMonitoringUstadz(tglMulaiStr, tglAkhirStr, userRole, userUnit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUsers = ss.getSheetByName("Users");
    if (!sheetUsers) return { sukses: false, pesan: "Gagal: Sheet 'Users' tidak ditemukan!" };
    
    var dataUsers = sheetUsers.getDataRange().getValues();
    var listUstadz = [];
    var setoranMap = {}; 
    
    var roleAkses = userRole ? userRole.toString().toLowerCase().trim() : "";
    var unitAkses = userUnit ? userUnit.toString().toLowerCase().trim() : "";
    
    // PERBAIKAN LOGIKA: Admin Pusat hanya jika spesifik adminpusat atau unit tertentu
    var isAdminPusat = (roleAkses.includes("adminpusat") || unitAkses === "semua" || unitAkses === "pusat"); 
    
    for (var i = 1; i < dataUsers.length; i++) {
      var roleDb = (dataUsers[i][2] != null ? dataUsers[i][2].toString().toLowerCase().trim() : "");
      var namaLengkap = (dataUsers[i][4] != null ? dataUsers[i][4].toString().trim() : "");
      var unitDb = (dataUsers[i][5] != null ? dataUsers[i][5].toString().trim() : "");
      
      if (namaLengkap === "") continue; 
      
      // Filter hanya untuk Ustadz (Bukan Admin/Kepsek)
      if (!roleDb.includes("admin") && !roleDb.includes("kepala")) {
        
        // JIKA BUKAN ADMIN PUSAT, LAKUKAN FILTER UNIT
        if (!isAdminPusat) {
           var unitDbLow = unitDb.toLowerCase();
           // Pastikan unit di DB sama persis dengan unit yang sedang login
           if (unitDbLow !== unitAkses) {
               continue; // Lewati jika unit tidak cocok
           }
        }

        var keyNama = namaLengkap.toLowerCase().replace(/[^a-z0-9]/g, '');
        listUstadz.push({
          nama_lengkap: namaLengkap,
          unit: unitDb || "-",
          jumlah_input: 0,
          status: "BELUM INPUT"
        });
        
        setoranMap[keyNama] = listUstadz.length - 1; 
      }
    }

    if (listUstadz.length === 0) {
      return { sukses: true, total_ustadz_aktif: 0, total_setoran: 0, data: [] };
    }
    
    // ... (Lanjutkan sisa kode perhitungan dataSetoran Anda seperti biasa) ...
    // Pastikan kode di bawahnya tetap sama karena logikanya sudah benar
    var totalSetoranDalamPeriode = 0;
    var targetMulai = tglMulaiStr.trim(); 
    var targetAkhir = tglAkhirStr.trim();
    
    var sheetSetoran = ss.getSheetByName("Setoran"); 
    if (sheetSetoran) {
      var dataSetoran = sheetSetoran.getDataRange().getValues();
      for (var j = 1; j < dataSetoran.length; j++) {
        var tglCell = dataSetoran[j][1];
        var namaUstadzDbRaw = dataSetoran[j][5];
        if (!tglCell || !namaUstadzDbRaw) continue;
        
        var keyUstadzDb = namaUstadzDbRaw.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); 
        var tglSetoranStr = "";
        
        if (Object.prototype.toString.call(tglCell) === '[object Date]') {
           tglSetoranStr = Utilities.formatDate(tglCell, "Asia/Jakarta", "yyyy-MM-dd");
        } else {
           tglSetoranStr = tglCell.toString().split(' ')[0];
        }
        
        if (tglSetoranStr >= targetMulai && tglSetoranStr <= targetAkhir) {
          if (setoranMap[keyUstadzDb] !== undefined) {
            var indexData = setoranMap[keyUstadzDb];
            listUstadz[indexData].jumlah_input += 1;
            listUstadz[indexData].status = "AKTIF";
            totalSetoranDalamPeriode++;
          }
        }
      }
    }
    
    var ustadzAktifCount = 0;
    for (var k = 0; k < listUstadz.length; k++) {
      if (listUstadz[k].jumlah_input > 0) ustadzAktifCount++;
    }
    
    listUstadz.sort(function(a, b) { return b.jumlah_input - a.jumlah_input; });
    
    return {
      sukses: true,
      total_ustadz_aktif: ustadzAktifCount,
      total_setoran: totalSetoranDalamPeriode,
      data: listUstadz
    };
  } catch (e) {
    return { sukses: false, pesan: "Error GAS: " + e.message };
  }
}

// ====================================================================
// FUNGSI BARU: AMBIL REKAP DETAIL 1 USTADZ UNTUK MODAL/GRAFIK
// ====================================================================
function getRekapDetailUstadzServer(namaUstadz, tglMulaiStr, tglAkhirStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSetoran = ss.getSheetByName("Setoran");
    if (!sheetSetoran) return { sukses: false, pesan: "Sheet Setoran tidak ditemukan!" };
    
    // 1. AMBIL DATA KELAS DARI SHEET "SISWA" (Untuk cek target)
    var sheetSiswa = ss.getSheetByName("Siswa");
    var mapKelas = {}; // Wadah menyimpan pasangan ID Siswa -> Kelas
    if (sheetSiswa) {
       var dataSiswa = sheetSiswa.getDataRange().getValues();
       var headers = dataSiswa[0] || [];
       var idxId = -1, idxKelas = -1;
       
       // Cari otomatis posisi kolom ID Siswa dan Kelas
       for(var c=0; c<headers.length; c++) {
          var h = headers[c].toString().toLowerCase().trim();
          if(h === "id siswa" || h === "id_siswa") idxId = c;
          if(h === "kelas") idxKelas = c;
       }
       if (idxId !== -1 && idxKelas !== -1) {
          for(var r=1; r<dataSiswa.length; r++) {
             if (dataSiswa[r][idxId]) mapKelas[dataSiswa[r][idxId].toString().trim()] = dataSiswa[r][idxKelas].toString();
          }
       }
    }

    // 2. PROSES DATA SETORAN
    var data = sheetSetoran.getDataRange().getValues();
    var hasil = [];
    var targetTercapai = 0;
    var targetBelum = 0;
    
    var targetMulai = tglMulaiStr.trim(); 
    var targetAkhir = tglAkhirStr.trim();
    var ustadzDicari = namaUstadz.toLowerCase().replace(/[^a-z0-9]/g, '');

    for (var i = 1; i < data.length; i++) {
      var tglCell = data[i][1]; // Kolom B: Tanggal
      var namaUstadzDbRaw = data[i][5]; // Kolom F: Ustadz
      
      if (!tglCell || !namaUstadzDbRaw) continue;
      var keyUstadzDb = namaUstadzDbRaw.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
      
      if (keyUstadzDb === ustadzDicari) {
         var tglSetoranStr = "";
         if (Object.prototype.toString.call(tglCell) === '[object Date]') {
            tglSetoranStr = Utilities.formatDate(tglCell, "Asia/Jakarta", "yyyy-MM-dd");
         } else {
            tglSetoranStr = tglCell.toString().split(' ')[0];
         }
         
         if (tglSetoranStr >= targetMulai && tglSetoranStr <= targetAkhir) {
            
            // PERBAIKAN URUTAN KOLOM SESUAI SCREENSHOT
            var idSiswa = (data[i][3] || "").toString().trim(); // Kolom D
            var namaSiswa = (data[i][4] || "-").toString(); // Kolom E
            var jilid = (data[i][6] || "").toString(); // Kolom G
            var halaman = (data[i][7] || "").toString(); // Kolom H
            var statusLancar = (data[i][8] || "").toString(); // Kolom I (Lancar/Diulang)
            var catatan = (data[i][9] || "").toString(); // Kolom J (Catatan)
            
            var capaianAkhir = jilid + (halaman ? " Hal " + halaman : "");
            var kelasSiswa = mapKelas[idSiswa] || ""; // Ambil kelas dari Map
            
            // 3. GUNAKAN LOGIKA TARGET ANDA
            var isTercapai = false;
            if (kelasSiswa && jilid) {
                isTercapai = cekTargetQiroahGS(kelasSiswa, jilid);
            }
            
            var statusTeks = "";
            if (isTercapai) {
              targetTercapai++;
              statusTeks = "Mencapai Target";
            } else {
              targetBelum++;
              statusTeks = "Belum Mencapai Target";
            }
            
            hasil.push({
              tgl: tglSetoranStr,
              siswa: namaSiswa,
              ustadz: namaUstadzDbRaw.toString(),
              capaian: capaianAkhir,
              catatan: catatan,
              statusLancar: statusLancar,
              statusTarget: statusTeks
            });
         }
      }
    }
    hasil.reverse();
    return {
      sukses: true,
      nama: namaUstadz,
      tercapai: targetTercapai,
      belum: targetBelum,
      total: targetTercapai + targetBelum,
      data: hasil
    };
    
  } catch(e) {
    return { sukses: false, pesan: e.message };
  }
}

// LOGIKA TARGET (SAMA PERSIS DENGAN JS ANDA, TAPI UNTUK SERVER)
function cekTargetQiroahGS(kelas, capaianStr) {
  if (!capaianStr || !kelas) return false;
  var klsNum = parseInt(kelas.toString().replace(/[^0-9]/g, '')) || 0;
  var teksCapaian = capaianStr.toString().toLowerCase();
  var bobot = 0;
  
  if (teksCapaian.includes("al quran") || teksCapaian.includes("al-quran") || teksCapaian.includes("al-qur'an") || teksCapaian.includes("al qur'an") || teksCapaian.includes("alquran")) {
    bobot = 7;
  } else if (teksCapaian.includes("at tartil 6")) { bobot = 6; }
  else if (teksCapaian.includes("at tartil 5")) { bobot = 5; }
  else if (teksCapaian.includes("at tartil 4")) { bobot = 4; }
  else if (teksCapaian.includes("at tartil 3")) { bobot = 3; }
  else if (teksCapaian.includes("at tartil 2")) { bobot = 2; }
  else if (teksCapaian.includes("at tartil 1")) { bobot = 1; }

  if (klsNum === 1) return (bobot >= 1);
  else if (klsNum === 2) return (bobot >= 4);
  else if (klsNum >= 3) return (bobot >= 7);
  return false;
}

// ====================================================================
// FUNGSI PERKEMBANGAN KELAS (MENGGUNAKAN LOGIKA CEK TARGET QIROAH)
// ====================================================================
function getPerkembanganKelasServer(kelasTarget, userUnit) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("Siswa");
    var sheetSetoran = ss.getSheetByName("Setoran");
    
    if (!sheetSiswa || !sheetSetoran) return { sukses: false, pesan: "Sheet Siswa atau Setoran tidak ditemukan!" };
    
    var targetKls = (kelasTarget || "").toString().toLowerCase().trim();
    var targetUnit = (userUnit || "").toString().toLowerCase().trim();
    
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    if (dataSiswa.length <= 1) return { sukses: true, data: [] };
    
    var headersSiswa = dataSiswa[0];
    var idxId = -1, idxNama = -1, idxKelas = -1, idxUnit = -1;
    
    // 1. Deteksi Header Kolom
    for (var c = 0; c < headersSiswa.length; c++) {
      var h = headersSiswa[c].toString().toLowerCase().trim();
      if (h.indexOf("id") !== -1 && idxId === -1) idxId = c;
      if (h.indexOf("nama") !== -1 && idxNama === -1) idxNama = c;
      if (h.indexOf("kelas") !== -1 && idxKelas === -1) idxKelas = c;
      if (h.indexOf("unit") !== -1 && idxUnit === -1) idxUnit = c;
    }
    
    if (idxNama === -1) idxNama = 1;  
    if (idxKelas === -1) idxKelas = 2; 
    
    var listSiswa = [];
    var mapSiswaKey = {}; 
    
    // 2. Filter Siswa berdasarkan Kelas & Unit
    for (var r = 1; r < dataSiswa.length; r++) {
      var namaSiswa = (dataSiswa[r][idxNama] || "").toString().trim();
      if (!namaSiswa) continue; 
      
      var klsDb = (idxKelas !== -1 && dataSiswa[r][idxKelas] != null ? dataSiswa[r][idxKelas].toString().toLowerCase().trim() : "");
      var unitDb = (idxUnit !== -1 && dataSiswa[r][idxUnit] != null ? dataSiswa[r][idxUnit].toString().toLowerCase().trim() : "");
      
      if (klsDb === targetKls) {
        if (targetUnit && !targetUnit.includes("semua") && !targetUnit.includes("pusat") && unitDb) {
          if (unitDb !== targetUnit && !unitDb.includes(targetUnit)) {
            continue; 
          }
        }
        
        var idSiswa = (idxId !== -1 && dataSiswa[r][idxId] != null ? dataSiswa[r][idxId].toString().trim() : "");
        var keyUnik = idSiswa ? idSiswa : namaSiswa.toLowerCase().replace(/[^a-z0-9]/g, '');
        
        var objSiswa = {
          nama: namaSiswa,
          awal: "Belum Setor",
          akhir: "Belum Setor",
          akhir_tercapai: false,
          sudahAdaAwal: false
        };
        
        listSiswa.push(objSiswa);
        mapSiswaKey[keyUnik] = listSiswa.length - 1;
      }
    }
    
    if (listSiswa.length === 0) return { sukses: true, data: [] };
    
    // 3. Ambil Setoran & Hitung Ketercapaian Target Menggunakan cekTargetQiroahGS()
    var dataSetoran = sheetSetoran.getDataRange().getValues();
    for (var i = 1; i < dataSetoran.length; i++) {
      var row = dataSetoran[i];
      var idS = (row[3] || "").toString().trim();   
      var namaS = (row[4] || row[2] || "").toString().trim(); 
      
      var keySetoran = idS ? idS : namaS.toLowerCase().replace(/[^a-z0-9]/g, '');
      
      if (mapSiswaKey[keySetoran] !== undefined) {
        var idxInList = mapSiswaKey[keySetoran];
        var jilid = (row[6] || "").toString().trim();  
        var halaman = (row[7] || "").toString().trim(); 
        var capaian = jilid + (halaman ? " Hal " + halaman : "");
        
        if (!capaian) continue;
        
        // Simpan setoran awal
        if (!listSiswa[idxInList].sudahAdaAwal) {
          listSiswa[idxInList].awal = capaian;
          listSiswa[idxInList].sudahAdaAwal = true;
        }
        
        // Simpan setoran terakhir
        listSiswa[idxInList].akhir = capaian;
        
        // PASTI DAN AKURAT: Cek target menggunakan fungsi cekTargetQiroahGS milik Anda
        listSiswa[idxInList].akhir_tercapai = cekTargetQiroahGS(kelasTarget, capaian);
      }
    }
    
    return { sukses: true, data: listSiswa };
    
  } catch(e) {
    return { sukses: false, pesan: "Error Server: " + e.message };
  }
}