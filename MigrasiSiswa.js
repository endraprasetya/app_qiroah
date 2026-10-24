function runMigrasiSiswa() {
  var db = getFirestoreDB(); // Memanggil fungsi koneksi dari FirebaseConfig.gs
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Siswa");
  var data = sheet.getDataRange().getValues();
  
  Logger.log("Memulai migrasi " + (data.length - 1) + " data Siswa...");

  // Looping mulai dari baris kedua (indeks 1) untuk melewati header
  for (var i = 1368; i < data.length; i++) {
    var idSiswa = data[i][0].toString().trim(); // Kolom A: ID_Siswa sebagai Document ID
    
    // Jika ID Siswa kosong, lewati baris ini (mencegah error baris kosong)
    if (!idSiswa) continue; 
    
    // Pemetaan data berdasarkan Screenshot (332)_2.jpg
    var dataSiswa = {
      id_siswa: idSiswa,                  // Kolom A (index 0)
      nama_siswa: data[i][1],             // Kolom B (index 1)
      id_ortu: data[i][2],                // Kolom C (index 2)
      kelas: data[i][3].toString(),       // Kolom D (index 3)
      id_pembimbing: data[i][4],          // Kolom E (index 4)
      no_wa_ortu: data[i][5].toString(),  // Kolom F (index 5) - Pastikan jadi string
      unit: data[i][6],                   // Kolom G (index 6)
      password_ortu: data[i][7].toString(),// Kolom H (index 7)
      nis: data[i][8] ? data[i][8].toString() : "" // Kolom I (index 8) - Jika kosong, diisi string kosong
    };
    
    try {
      // Menyimpan ke collection 'Siswa' dengan nama dokumen menggunakan ID_Siswa
      db.createDocument("Siswa/" + idSiswa, dataSiswa);
      Logger.log("Berhasil memigrasi Siswa: " + dataSiswa.nama_siswa + " (ID: " + idSiswa + ")");
    } catch (e) {
      Logger.log("Gagal memigrasi ID Siswa " + idSiswa + " | Error: " + e.message);
    }
  }
  
  Logger.log("=== PROSES MIGRASI SISWA SELESAI ===");
}