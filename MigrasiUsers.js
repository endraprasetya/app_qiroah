function runMigrasiUsers() {
  var db = getFirestoreDB(); // Memanggil koneksi dari FirebaseConfig.gs
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  
  Logger.log("Memulai migrasi " + (data.length - 1) + " data Users...");

  // Looping mulai dari baris kedua (indeks 1) untuk melewati header
  for (var i = 1; i < data.length; i++) {
    var username = data[i][0].toString().trim(); // Kolom A (Indeks 0)
    
    // Jika Username kosong, lewati baris ini
    if (!username) continue;
    
    // Memetakan data sesuai urutan kolom di Sheet Users terbaru
    var dataUser = {
      username: username,                           // Kolom A (Indeks 0)
      password: data[i][1].toString().trim(),       // Kolom B (Indeks 1)
      role: data[i][2].toString().trim(),           // Kolom C (Indeks 2)
      target_id: data[i][3].toString().trim(),      // Kolom D (Indeks 3)
      nama_lengkap: data[i][4].toString().trim(),   // Kolom E (Indeks 4)
      unit: data[i][5].toString().trim(),           // Kolom F (Indeks 5)
      status_setup: data[i][6].toString().trim()    // Kolom G (Indeks 6)
    };
    
    try {
      // Menyimpan ke collection 'Users' dengan nama dokumen menggunakan Username
      db.createDocument("Users/" + username, dataUser);
      Logger.log("Berhasil memigrasi User: " + dataUser.username + " (" + dataUser.nama_lengkap + ")");
    } catch (e) {
      Logger.log("Gagal memigrasi Username " + username + " | Error: " + e.message);
    }
  }
  
  Logger.log("=== PROSES MIGRASI USERS SELESAI ===");
}