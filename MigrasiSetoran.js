// ====================================================================
// FUNGSI MIGRASI SETORAN (DENGAN BATCH DAN KOLOM S)
// ====================================================================
function runMigrasiSetoranBatch() {
  var db = getFirestoreDB(); // Pastikan library/fungsi ini mendukung batch atau Anda sesuaikan
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setoran");
  var data = sheet.getDataRange().getValues();
  
  Logger.log("Memulai migrasi " + (data.length - 1) + " data setoran...");

  var batchData = [];
  var batchLimit = 1;

  for (var i = 10895; i < data.length; i++) {
    var idQiroah = data[i][0].toString().trim(); // Kolom A
    
    if (!idQiroah) continue;
    
    // Siapkan struktur payload
    var payloadFirestore = {
      tanggal: data[i][1],             // Kolom B
      unit: data[i][2],                // Kolom C
      id_siswa: data[i][3].toString(), // Kolom D
      nama_siswa: data[i][4],          // Kolom E
      ustadz: data[i][5],              // Kolom F
      jilid: data[i][6],               // Kolom G
      halaman: data[i][7],             // Kolom H
      status_lancar: data[i][8],       // Kolom I
      catatan: data[i][9],             // Kolom J
      tugas_rumah: data[i][10],        // Kolom K
      id_audio: data[i][11],           // Kolom L
      status_setoran: data[i][12],     // Kolom M 
      nilai: data[i][13],              // Kolom N 
      catatan_audio: data[i][14],      // Kolom O 
      murojaah_di_rumah: data[i][15],  // Kolom P 
      membaca_di_sekolah: data[i][16], // Kolom Q 
      murojaah_bersama_teman: data[i][17], // Kolom R
      nilai_audio: data[i][18]         // Kolom S (TERBARU)
    };

    // Tambahkan ke keranjang batch
    batchData.push({
      id: idQiroah,
      data: payloadFirestore
    });

    // Jika keranjang sudah isi 100, KIRIM!
    if (batchData.length >= batchLimit || i === data.length - 1) {
      try {
        // PERHATIAN: Perintah 'batchCreate' bergantung pada library Firestore yang Anda pakai.
        // Jika tidak ada batch, ubah menjadi looping db.createDocument biasa di dalam blok ini
        
        // --- CONTOH JIKA LIBRARY ANDA MENDUKUNG BATCH ---
        // db.batchCreate("Setoran", batchData); 
        
        // --- JIKA LIBRARY TIDAK MENDUKUNG BATCH, GUNAKAN LOOPING BIASA (MASIH ADA RISIKO TIMEOUT) ---
        for (var k = 0; k < batchData.length; k++) {
            db.createDocument("Setoran/" + batchData[k].id, batchData[k].data);
        }

        Logger.log("Berhasil memigrasi baris s.d " + (i + 1));
      } catch (e) {
        Logger.log("Gagal migrasi pada blok baris " + (i + 1) + ". Error: " + e.message);
      }
      batchData = []; // Kosongkan keranjang setelah kirim
    }
  }
  
  Logger.log("=== PROSES MIGRASI SETORAN SELESAI ===");
}