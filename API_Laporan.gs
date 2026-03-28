// --- FUNGSI AMBIL DATA STOCK MASTER ---
function ambilDataStock() {
  try {
    const ss = SpreadsheetApp.openById('1J_q7PBRkRF5JaEtbadxMrO6NbnYFue02P5SNEEEjaSo');
    const sheet = ss.getSheetByName('STOCK MASTER'); 
    const data = sheet.getDataRange().getValues();
    let hasilStock = [];
    
    // Mulai dari baris ke-2 (index 1) untuk melewati header
    for (let i = 1; i < data.length; i++) {
      // Jika kolom Nama Barang kosong, lewati baris ini
      if (!data[i][1]) continue; 

      hasilStock.push({
        namaBarang: data[i][1], // Kolom B (Index 1)
        masuk: data[i][5],      // Kolom F (Index 5)
        keluar: data[i][6],     // Kolom G (Index 6)
        rusak: data[i][7],      // Kolom H (Index 7)
        tersedia: data[i][8]    // Kolom I (Index 8)
      });
    }
    
    return hasilStock;
  } catch (error) {
    return []; 
  }
}
