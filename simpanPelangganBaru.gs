// --- FUNGSI TAMBAH PELANGGAN ---
function simpanPelangganBaru(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('CUST'); // Pastikan nama sheet sesuai
    
    // Pecah objek 'data' menjadi array agar masuk ke kolom yang terpisah
    // Sesuaikan urutannya dengan susunan kolom di sheet CUST Anda
    const rowData = [
      data.namaPelanggan, // Akan masuk ke Kolom A
      data.noHp,          // Akan masuk ke Kolom B
      data.alamat         // Akan masuk ke Kolom C
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: "Pelanggan berhasil ditambahkan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}