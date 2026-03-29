function simpanPenjualan(formObject) {
  try {
    // 1. Hubungkan ke Spreadsheet dan pilih sheet 'data entry'
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data entry");
    
    // Pastikan sheet 'data entry' ditemukan
    if (!sheet) {
      throw new Error("Sheet 'data entry' tidak ditemukan. Pastikan namanya sama persis.");
    }

    // 2. Buat ID Transaksi unik
    var idTransaksi = "TRX-" + new Date().getTime();
    var waktu = new Date();
    
    // 3. Ambil data dari form HTML 
    var namaBarang = formObject.namaBarang;
    var jumlah = formObject.jumlah;
    
    // 4. Masukkan data ke baris baru di sheet 'data entry'
    // Asumsi urutan kolom: A=ID Transaksi, B=Waktu, C=Nama Barang, D=Jumlah
    sheet.appendRow([idTransaksi, waktu, namaBarang, jumlah]);
    
    // 5. Kembalikan ID transaksi ke HTML
    return idTransaksi;
    
  } catch (error) {
    throw new Error(error.message);
  }
}