const SPREADSHEET_ID = '1J_q7PBRkRF5JaEtbadxMrO6NbnYFue02P5SNEEEjaSo';

// --- FUNGSI TAMBAH BARANG (PART MASTER) ---
function simpanBarangBaru(data) {
  try {
    const ss = SpreadsheetApp.openById('1J_q7PBRkRF5JaEtbadxMrO6NbnYFue02P5SNEEEjaSo');
    const sheet = ss.getSheetByName('PART MASTER');
    
    const rowData = [
      "'" + data.kodeBarang, 
      data.namaBarang,
      data.kategori,
      data.unit,
      "", // Kolom Stock dikosongkan (diisi otomatis oleh rumus/DATA ENTRY)
      data.hargaJual,
      data.hargaBeli,
      ""  // Kolom Margin dikosongkan (diisi otomatis oleh rumus)
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: "Data barang berhasil ditambahkan ke PART MASTER!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// --- FUNGSI TAMBAH PELANGGAN ---
function simpanPelangganBaru(namaPelanggan) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('CUST');
    
    // Sheet CUST Anda hanya berisi 1 kolom (Nama)
    sheet.appendRow([namaPelanggan]);
    return { success: true, message: "Pelanggan berhasil ditambahkan!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// --- FUNGSI SIMPAN TRANSAKSI (PENJUALAN/PEMBELIAN/RETUR) ---
// Gunakan fungsi ini untuk SEMUA Form Transaksi Anda
function simpanTransaksi(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DATA ENTRY');
    
    // Susun data sesuai urutan kolom DATA ENTRY: 
    // Tanggal, Nomor Nota, Customer, No Barang, Nama Barang, Kategori, Sat, Harga, Qty, Kode Mutasi, Keterangan
    
    // Penting: Kode Mutasi disesuaikan dengan form yang mengirim (1=Masuk, 2=Keluar, 3=Rusak)
    const rowData = [
      data.tanggal, 
      data.nomorNota,
      data.customer,
      data.kodeBarang,
      data.namaBarang,
      data.kategori,
      data.satuan,
      data.harga,
      data.qty,
      data.kodeMutasi, 
      data.keterangan
    ];
    
    sheet.appendRow(rowData);
    return { success: true, message: "Transaksi berhasil disimpan ke Data Entry!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}