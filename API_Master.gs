const SPREADSHEET_ID = '1J_q7PBRkRF5JaEtbadxMrO6NbnYFue02P5SNEEEjaSo';

/**
 * MENGAMBIL DATA AWAL (NOTA & LIST PREDIKTIF)
 */
function getInitialData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Generate Next Nota (SHAZA000001)
  const sheetEntry = ss.getSheetByName('DATA ENTRY');
  const lastRow = sheetEntry.getLastRow();
  let nextNum = 1;
  
  if (lastRow > 1) {
    const lastNota = sheetEntry.getRange(lastRow, 2).getValue().toString();
    const numPart = lastNota.replace("SHAZA", "");
    if (!isNaN(parseInt(numPart))) {
      nextNum = parseInt(numPart) + 1;
    }
  }
  const formattedNota = "SHAZA" + nextNum.toString().padStart(6, '0');

  // 2. Data Customer (CUST Kolom A)
  const sheetCust = ss.getSheetByName('CUST');
  const custData = sheetCust.getRange("A2:A" + sheetCust.getLastRow()).getValues()
                   .map(r => r[0]).filter(i => i !== "");

  // 3. Data Barang (PART MASTER Kolom B)
  const sheetPart = ss.getSheetByName('PART MASTER');
  const partData = sheetPart.getRange("B2:B" + sheetPart.getLastRow()).getValues()
                   .map(r => r[0]).filter(i => i !== "");

  return {
    nextNota: formattedNota,
    customer: custData,
    barang: partData
  };
}

/**
 * SIMPAN KE DATA ENTRY (11 KOLOM)
 */
function simpanTransaksi(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DATA ENTRY');
    
    // Kolom D, F, G, H dikirim "" agar ARRAYFORMULA tidak rusak (REF!)
    const rowData = [
      data.tanggal,     // A
      data.nomorNota,   // B
      data.customer,    // C
      data.kodeBarang,  // D (Dikosongkan)
      data.namaBarang,  // E
      data.kategori,    // F (Dikosongkan)
      data.satuan,      // G (Dikosongkan)
      data.harga,       // H (Dikosongkan)
      data.qty,         // I
      data.kodeMutasi,  // J
      data.keterangan   // K
    ];
    
    sheet.appendRow(rowData);
    return true;
  } catch (e) {
    throw new Error("Gagal Simpan: " + e.message);
  }
}