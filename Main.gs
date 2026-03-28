function doGet(e) {
  // Memuat file Index.html sebagai template utama
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Manajemen Penjualan')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi krusial untuk memanggil file HTML lain ke dalam Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
