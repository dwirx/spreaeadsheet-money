// Code.gs - Advanced Personal Finance Management System (FIXED)

// ===== MENU & INITIALIZATION =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Keuangan Pribadi Pro')
    .addItem('🚀 Setup Awal', 'setupComplete')
    .addItem('➕ Transaksi Baru', 'showTransactionDialog')
    .addItem('💳 Transfer Antar Wallet', 'showTransferDialog')
    .addItem('💸 Catat Utang/Piutang', 'showUtangPiutangDialog')
    .addItem('🎯 Atur Anggaran', 'showBudgetDialog') // New Menu Item
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Laporan')
      .addItem('📅 Laporan Mingguan', 'generateWeeklyReport')
      .addItem('📆 Laporan Bulanan', 'generateMonthlyReport')
      .addItem('📈 Laporan 2 Bulanan', 'generateBimonthlyReport')
      .addItem('🗓️ Laporan Semester', 'generateSemesterReport')
      .addItem('📋 Laporan Tahunan', 'generateYearlyReport'))
    .addSeparator()
    .addItem('🔧 Pengaturan', 'showSettings')
    .addItem('📱 Update Dashboard', 'updateDashboard')
    .addItem('🗑️ Reset Data', 'resetAllData')
    .addItem('❓ Bantuan', 'showHelp')
    .addToUi();
}

// ===== COMPLETE SETUP FUNCTION =====
function setupComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Konfirmasi setup
  const response = ui.alert(
    '🚀 Setup Template Keuangan',
    'Ini akan membuat template baru dan menghapus semua sheet yang ada. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Hapus semua sheet kecuali yang pertama
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // Setup sheets
  setupDashboard(sheets[0]);
  setupTransaksi(ss.insertSheet('Transaksi'));
  setupWallets(ss.insertSheet('Wallets'));
  setupKategori(ss.insertSheet('Kategori'));
  setupUtangPiutang(ss.insertSheet('UtangPiutang'));
  setupBudget(ss.insertSheet('Budget')); // New Sheet Setup
  setupLaporan(ss.insertSheet('Laporan'));
  setupPengaturan(ss.insertSheet('Pengaturan'));
  
  // Initialize default data
  initializeDefaultData();
  
  ui.alert('✅ Sukses!', 'Template berhasil dibuat! Silakan mulai dengan menambah transaksi.', ui.ButtonSet.OK);
}

// ===== DASHBOARD SETUP =====
function setupDashboard(sheet) {
  sheet.setName('Dashboard');
  sheet.clear();
  
  // Header
  sheet.getRange('A1:H1').merge();
  sheet.getRange('A1').setValue('💰 DASHBOARD KEUANGAN PRIBADI');
  sheet.getRange('A1').setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#1a73e8').setFontColor('white');
  
  // Current Date & Time
  sheet.getRange('A2:H2').merge();
  sheet.getRange('A2').setFormula('="Update Terakhir: " & TEXT(NOW(), "dd mmmm yyyy HH:mm")');
  sheet.getRange('A2').setHorizontalAlignment('center').setFontStyle('italic');
  
  // === SALDO WALLETS ===
  sheet.getRange('A4').setValue('💳 SALDO WALLETS');
  sheet.getRange('A4').setFontSize(14).setFontWeight('bold').setBackground('#E3F2FD');
  sheet.getRange('A4:D4').merge();
  
  // Wallet headers
  const walletHeaders = ['Wallet', 'Saldo', 'Update', 'Persentase'];
  sheet.getRange(5, 1, 1, 4).setValues([walletHeaders]);
  sheet.getRange(5, 1, 1, 4).setFontWeight('bold').setBackground('#BBDEFB');
  
  // Wallet data dengan formula yang diperbaiki
  const wallets = ['Bank BRI', 'Bank Jago', 'BSI', 'DANA', 'ShopeePay'];
  for (let i = 0; i < wallets.length; i++) {
    const row = 6 + i;
    sheet.getRange(row, 1).setValue(wallets[i]);
    
    // Formula saldo dipecah untuk menghindari error
    const formulaParts = [
      "SUMIFS(Transaksi!E:E,Transaksi!H:H,\"" + wallets[i] + "\",Transaksi!B:B,\"Pemasukan\")",
      "SUMIFS(Transaksi!E:E,Transaksi!H:H,\"" + wallets[i] + "\",Transaksi!B:B,\"Pengeluaran\")",
      "SUMIFS(Transaksi!E:E,Transaksi!H:H,\"" + wallets[i] + "\",Transaksi!B:B,\"Transfer Masuk\")",
      "SUMIFS(Transaksi!E:E,Transaksi!H:H,\"" + wallets[i] + "\",Transaksi!B:B,\"Transfer Keluar\")"
    ];
    
    const formula = "=" + formulaParts[0] + "-" + formulaParts[1] + "+" + formulaParts[2] + "-" + formulaParts[3];
    sheet.getRange(row, 2).setFormula(formula);
    
    // Formula update terakhir
    sheet.getRange(row, 3).setFormula('=IFERROR(TEXT(MAX(FILTER(Transaksi!A:A,Transaksi!H:H="' + wallets[i] + '")),"dd mmm"),"Belum ada")');
    
    // Formula persentase
    sheet.getRange(row, 4).setFormula('=IFERROR(B' + row + '/B12,0)');
  }
  
  // Total row
  sheet.getRange('A11:D11').setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A12').setValue('TOTAL');
  sheet.getRange('A12').setFontWeight('bold');
  sheet.getRange('B12').setFormula('=SUM(B6:B10)');
  sheet.getRange('D12').setValue('100%');
  
  // Format currency dan persentase
  sheet.getRange('B6:B12').setNumberFormat('"Rp "#,##0');
  sheet.getRange('D6:D12').setNumberFormat('0%');
  
  // === RINGKASAN BULANAN ===
  sheet.getRange('F4').setValue('📊 RINGKASAN BULAN INI');
  sheet.getRange('F4').setFontSize(14).setFontWeight('bold').setBackground('#E8F5E9');
  sheet.getRange('F4:H4').merge();
  
  const summaryLabels = [
    ['Pemasukan:', '=SUMIFS(Transaksi!E:E,Transaksi!B:B,"Pemasukan",Transaksi!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Transaksi!A:A,"<="&EOMONTH(TODAY(),0))'],
    ['Pengeluaran:', '=SUMIFS(Transaksi!E:E,Transaksi!B:B,"Pengeluaran",Transaksi!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),Transaksi!A:A,"<="&EOMONTH(TODAY(),0))'],
    ['Selisih:', '=G5-G6'],
    ['', ''],
    ['Rata-rata/hari:', '=IFERROR(G6/DAY(TODAY()),0)']
  ];
  
  for (let i = 0; i < summaryLabels.length; i++) {
    sheet.getRange(5 + i, 6).setValue(summaryLabels[i][0]);
    sheet.getRange(5 + i, 7).setFormula(summaryLabels[i][1]);
  }
  
  sheet.getRange('G5:G9').setNumberFormat('"Rp "#,##0');
  sheet.getRange('G7').setFontWeight('bold');
  
  // Conditional formatting untuk selisih
  const selisihRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#C8E6C9') // Green
    .setRanges([sheet.getRange('G7')])
    .build();
    
  const selisihNegativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#FFCDD2') // Red
    .setRanges([sheet.getRange('G7')])
    .build();
  
  // === BUDGET SUMMARY (BULAN INI) ===
  sheet.getRange('A37').setValue('📊 RINGKASAN BUDGET BULANAN');
  sheet.getRange('A37').setFontSize(14).setFontWeight('bold').setBackground('#E3F2FD');
  sheet.getRange('A37:D37').merge();
  const budgetHeaders = ['Total Anggaran', 'Total Pengeluaran', 'Sisa', 'Penggunaan'];
  sheet.getRange(38, 1, 1, 4).setValues([budgetHeaders]);
  sheet.getRange(38, 1, 1, 4).setFontWeight('bold').setBackground('#BBDEFB');
  sheet.getRange('A39').setFormula('=SUM(Budget!B:B)');
  sheet.getRange('B39').setFormula('=SUM(Budget!C:C)');
  sheet.getRange('C39').setFormula('=A39-B39');
  sheet.getRange('D39').setFormula('=IFERROR(B39/A39, 0)');
  sheet.getRange('A39:C39').setNumberFormat('"Rp "#,##0');
  sheet.getRange('D39').setNumberFormat('0%');

  const budgetSisaRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('C39')])
    .build();
  const budgetSisaNegativeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#FFCDD2')
    .setRanges([sheet.getRange('C39')])
    .build();

  // === ANALISIS KESEHATAN FINANSIAL (50/30/20) ===
  sheet.getRange('E27').setValue('❤️ ANALISIS KESEHATAN FINANSIAL (50/30/20)');
  sheet.getRange('E27').setFontSize(14).setFontWeight('bold').setBackground('#FCE4EC');
  sheet.getRange('E27:I27').merge();

  const analysisHeaders = [['Tipe', 'Alokasi Ideal', 'Alokasi Aktual', 'Jumlah', 'Status']];
  sheet.getRange(28, 5, 1, 5).setValues(analysisHeaders); // E28:I28
  sheet.getRange('E28:I28').setFontWeight('bold').setBackground('#F8BBD0');

  const analysisData = [
    ['Kebutuhan (Needs)', '50%'],
    ['Keinginan (Wants)', '30%'],
    ['Tabungan/Investasi', '20%']
  ];
  sheet.getRange(29, 5, 3, 2).setValues(analysisData); // E29:F31

  // Formulas for 50/30/20 analysis
  const totalIncomeFormula = 'G5'; // Total Pemasukan Bulan Ini
  sheet.getRange('G29').setFormula(`=H29/${totalIncomeFormula}`); // Aktual % Kebutuhan
  sheet.getRange('H29').setFormula('=SUM(SUMIFS(Transaksi!E:E, Transaksi!C:C, FILTER(Kategori!D:D, Kategori!F:F="Kebutuhan"), Transaksi!A:A, ">="&EOMONTH(TODAY(),-1)+1, Transaksi!A:A, "<="&EOMONTH(TODAY(),0)))');
  sheet.getRange('I29').setFormula('=IF(G29<=F29, "✅ Bagus", "⚠️ Evaluasi")');

  sheet.getRange('G30').setFormula(`=H30/${totalIncomeFormula}`); // Aktual % Keinginan
  sheet.getRange('H30').setFormula('=SUM(SUMIFS(Transaksi!E:E, Transaksi!C:C, FILTER(Kategori!D:D, Kategori!F:F="Keinginan"), Transaksi!A:A, ">="&EOMONTH(TODAY(),-1)+1, Transaksi!A:A, "<="&EOMONTH(TODAY(),0)))');
  sheet.getRange('I30').setFormula('=IF(G30<=F30, "✅ Bagus", "⚠️ Evaluasi")');

  sheet.getRange('G31').setFormula(`=H31/${totalIncomeFormula}`); // Aktual % Tabungan
  sheet.getRange('H31').setFormula(`=SUM(SUMIFS(Transaksi!E:E, Transaksi!C:C, FILTER(Kategori!D:D, Kategori!F:F="Tabungan/Investasi"), Transaksi!A:A, ">="&EOMONTH(TODAY(),-1)+1, Transaksi!A:A, "<="&EOMONTH(TODAY(),0))) + SUMIFS(UtangPiutang!D:D, UtangPiutang!A:A, "Piutang", UtangPiutang!G:G, "Lunas", UtangPiutang!E:E, ">="&EOMONTH(TODAY(),-1)+1)`);
  sheet.getRange('I31').setFormula('=IF(G31>=F31, "✅ Bagus", "⚠️ Tingkatkan")');

  // Formatting
  sheet.getRange('G29:G31').setNumberFormat('0%');
  sheet.getRange('H29:H31').setNumberFormat('"Rp "#,##0');
  sheet.getRange('E29:I31').setHorizontalAlignment('center');

  // Conditional formatting for Status
  const statusBagusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('✅')
    .setBackground('#C8E6C9')
    .setRanges([sheet.getRange('I29:I31')])
    .build();
  const statusEvaluasiRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️')
    .setBackground('#FFEBEE')
    .setRanges([sheet.getRange('I29:I31')])
    .build();

  // Rekomendasi Section
  sheet.getRange('E33').setValue('💡 REKOMENDASI UNTUK ANDA:');
  sheet.getRange('E33').setFontWeight('bold');
  sheet.getRange('E34').setFormula('=IF(I29="⚠️ Evaluasi", "Pengeluaran Kebutuhan Anda melebihi 50% dari pendapatan. Coba tinjau kembali pos-pos seperti \'Makanan\' atau \'Tagihan\' untuk efisiensi.", IF(I30="⚠️ Evaluasi", "Pengeluaran Keinginan Anda melebihi 30%. Pertimbangkan untuk mengurangi frekuensi \'Hiburan\' atau \'Belanja\' bulan depan.", IF(I31="⚠️ Tingkatkan", "Alokasi Tabungan Anda di bawah 20%. Coba sisihkan sebagian pendapatan di awal bulan atau cari sumber pemasukan tambahan.", "🏆 Selamat! Alokasi keuangan Anda sudah sangat baik. Pertahankan!")))');
  sheet.getRange('E34:I35').merge().setWrap(true).setVerticalAlignment('top').setFontStyle('italic');


  // === TOP KATEGORI ===
  sheet.getRange('A14').setValue('🏆 TOP 5 PENGELUARAN (BULAN INI)');
  sheet.getRange('A14').setFontSize(14).setFontWeight('bold').setBackground('#FFF3E0');
  sheet.getRange('A14:D14').merge();
  
  const kategoriHeaders = ['Kategori', 'Jumlah', 'Transaksi', '%'];
  sheet.getRange(15, 1, 1, 4).setValues([kategoriHeaders]);
  sheet.getRange(15, 1, 1, 4).setFontWeight('bold').setBackground('#FFE0B2');
  
  // === GRAFIK AREA ===
  sheet.getRange('F14').setValue('📈 GRAFIK & ANALISIS');
  sheet.getRange('F14').setFontSize(14).setFontWeight('bold').setBackground('#F3E5F5');
  sheet.getRange('F14:H14').merge();
  
  // Chart placeholder
  sheet.getRange('F15:H25').merge();
  sheet.getRange('F15').setValue('Grafik akan muncul setelah update dashboard');
  sheet.getRange('F15').setHorizontalAlignment('center').setVerticalAlignment('middle');

  // === RINGKASAN UTANG & PIUTANG ===
  sheet.getRange('F11').setValue('📊 RINGKASAN UTANG & PIUTANG');
  sheet.getRange('F11').setFontSize(12).setFontWeight('bold').setBackground('#FFFDE7');
  sheet.getRange('F11:H11').merge();
  const utangHeaders = [
    ['Total Utang Saya:', '=SUMIFS(UtangPiutang!D:D, UtangPiutang!A:A, "Utang", UtangPiutang!G:G, "Belum Lunas")'],
    ['Total Piutang (Uang Saya di Lain):', '=SUMIFS(UtangPiutang!D:D, UtangPiutang!A:A, "Piutang", UtangPiutang!G:G, "Belum Lunas")']
  ];
  sheet.getRange('F12:G13').setValues(utangHeaders);
  sheet.getRange('G12:G13').setNumberFormat('"Rp "#,##0');
  
  // === JATUH TEMPO TERDEKAT ===
  sheet.getRange('A29').setValue('🔔 JATUH TEMPO TERDEKAT (UTANG/PIUTANG)');
  sheet.getRange('A29').setFontSize(14).setFontWeight('bold').setBackground('#FFEBEE');
  sheet.getRange('A29:D29').merge();
  const jatuhTempoHeaders = ['Pihak', 'Deskripsi', 'Jumlah', 'Jatuh Tempo'];
  sheet.getRange(30, 1, 1, 4).setValues([jatuhTempoHeaders]);
  sheet.getRange(30, 1, 1, 4).setFontWeight('bold').setBackground('#FFCDD2');
  sheet.getRange('A31:D35').setFormula('=IFERROR(QUERY(UtangPiutang!B:F, "SELECT B, C, D, F WHERE G = \'Belum Lunas\' AND F IS NOT NULL ORDER BY F ASC LIMIT 5"), "")');
  sheet.getRange('C31:C35').setNumberFormat('"Rp "#,##0');
  sheet.getRange('D31:D35').setNumberFormat('dd mmmm yyyy');

  // === TRANSAKSI TERAKHIR ===
  sheet.getRange('A22').setValue('🕐 5 TRANSAKSI TERAKHIR');
  sheet.getRange('A22').setFontSize(14).setFontWeight('bold').setBackground('#E1F5FE');
  sheet.getRange('A22:D22').merge();
  
  const transHeaders = ['Tanggal', 'Deskripsi', 'Jumlah', 'Wallet'];
  sheet.getRange(23, 1, 1, 4).setValues([transHeaders]);
  sheet.getRange(23, 1, 1, 4).setFontWeight('bold').setBackground('#B3E5FC');
  
  // Formulas untuk transaksi terakhir
  for (let i = 0; i < 5; i++) {
    const row = 24 + i;
    const index = i + 1;
    sheet.getRange(row, 1).setFormula('=IFERROR(INDEX(SORT(Transaksi!A:A,1,FALSE),' + index + '),"")');
    sheet.getRange(row, 2).setFormula('=IFERROR(INDEX(SORT(Transaksi!A:K,1,FALSE),' + index + ',4),"")');
    sheet.getRange(row, 3).setFormula('=IFERROR(INDEX(SORT(Transaksi!A:K,1,FALSE),' + index + ',5),"")');
    sheet.getRange(row, 4).setFormula('=IFERROR(INDEX(SORT(Transaksi!A:K,1,FALSE),' + index + ',8),"")');
  }
  
  sheet.getRange('A24:A28').setNumberFormat('dd mmmm yyyy');
  sheet.getRange('C24:C28').setNumberFormat('"Rp "#,##0');
  
  // Set column widths
  sheet.setColumnWidths(1, 1, 120);
  sheet.setColumnWidths(2, 1, 100);
  sheet.setColumnWidths(3, 1, 100);
  sheet.setColumnWidths(4, 1, 80);
  sheet.setColumnWidths(5, 1, 20);
  sheet.setColumnWidths(6, 1, 120);
  sheet.setColumnWidths(7, 1, 100);
  sheet.setColumnWidths(8, 1, 100);
  sheet.setColumnWidths(9, 1, 100); // New column for 50/30/20
  
  // Apply conditional formatting rules
  sheet.setConditionalFormatRules([selisihRule, selisihNegativeRule, budgetSisaRule, budgetSisaNegativeRule, statusBagusRule, statusEvaluasiRule]);
  
  // Freeze header
  sheet.setFrozenRows(4);
}

// ===== TRANSAKSI SHEET SETUP =====
function setupTransaksi(sheet) {
  sheet.clear();
  
  const headers = [
    'Tanggal', 'Jenis', 'Kategori', 'Deskripsi', 'Jumlah',
    'Status', 'Metode', 'Wallet', 'Tags', 'Catatan', 'Input By'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#263238');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Set column widths
  const widths = [100, 100, 120, 200, 120, 80, 100, 100, 150, 200, 100];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  
  // Data validations
  const jenisValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pemasukan', 'Pengeluaran', 'Transfer Masuk', 'Transfer Keluar'], true)
    .build();
  sheet.getRange('B2:B').setDataValidation(jenisValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Lunas', 'Pending', 'Batal'], true)
    .build();
  sheet.getRange('F2:F').setDataValidation(statusValidation);
  
  const metodeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tunai', 'Transfer', 'Kartu Debit', 'E-Wallet', 'QRIS'], true)
    .build();
  sheet.getRange('G2:G').setDataValidation(metodeValidation);
  
  const walletValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bank BRI', 'Bank Jago', 'BSI', 'DANA', 'ShopeePay', 'Cash'], true)
    .build();
  sheet.getRange('H2:H').setDataValidation(walletValidation);
  
  // Number formats
  sheet.getRange('A2:A').setNumberFormat('dd mmmm yyyy');
  sheet.getRange('E2:E').setNumberFormat('"Rp "#,##0');
  
  // Conditional formatting
  // Pemasukan - hijau
  const pemasukanRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="Pemasukan"')
    .setBackground('#E8F5E9')
    .setRanges([sheet.getRange('A2:K')])
    .build();
    
  // Pengeluaran - merah muda
  const pengeluaranRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2="Pengeluaran"')
    .setBackground('#FFEBEE')
    .setRanges([sheet.getRange('A2:K')])
    .build();
    
  // Transfer - biru
  const transferRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=OR($B2="Transfer Masuk",$B2="Transfer Keluar")')
    .setBackground('#E3F2FD')
    .setRanges([sheet.getRange('A2:K')])
    .build();
    
  // Pending - kuning
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="Pending"')
    .setBackground('#FFF9C4')
    .setRanges([sheet.getRange('A2:K')])
    .build();
  
  sheet.setConditionalFormatRules([pemasukanRule, pengeluaranRule, transferRule, pendingRule]);
  
  // Freeze header
  sheet.setFrozenRows(1);
}

// ===== WALLETS SHEET SETUP =====
function setupWallets(sheet) {
  sheet.clear();
  
  sheet.getRange('A1').setValue('💳 DAFTAR WALLETS & REKENING');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:F1').merge();
  
  const headers = ['Nama Wallet', 'Jenis', 'No. Rekening/ID', 'Saldo Awal', 'Status', 'Keterangan'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');
  
  // Default wallets
  const defaultWallets = [
    ['Bank BRI', 'Bank', '', 0, 'Aktif', 'Rekening utama'],
    ['Bank Jago', 'Bank', '', 0, 'Aktif', 'Tabungan digital'],
    ['BSI', 'Bank', '', 0, 'Aktif', 'Bank Syariah'],
    ['DANA', 'E-Wallet', '', 0, 'Aktif', 'E-wallet utama'],
    ['ShopeePay', 'E-Wallet', '', 0, 'Aktif', 'Untuk belanja online'],
    ['Cash', 'Tunai', '-', 0, 'Aktif', 'Uang tunai']
  ];
  
  sheet.getRange(4, 1, defaultWallets.length, headers.length).setValues(defaultWallets);
  
  // Format
  sheet.getRange('D4:D').setNumberFormat('"Rp "#,##0');
  
  // Validation
  const jenisWalletValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Bank', 'E-Wallet', 'Tunai', 'Investasi'], true)
    .build();
  sheet.getRange('B4:B').setDataValidation(jenisWalletValidation);
  
  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Aktif', 'Non-Aktif'], true)
    .build();
  sheet.getRange('E4:E').setDataValidation(statusValidation);
  
  // Column widths
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 200);
}

// ===== KATEGORI SHEET SETUP =====
function setupKategori(sheet) {
  sheet.clear();
  
  sheet.getRange('A1').setValue('🏷️ KATEGORI TRANSAKSI');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:F1').merge(); // Extend merge area
  
  // Kategori Pemasukan
  sheet.getRange('A3').setValue('💰 KATEGORI PEMASUKAN');
  sheet.getRange('A3').setFontWeight('bold').setBackground('#C8E6C9');
  sheet.getRange('A3:B3').merge();
  
  const pemasukanKategori = [
    ['Gaji', 'Gaji bulanan'],
    ['Bonus', 'Bonus & insentif'],
    ['Freelance', 'Penghasilan freelance'],
    ['Investasi', 'Return investasi'],
    ['Hadiah', 'Hadiah & pemberian'],
    ['Lainnya', 'Pemasukan lainnya']
  ];
  
  sheet.getRange(4, 1, pemasukanKategori.length, 2).setValues(pemasukanKategori);
  
  // Kategori Pengeluaran
  sheet.getRange('D3').setValue('💸 KATEGORI PENGELUARAN');
  sheet.getRange('D3').setFontWeight('bold').setBackground('#FFCDD2');
  sheet.getRange('D3:F3').merge(); // Extend merge area for new header
  
  const pengeluaranHeaders = ['Kategori', 'Deskripsi', 'Tipe (50/30/20)'];
  sheet.getRange(4, 4, 1, 3).setValues([pengeluaranHeaders]).setFontWeight('bold');

  const pengeluaranKategori = [
    ['Makanan', 'Makan & minum', 'Kebutuhan'],
    ['Transportasi', 'Bensin, parkir, toll', 'Kebutuhan'],
    ['Belanja', 'Kebutuhan sehari-hari', 'Kebutuhan'],
    ['Tagihan', 'Listrik, air, internet', 'Kebutuhan'],
    ['Kesehatan', 'Obat & perawatan', 'Kebutuhan'],
    ['Hiburan', 'Entertainment', 'Keinginan'],
    ['Pendidikan', 'Kursus & buku', 'Keinginan'],
    ['Fashion', 'Pakaian & aksesoris', 'Keinginan'],
    ['Gadget', 'Elektronik', 'Keinginan'],
    ['Sosial', 'Hadiah & donasi', 'Keinginan'],
    ['Investasi', 'Tabungan & investasi', 'Tabungan/Investasi'],
    ['Darurat', 'Pengeluaran darurat', 'Kebutuhan'],
    ['Rumah', 'Perawatan rumah', 'Kebutuhan'],
    ['Lainnya', 'Pengeluaran lainnya', 'Keinginan']
  ];
  
  sheet.getRange(5, 4, pengeluaranKategori.length, 3).setValues(pengeluaranKategori);
  
  // Data validation for Tipe
  const tipeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Kebutuhan', 'Keinginan', 'Tabungan/Investasi'], true)
    .setAllowInvalid(false)
    .setHelpText('Pilih salah satu tipe: Kebutuhan (Needs), Keinginan (Wants), atau Tabungan/Investasi (Savings).')
    .build();
  sheet.getRange('F5:F').setDataValidation(tipeValidation);

  // Format
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 20);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 150); // Width for new Tipe column
}

// ===== BUDGET SHEET SETUP =====
function setupBudget(sheet) {
  sheet.clear();
  sheet.setName('Budget');

  sheet.getRange('A1').setValue('🎯 PERENCANAAN ANGGARAN');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:E1').merge();

  sheet.getRange('A2').setValue('Atur alokasi dana bulanan Anda di sini. Data "Aktual" akan terisi otomatis dari sheet Transaksi.');
  sheet.getRange('A2:E2').merge().setFontStyle('italic');

  const headers = ['Kategori', 'Anggaran Bulanan', 'Pengeluaran Aktual', 'Sisa Anggaran', 'Penggunaan (%)'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');

  // Get expense categories
  const kategoriSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kategori');
  const pengeluaranKategori = kategoriSheet.getRange('D5:D' + kategoriSheet.getLastRow()).getValues().flat().filter(String);

  // Populate categories and formulas
  pengeluaranKategori.forEach((kategori, index) => {
    const row = 5 + index;
    sheet.getRange(row, 1).setValue(kategori);
    sheet.getRange(row, 2).setValue(0); // Default budget
    sheet.getRange(row, 3).setFormula(`=SUMIFS(Transaksi!E:E, Transaksi!C:C, A${row}, Transaksi!B:B, "Pengeluaran", Transaksi!A:A, ">="&EOMONTH(TODAY(),-1)+1, Transaksi!A:A, "<="&EOMONTH(TODAY(),0))`);
    sheet.getRange(row, 4).setFormula(`=B${row}-C${row}`);
    sheet.getRange(row, 5).setFormula(`=IFERROR(C${row}/B${row}, 0)`);
  });

  // Formatting
  sheet.getRange('B:D').setNumberFormat('"Rp "#,##0');
  sheet.getRange('E:E').setNumberFormat('0.00%');
  sheet.setColumnWidths(1, 1, 150);
  sheet.setColumnWidths(2, 4, 130);

  // Conditional formatting for usage percentage
  const usageRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#FFCDD2') // Red for high usage
    .setGradientMidpointWithValue('#FFF9C4', SpreadsheetApp.InterpolationType.PERCENT, '75')
    .setGradientMinpoint('#C8E6C9') // Green for low usage
    .setRanges([sheet.getRange('E5:E')])
    .build();
  
  sheet.setConditionalFormatRules([usageRule]);
  sheet.setFrozenRows(4);
}

// ===== UTANG/PIUTANG SHEET SETUP =====
function setupUtangPiutang(sheet) {
  sheet.clear();
  sheet.setName('UtangPiutang');

  sheet.getRange('A1').setValue('💸 DAFTAR UTANG & PIUTANG');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A1:G1').merge();

  const headers = ['Jenis', 'Pihak Terkait', 'Deskripsi', 'Jumlah', 'Tanggal Catat', 'Jatuh Tempo', 'Status', 'Catatan'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length).setFontWeight('bold').setBackground('#FFFDE7');

  // Data validations
  const jenisValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Utang', 'Piutang'], true)
    .build();
  sheet.getRange('A4:A').setDataValidation(jenisValidation);

  const statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Belum Lunas', 'Lunas', 'Batal'], true)
    .build();
  sheet.getRange('G4:G').setDataValidation(statusValidation);

  // Formatting
  sheet.getRange('D4:D').setNumberFormat('"Rp "#,##0');
  sheet.getRange('E4:F').setNumberFormat('dd mmmm yyyy');

  // Conditional formatting
  const utangRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A4="Utang"')
    .setBackground('#FFEBEE') // Light red
    .setRanges([sheet.getRange('A4:H')])
    .build();
  
  const piutangRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A4="Piutang"')
    .setBackground('#E8F5E9') // Light green
    .setRanges([sheet.getRange('A4:H')])
    .build();

  const lunasRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G4="Lunas"')
    .setFontColor('#9E9E9E')
    .setStrikethrough(true)
    .setRanges([sheet.getRange('A4:H')])
    .build();

  const jatuhTempoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F4 < TODAY(), $G4 = "Belum Lunas", NOT(ISBLANK($F4)))')
    .setBackground('#F44336') // Red
    .setFontColor('white')
    .setRanges([sheet.getRange('A4:H')])
    .build();

  sheet.setConditionalFormatRules([utangRule, piutangRule, lunasRule, jatuhTempoRule]);
  sheet.setFrozenRows(3);
}


// ===== LAPORAN SHEET SETUP =====
function setupLaporan(sheet) {
  sheet.clear();
  
  sheet.getRange('A1').setValue('📊 LAPORAN KEUANGAN');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold');
  sheet.getRange('A1:H1').merge();
  
  sheet.getRange('A3').setValue('Pilih jenis laporan dari menu untuk generate laporan');
  sheet.getRange('A3').setFontStyle('italic');
  
  // Template area untuk laporan
  sheet.getRange('A5').setValue('PERIODE LAPORAN:');
  sheet.getRange('A5').setFontWeight('bold');
  
  sheet.getRange('A7').setValue('RINGKASAN:');
  sheet.getRange('A7').setFontWeight('bold');
  
  sheet.getRange('A15').setValue('DETAIL TRANSAKSI:');
  sheet.getRange('A15').setFontWeight('bold');
}

// ===== PENGATURAN SHEET SETUP =====
function setupPengaturan(sheet) {
  sheet.clear();
  
  sheet.getRange('A1').setValue('⚙️ PENGATURAN');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  
  const settings = [
    ['Nama Pengguna:', Session.getActiveUser().getEmail()],
    ['Mata Uang:', 'IDR (Rupiah)'],
    ['Format Tanggal:', 'DD/MM/YYYY'],
    ['Batas Pengeluaran Harian:', 500000],
    ['Notifikasi Email:', 'Tidak Aktif'],
    ['Backup Otomatis:', 'Aktif'],
    ['Tema Warna:', 'Default']
  ];
  
  sheet.getRange(3, 1, settings.length, 2).setValues(settings);
  sheet.getRange('B6').setNumberFormat('"Rp "#,##0');
  
  // Validation
  const notifValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Aktif', 'Tidak Aktif'], true)
    .build();
  sheet.getRange('B7').setDataValidation(notifValidation);
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
}

// ===== INITIALIZE DEFAULT DATA =====
function initializeDefaultData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kategoriSheet = ss.getSheetByName('Kategori');
  
  // Create named ranges for kategori
  const pemasukanRange = kategoriSheet.getRange('A4:A9');
  const pengeluaranRange = kategoriSheet.getRange('D5:D' + kategoriSheet.getLastRow());
  
  ss.setNamedRange('KategoriPemasukan', pemasukanRange);
  ss.setNamedRange('KategoriPengeluaran', pengeluaranRange);
}

// ===== BUDGET DIALOG =====
function showBudgetDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BudgetForm')
    .setWidth(650)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 Atur Anggaran Bulanan');
}

// ===== GET BUDGET DATA (FOR FORM) =====
function getBudgetData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Budget');
    const dataRange = sheet.getRange('A5:B' + sheet.getLastRow());
    const values = dataRange.getValues();
    
    return values.map(row => ({
      kategori: row[0],
      anggaran: row[1]
    })).filter(item => item.kategori);
  } catch (e) {
    return []; // Return empty array on error
  }
}

// ===== UPDATE BUDGET =====
function updateBudget(budgetData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Budget');
    const categories = sheet.getRange('A5:A' + sheet.getLastRow()).getValues().flat();
    
    const newBudgets = categories.map(category => {
      if (budgetData.hasOwnProperty(category)) {
        return [parseFloat(budgetData[category])];
      }
      return [0]; // Default to 0 if not found
    });

    sheet.getRange('B5:B' + (4 + newBudgets.length)).setValues(newBudgets);
    
    // Update dashboard to reflect changes
    updateDashboard();

    return { success: true, message: 'Anggaran berhasil diperbarui!' };
  } catch (error) {
    Logger.log(error);
    return { success: false, message: 'Error: ' + error.toString() };
  }
}


// ===== TRANSACTION DIALOG =====
function showTransactionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TransactionForm')
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Tambah Transaksi Baru');
}

// ===== UTANG/PIUTANG DIALOG =====
function showUtangPiutangDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UtangPiutangForm')
    .setWidth(500)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '💸 Catat Utang/Piutang Baru');
}

// ===== ADD UTANG/PIUTANG FUNCTION =====
function addUtangPiutang(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('UtangPiutang');
    const newRow = sheet.getLastRow() + 1;
    const cleanJumlah = data.jumlah.replace(/\./g, ''); // Remove dots

    const rowData = [
      data.jenis,
      data.pihak,
      data.deskripsi,
      parseFloat(cleanJumlah),
      data.tanggalCatat ? new Date(data.tanggalCatat) : new Date(), // Use provided date or today
      data.tanggalJatuhTempo ? new Date(data.tanggalJatuhTempo) : null,
      'Belum Lunas',
      data.catatan || ''
    ];

    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.sort(5, false); // Sort by Tanggal Catat descending

    const transaksiSheet = ss.getSheetByName('Transaksi');

    // If it's a Piutang and a wallet is selected, add an income transaction
    if (data.jenis === 'Piutang' && data.wallet) {
      const pemasukanData = [
        new Date(data.tanggalCatat),
        'Pemasukan',
        'Piutang', // Or a more specific category
        'Penerimaan piutang dari ' + data.pihak,
        parseFloat(cleanJumlah),
        'Lunas',
        'Transfer',
        data.wallet,
        'piutang',
        'Otomatis dari pencatatan piutang',
        Session.getActiveUser().getEmail()
      ];
      transaksiSheet.appendRow(pemasukanData);
    }

    // If it's an Utang and a source wallet is selected, add an expense transaction
    if (data.jenis === 'Utang' && data.walletUtang) {
        const pengeluaranData = [
            new Date(data.tanggalCatat),
            'Pengeluaran',
            'Utang', // Or a more specific category
            'Dana untuk utang ke ' + data.pihak,
            parseFloat(cleanJumlah),
            'Lunas',
            'Transfer',
            data.walletUtang,
            'utang',
            'Otomatis dari pencatatan utang',
            Session.getActiveUser().getEmail()
        ];
        transaksiSheet.appendRow(pengeluaranData);
    }
    
    // Sort transactions sheet after potential additions and flush
    if (transaksiSheet.getLastRow() > 1) {
      const dataRange = transaksiSheet.getRange(2, 1, transaksiSheet.getLastRow() - 1, transaksiSheet.getLastColumn());
      dataRange.sort({column: 1, ascending: false});
    }
    
    updateDashboard();
    SpreadsheetApp.flush(); // Force immediate update

    return { success: true, message: 'Data utang/piutang berhasil ditambahkan!' };
  } catch (error) {
    Logger.log(error);
    return { success: false, message: 'Error: ' + error.toString() };
  }
}


// ===== ADD TRANSACTION FUNCTION =====
function addTransaction(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transaksi');
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    const cleanJumlah = data.jumlah.replace(/\./g, ''); // Remove dots
    
    const rowData = [
      new Date(data.tanggal),
      data.jenis,
      data.kategori,
      data.deskripsi,
      parseFloat(cleanJumlah),
      data.status || 'Lunas',
      data.metode,
      data.wallet,
      data.tags || '',
      data.catatan || '',
      Session.getActiveUser().getEmail()
    ];
    
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Auto-sort by date (newest first)
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.sort({column: 1, ascending: false});
    
    // Update dashboard and force flush
    updateDashboard();
    SpreadsheetApp.flush();
    
    return {success: true, message: 'Transaksi berhasil ditambahkan!'};
  } catch (error) {
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== GET INITIAL DATA (FOR FORMS) =====
function getInitialData() {
  return {
    wallets: getWallets(),
    categories: getCategories()
  };
}

// ===== GET CATEGORIES FUNCTION =====
function getCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Kategori');
  
  const pemasukan = sheet.getRange('A4:A' + sheet.getLastRow()).getValues().flat().filter(val => val);
  const pengeluaran = sheet.getRange('D5:D' + sheet.getLastRow()).getValues().flat().filter(val => val);
  
  return {
    pemasukan: pemasukan,
    pengeluaran: pengeluaran
  };
}

// ===== GET WALLETS FUNCTION =====
function getWallets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Wallets');
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 4) return [];
  
  const wallets = sheet.getRange(4, 1, lastRow - 3, 5).getValues();
  return wallets.filter(wallet => wallet[4] === 'Aktif').map(wallet => wallet[0]);
}

// ===== UPDATE DASHBOARD =====
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Dashboard');
  const transaksiSheet = ss.getSheetByName('Transaksi');
  
  if (!transaksiSheet || transaksiSheet.getLastRow() < 2) return;
  
  // Update top categories
  updateTopCategories(dashboardSheet, transaksiSheet);
  
  // Create/update charts
  updateCharts(dashboardSheet, transaksiSheet);
  
  SpreadsheetApp.flush();
}

// ===== UPDATE TOP CATEGORIES =====
function updateTopCategories(dashboardSheet, transaksiSheet) {
  const currentMonth = new Date();
  const startDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth(), 1);
  const endDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 0);
  
  // Get all transactions for current month
  const data = transaksiSheet.getDataRange().getValues();
  const headers = data[0];
  const transactions = data.slice(1);
  
  // Filter pengeluaran for current month
  const pengeluaran = transactions.filter(row => {
    const date = new Date(row[0]);
    return row[1] === 'Pengeluaran' && 
           date >= startDate && 
           date <= endDate;
  });
  
  // Group by category
  const categoryTotals = {};
  pengeluaran.forEach(row => {
    const category = row[2];
    const amount = row[4];
    if (!categoryTotals[category]) {
      categoryTotals[category] = {total: 0, count: 0};
    }
    categoryTotals[category].total += amount;
    categoryTotals[category].count += 1;
  });
  
  // Sort and get top 5
  const sortedCategories = Object.entries(categoryTotals)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 5);
  
  // Update dashboard
  const startRow = 16;
  dashboardSheet.getRange(startRow, 1, 5, 4).clearContent();
  
  sortedCategories.forEach((cat, index) => {
    const row = startRow + index;
    dashboardSheet.getRange(row, 1).setValue(cat[0]);
    dashboardSheet.getRange(row, 2).setValue(cat[1].total);
    dashboardSheet.getRange(row, 3).setValue(cat[1].count);
    dashboardSheet.getRange(row, 4).setFormula('=B' + row + '/$G$6');
  });
  
  dashboardSheet.getRange(startRow, 2, 5, 1).setNumberFormat('"Rp "#,##0');
  dashboardSheet.getRange(startRow, 4, 5, 1).setNumberFormat('0%');
}

// ===== UPDATE CHARTS =====
function updateCharts(dashboardSheet, transaksiSheet) {
  // Remove existing charts
  const charts = dashboardSheet.getCharts();
  charts.forEach(chart => dashboardSheet.removeChart(chart));
  
  // Create pie chart for wallet distribution
  const walletChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dashboardSheet.getRange('A6:B10'))
    .setPosition(15, 6, 0, 0)
    .setOption('title', 'Distribusi Saldo per Wallet')
    .setOption('width', 450)
    .setOption('height', 300)
    .setOption('pieHole', 0.4)
    .build();
    
  dashboardSheet.insertChart(walletChart);

  // Create column chart for monthly income vs expense
  const summaryRange = dashboardSheet.getRange('F5:G6');
  const incomeExpenseChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(summaryRange)
    .setPosition(15, 9, 0, 0) // Repositioned
    .setOption('title', 'Pemasukan vs Pengeluaran Bulan Ini')
    .setOption('width', 450)
    .setOption('height', 300)
    .build();
  
  dashboardSheet.insertChart(incomeExpenseChart);

  // Create Donut chart for 50/30/20
  const analysisRange = dashboardSheet.getRange('E29:E31');
  const analysisValuesRange = dashboardSheet.getRange('H29:H31');
  const donutChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(analysisRange)
    .addRange(analysisValuesRange)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setPosition(28, 9, 0, 0) // Repositioned
    .setOption('title', 'Alokasi Dana (50/30/20)')
    .setOption('width', 450)
    .setOption('height', 300)
    .setOption('pieHole', 0.4)
    .setOption('colors', ['#4CAF50', '#FFC107', '#2196F3'])
    .build();

  dashboardSheet.insertChart(donutChart);
}

// ===== GENERATE REPORTS =====
function generateWeeklyReport() {
  generateReport('Mingguan', 7);
}

function generateMonthlyReport() {
  generateReport('Bulanan', 30);
}

function generateBimonthlyReport() {
  generateReport('2 Bulanan', 60);
}

function generateSemesterReport() {
  generateReport('Semester', 180);
}

function generateYearlyReport() {
  generateReport('Tahunan', 365);
}

function generateReport(period, days) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const laporanSheet = ss.getSheetByName('Laporan');
  const transaksiSheet = ss.getSheetByName('Transaksi');
  
  laporanSheet.clear();
  
  // Header
  laporanSheet.getRange('A1').setValue('📊 LAPORAN ' + period.toUpperCase());
  laporanSheet.getRange('A1').setFontSize(18).setFontWeight('bold');
  laporanSheet.getRange('A1:H1').merge();
  
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  
  laporanSheet.getRange('A3').setValue('Periode: ' + startDate.toLocaleDateString('id-ID') + ' - ' + endDate.toLocaleDateString('id-ID'));
  
  // Get data
  const data = transaksiSheet.getDataRange().getValues();
  const headers = data[0];
  const transactions = data.slice(1).filter(row => {
    const date = new Date(row[0]);
    return date >= startDate && date <= endDate;
  });
  
  // Calculate summary
  const summary = {
    totalPemasukan: 0,
    totalPengeluaran: 0,
    totalTransfer: 0,
    countPemasukan: 0,
    countPengeluaran: 0,
    walletSummary: {},
    categorySummary: {}
  };
  
  transactions.forEach(row => {
    const jenis = row[1];
    const kategori = row[2];
    const jumlah = row[4];
    const wallet = row[7];
    
    if (jenis === 'Pemasukan') {
      summary.totalPemasukan += jumlah;
      summary.countPemasukan++;
    } else if (jenis === 'Pengeluaran') {
      summary.totalPengeluaran += jumlah;
      summary.countPengeluaran++;
    }
    
    // Wallet summary
    if (!summary.walletSummary[wallet]) {
      summary.walletSummary[wallet] = {in: 0, out: 0};
    }
    if (jenis === 'Pemasukan' || jenis === 'Transfer Masuk') {
      summary.walletSummary[wallet].in += jumlah;
    } else {
      summary.walletSummary[wallet].out += jumlah;
    }
    
    // Category summary
    if (!summary.categorySummary[kategori]) {
      summary.categorySummary[kategori] = 0;
    }
    summary.categorySummary[kategori] += jumlah;
  });
  
  // Display summary
  laporanSheet.getRange('A5').setValue('RINGKASAN:');
  laporanSheet.getRange('A5').setFontWeight('bold').setBackground('#E3F2FD');
  
  const summaryData = [
    ['Total Pemasukan:', summary.totalPemasukan, 'Jumlah Transaksi:', summary.countPemasukan],
    ['Total Pengeluaran:', summary.totalPengeluaran, 'Jumlah Transaksi:', summary.countPengeluaran],
    ['Selisih:', summary.totalPemasukan - summary.totalPengeluaran, '', ''],
    ['Rata-rata Pemasukan/hari:', summary.totalPemasukan / days, '', ''],
    ['Rata-rata Pengeluaran/hari:', summary.totalPengeluaran / days, '', '']
  ];
  
  laporanSheet.getRange(6, 1, summaryData.length, 4).setValues(summaryData);
  laporanSheet.getRange('B6:B10').setNumberFormat('"Rp "#,##0');
  
  // Wallet summary
  laporanSheet.getRange('A12').setValue('RINGKASAN PER WALLET:');
  laporanSheet.getRange('A12').setFontWeight('bold').setBackground('#E8F5E9');
  
  const walletHeaders = ['Wallet', 'Masuk', 'Keluar', 'Selisih'];
  laporanSheet.getRange(13, 1, 1, walletHeaders.length).setValues([walletHeaders]);
  laporanSheet.getRange(13, 1, 1, walletHeaders.length).setFontWeight('bold');
  
  let walletRow = 14;
  Object.entries(summary.walletSummary).forEach(([wallet, data]) => {
    laporanSheet.getRange(walletRow, 1).setValue(wallet);
    laporanSheet.getRange(walletRow, 2).setValue(data.in);
    laporanSheet.getRange(walletRow, 3).setValue(data.out);
    laporanSheet.getRange(walletRow, 4).setValue(data.in - data.out);
    walletRow++;
  });
  
  laporanSheet.getRange(14, 2, walletRow - 14, 3).setNumberFormat('"Rp "#,##0');
  
  // Top categories
  const categoryRow = walletRow + 2;
  laporanSheet.getRange(categoryRow, 1).setValue('TOP KATEGORI:');
  laporanSheet.getRange(categoryRow, 1).setFontWeight('bold').setBackground('#FFF3E0');
  
  const sortedCategories = Object.entries(summary.categorySummary)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  laporanSheet.getRange(categoryRow + 1, 1).setValue('Kategori');
  laporanSheet.getRange(categoryRow + 1, 2).setValue('Total');
  laporanSheet.getRange(categoryRow + 1, 1, 1, 2).setFontWeight('bold');
  
  sortedCategories.forEach((cat, index) => {
    laporanSheet.getRange(categoryRow + 2 + index, 1).setValue(cat[0]);
    laporanSheet.getRange(categoryRow + 2 + index, 2).setValue(cat[1]);
  });
  
  laporanSheet.getRange(categoryRow + 2, 2, sortedCategories.length, 1).setNumberFormat('"Rp "#,##0');
  
  SpreadsheetApp.getUi().alert('✅ Laporan ' + period + ' berhasil dibuat!');
}

// ===== TRANSFER DIALOG =====
function showTransferDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TransferForm')
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '💳 Transfer Antar Wallet');
}

// ===== PROCESS TRANSFER =====
function processTransfer(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Transaksi');
    const cleanJumlah = parseFloat(data.jumlah.replace(/\./g, ''));
    const cleanBiayaAdmin = data.biayaAdmin ? parseFloat(data.biayaAdmin.replace(/\./g, '')) : 0;
    const user = Session.getActiveUser().getEmail();
    const tanggal = new Date(data.tanggal);

    if (data.walletAsal === data.walletTujuan) {
      throw new Error('Wallet asal dan tujuan tidak boleh sama.');
    }
    if (isNaN(cleanJumlah) || cleanJumlah <= 0) {
      throw new Error('Jumlah transfer tidak valid.');
    }

    const transactions = [];

    // 1. Transfer Keluar
    const transferOut = [
      tanggal,
      'Transfer Keluar',
      'Transfer',
      'Transfer ke ' + data.walletTujuan,
      cleanJumlah,
      'Lunas',
      'Transfer',
      data.walletAsal,
      'transfer',
      data.catatan || '',
      user
    ];
    transactions.push(transferOut);

    // 2. Transfer Masuk
    const transferIn = [
      tanggal,
      'Transfer Masuk',
      'Transfer',
      'Transfer dari ' + data.walletAsal,
      cleanJumlah,
      'Lunas',
      'Transfer',
      data.walletTujuan,
      'transfer',
      data.catatan || '',
      user
    ];
    transactions.push(transferIn);

    // 3. Biaya Admin (jika ada)
    if (cleanBiayaAdmin > 0) {
      const adminFee = [
        tanggal,
        'Pengeluaran',
        'Biaya Admin',
        'Biaya admin transfer ke ' + data.walletTujuan,
        cleanBiayaAdmin,
        'Lunas',
        'Transfer',
        data.walletAsal, // Biaya admin diambil dari wallet asal
        'biaya_admin',
        'Otomatis dari transfer',
        user
      ];
      transactions.push(adminFee);
    }

    // Write all transactions to the sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, transactions.length, transactions[0].length).setValues(transactions);

    // Sort the entire sheet by date
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.sort({column: 1, ascending: false});

    // Update dashboard and force flush
    updateDashboard();
    SpreadsheetApp.flush();

    return {success: true, message: 'Transfer berhasil diproses!'};
  } catch (error) {
    Logger.log(error);
    return {success: false, message: 'Error: ' + error.toString()};
  }
}

// ===== SHOW SETTINGS =====
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsForm')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ Pengaturan');
}

// ===== RESET DATA =====
function resetAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Peringatan!',
    'Ini akan menghapus SEMUA data transaksi. Apakah Anda yakin?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transaksi');
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    updateDashboard();
    ui.alert('✅ Semua data transaksi telah dihapus.');
  }
}

// ===== SHOW HELP =====
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = '💰 PANDUAN PENGGUNAAN SISTEM KEUANGAN PRIBADI\n\n' +
    '📋 FITUR UTAMA:\n' +
    '• Multi-Wallet: Kelola saldo dari berbagai bank & e-wallet\n' +
    '• Kategori: Organisir transaksi berdasarkan kategori\n' +
    '• Laporan: Generate laporan mingguan hingga tahunan\n' +
    '• Dashboard: Lihat ringkasan keuangan real-time\n' +
    '• Transfer: Catat perpindahan uang antar wallet\n\n' +
    '🚀 CARA MEMULAI:\n' +
    '1. Klik "Setup Awal" untuk membuat template\n' +
    '2. Tambah transaksi melalui menu "Transaksi Baru"\n' +
    '3. Lihat ringkasan di Dashboard (otomatis update)\n' +
    '4. Generate laporan sesuai kebutuhan\n\n' +
    '💡 TIPS:\n' +
    '• Gunakan kategori untuk tracking pengeluaran\n' +
    '• Set reminder untuk input transaksi harian\n' +
    '• Review laporan bulanan untuk evaluasi\n' +
    '• Backup spreadsheet secara berkala\n\n' +
    '📧 Butuh bantuan? Hubungi support';
  
  ui.alert('❓ Bantuan', helpText, ui.ButtonSet.OK);
}