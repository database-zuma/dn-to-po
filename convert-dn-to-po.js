const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Path ke file template PO Accurate (disimpan di folder script)
const TEMPLATE_PATH = path.join(__dirname, 'template', 'purchase-order-import-file.xlsx');

// Warna sesuai template Accurate
const STYLES = {
  HEADER: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } },
    font: { color: { argb: 'FFFFFFFF' }, bold: true }
  },
  ITEM: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } },
    font: { color: { argb: 'FFFFFFFF' }, bold: true }
  },
  EXPENSE: {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF6600' } },
    font: { color: { argb: 'FFFFFFFF' }, bold: true }
  }
};

function parseDN(filePath) {
  const wb = XLSX.readFile(filePath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  let dnNumber = '';
  let dnDate = '';
  let customerName = '';
  let warehouse = '';
  const items = [];

  let inItems = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // Detect DN Number (column 25, index 25)
    if (row[25] && String(row[25]).startsWith('DN/')) {
      dnNumber = String(row[25]).trim();
    }

    // Detect Date (column 12) - format "13 Feb 2026"
    if (row[12] && typeof row[12] === 'string' && /^\d{1,2}\s\w+\s\d{4}$/.test(String(row[12]).trim())) {
      dnDate = String(row[12]).trim();
    }
    if (row[12] instanceof Date) {
      dnDate = formatDate(row[12]);
    }

    // Detect Customer Name (col 1) - look for CV/PT patterns
    if (row[1] && /^(CV|PT)\s/i.test(String(row[1]).trim()) && !inItems) {
      customerName = String(row[1]).trim();
    }

    // Detect Warehouse (column 12, contains "Warehouse")
    if (row[12] && String(row[12]).includes('Warehouse')) {
      warehouse = String(row[12]).trim();
    }

    // Detect item header row
    if (row[1] === 'Item Kode' && row[22] === 'Qty') {
      inItems = true;
      continue;
    }

    // Parse item rows
    if (inItems) {
      const kode = String(row[1] || '').trim();
      const nama = String(row[7] || '').trim();
      const qty = row[22];
      const unit = String(row[31] || '').trim();

      if (!kode && !nama && !qty) {
        inItems = false;
        continue;
      }

      if (kode === 'Item Kode') continue;

      if (kode && qty) {
        items.push({ kode, nama, qty: Number(qty), unit: unit || 'PAIR' });
      }
    }
  }

  return { dnNumber, dnDate, customerName, warehouse, items };
}

function formatDate(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

function convertDateToAccurate(dateStr) {
  const months = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
    'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12',
    'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04', 'Mei': '05',
    'Juni': '06', 'Juli': '07', 'Agustus': '08', 'September': '09', 'Oktober': '10',
    'November': '11', 'Desember': '12'
  };
  const parts = dateStr.trim().split(/\s+/);
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = months[parts[1]] || '01';
    const year = parts[2];
    return `${day}/${month}/${year}`;
  }
  return dateStr;
}

function applyRowStyle(row, style, colCount) {
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.fill = style.fill;
    cell.font = style.font;
  }
}

function applyCellStyle(row, colIndex, style) {
  const cell = row.getCell(colIndex);
  cell.fill = style.fill;
  cell.font = style.font;
}

async function main() {
  console.log('============================================');
  console.log('  ZUMA - Konversi DN (DDD) ke PO Import');
  console.log('  Untuk import ke Accurate Online');
  console.log('============================================');
  console.log('');
  console.log('Penggunaan:');
  console.log('  node convert-dn-to-po.js <file_dn> <entitas>');
  console.log('  Contoh: node convert-dn-to-po.js dn.xlsx MBB');
  console.log('');

  const dnFile = (process.argv[2] || '').replace(/"/g, '').trim();
  const entity = (process.argv[3] || '').toUpperCase().trim();

  if (!dnFile) {
    console.log('Error: Path file DN tidak boleh kosong.');
    console.log('Contoh: node convert-dn-to-po.js "C:\\path\\dn.xlsx" MBB');
    return;
  }
  if (!entity || !['MBB', 'UBB'].includes(entity)) {
    console.log('Error: Entitas harus MBB atau UBB.');
    console.log('Contoh: node convert-dn-to-po.js "C:\\path\\dn.xlsx" MBB');
    return;
  }

  console.log(`Membaca file DN: ${dnFile}`);
  console.log('');

  const dnData = parseDN(dnFile);

  console.log('--- Info DN ---');
  console.log(`No DN      : ${dnData.dnNumber}`);
  console.log(`Tanggal    : ${dnData.dnDate}`);
  console.log(`Customer   : ${dnData.customerName}`);
  console.log(`Warehouse  : ${dnData.warehouse}`);
  console.log(`Total Item : ${dnData.items.length} SKU`);
  console.log(`Total Qty  : ${dnData.items.reduce((s, i) => s + i.qty, 0)} pairs`);
  console.log(`Entitas    : ${entity}`);
  console.log('');

  // --- Build workbook using ExcelJS (with template as base) ---
  const wb = new ExcelJS.Workbook();

  // Try to load template for explanation sheets
  let hasTemplate = false;
  if (fs.existsSync(TEMPLATE_PATH)) {
    await wb.xlsx.readFile(TEMPLATE_PATH);
    hasTemplate = true;
    console.log('Template Accurate ditemukan, menggunakan format template asli.');
  } else {
    console.log('Template tidak ditemukan, membuat file baru.');
  }

  // Get or create the main data sheet
  let ws;
  if (hasTemplate) {
    ws = wb.getWorksheet('Template Pesanan Pembelian') || wb.getWorksheet(1);
    // Clear existing data rows (keep structure)
    // Remove all rows first
    const rowCount = ws.rowCount;
    for (let i = rowCount; i >= 1; i--) {
      ws.spliceRows(i, 1);
    }
  } else {
    ws = wb.addWorksheet('Template Pesanan Pembelian');
  }

  const COL_COUNT = 49;

  // --- Row 1: HEADER labels ---
  const headerLabels = [
    "HEADER", "No Form", "Tgl Pesanan", "No Pemasok", "Alamat Kirim", "Kena PPN",
    "Total Termasuk PPN", "Diskon Pesanan (%)", "Diskon Pesanan (Rp)", "Keterangan",
    "Nama Cabang", "Pengiriman", "Tgl Pengiriman", "FOB", "Syarat Pembayaran",
    "Mata Uang", "Kurs Saldo (Jika Asing)",
    "Kustom Karakter 1", "Kustom Karakter 2", "Kustom Karakter 3",
    "Kustom Karakter 4", "Kustom Karakter 5", "Kustom Karakter 6",
    "Kustom Karakter 7", "Kustom Karakter 8", "Kustom Karakter 9", "Kustom Karakter 10",
    "Kustom Angka 1", "Kustom Angka 2", "Kustom Angka 3", "Kustom Angka 4",
    "Kustom Angka 5", "Kustom Angka 6", "Kustom Angka 7", "Kustom Angka 8",
    "Kustom Angka 9", "Kustom Angka 10",
    "Kustom Tanggal 1", "Kustom Tanggal 2",
    "", "", "", "", "", "", "", "", "", ""
  ];
  const r1 = ws.addRow(headerLabels);
  applyRowStyle(r1, STYLES.HEADER, COL_COUNT);

  // --- Row 2: ITEM labels ---
  const itemLabels = [
    "ITEM", "Kode Barang", "Nama Barang", "Kuantitas", "Satuan", "Harga Satuan",
    "Diskon Barang (%)", "Diskon Barang (Rp)", "Catatan Barang", "Nama Dept Barang",
    "No Proyek Barang", "Nama Gudang",
    "Kustom Karakter 1", "Kustom Karakter 2", "Kustom Karakter 3",
    "Kustom Karakter 4", "Kustom Karakter 5", "Kustom Karakter 6",
    "Kustom Karakter 7", "Kustom Karakter 8", "Kustom Karakter 9",
    "Kustom Karakter 10", "Kustom Karakter 11", "Kustom Karakter 12",
    "Kustom Karakter 13", "Kustom Karakter 14", "Kustom Karakter 15",
    "Kustom Angka1", "Kustom Angka 2", "Kustom Angka 3", "Kustom Angka 4",
    "Kustom Angka 5", "Kustom Angka 6", "Kustom Angka 7", "Kustom Angka 8",
    "Kustom Angka 9", "Kustom Angka 10",
    "Kustom Tanggal 1", "Kustom Tanggal 2",
    "Kategori Keuangan 1", "Kategori Keuangan 2", "Kategori Keuangan 3",
    "Kategori Keuangan 4", "Kategori Keuangan 5", "Kategori Keuangan 6",
    "Kategori Keuangan 7", "Kategori Keuangan 8", "Kategori Keuangan 9",
    "Kategori Keuangan 10"
  ];
  const r2 = ws.addRow(itemLabels);
  applyRowStyle(r2, STYLES.ITEM, COL_COUNT);

  // --- Row 3: EXPENSE labels ---
  const expenseLabels = [
    "EXPENSE", "No Biaya", "Nama Biaya", "Nilai Biaya", "Catatan Biaya",
    "Nama Dept Biaya", "No Proyek Biaya",
    "Kategori Keuangan 1", "Kategori Keuangan 2", "Kategori Keuangan 3",
    "Kategori Keuangan 4", "Kategori Keuangan 5", "Kategori Keuangan 6",
    "Kategori Keuangan 7", "Kategori Keuangan 8", "Kategori Keuangan 9",
    "Kategori Keuangan 10",
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
  ];
  const r3 = ws.addRow(expenseLabels);
  applyRowStyle(r3, STYLES.EXPENSE, COL_COUNT);

  // --- Row 4: HEADER data (only column A gets color) ---
  const tglPesanan = convertDateToAccurate(dnData.dnDate);
  const keterangan = `PO dari ${dnData.dnNumber}`;

  const headerData = new Array(COL_COUNT).fill('');
  headerData[0] = 'HEADER';
  headerData[1] = '';                // No Form (auto)
  headerData[2] = tglPesanan;
  headerData[3] = '';                // No Pemasok (diisi manual di Accurate)
  headerData[4] = '';                // Alamat Kirim
  headerData[5] = 'Ya';             // Kena PPN
  headerData[6] = 'Ya';             // Total Termasuk PPN
  headerData[9] = keterangan;
  headerData[15] = 'IDR';
  const r4 = ws.addRow(headerData);
  applyCellStyle(r4, 1, STYLES.HEADER);

  // --- Item rows (only column A gets color) ---
  for (const item of dnData.items) {
    const itemData = new Array(COL_COUNT).fill('');
    itemData[0] = 'ITEM';
    itemData[1] = item.kode;
    itemData[2] = item.nama;
    itemData[3] = item.qty;
    itemData[4] = item.unit;
    const ri = ws.addRow(itemData);
    applyCellStyle(ri, 1, STYLES.ITEM);
  }

  // --- Output ---
  const dnBaseName = dnData.dnNumber.replace(/\//g, '-') || 'DN';
  const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const timeDetail = new Date().toTimeString().slice(0,8).replace(/:/g,'');
  const outputFileName = `PO-${entity}-dari-${dnBaseName}-${timestamp}-${timeDetail}.xlsx`;
  const outputDir = path.dirname(dnFile);
  const outputPath = path.join(outputDir, outputFileName);

  await wb.xlsx.writeFile(outputPath);

  console.log('');
  console.log('File PO berhasil dibuat!');
  console.log(`   File: ${outputPath}`);
  console.log(`   Entity: ${entity}`);
  console.log(`   Pemasok: (diisi manual di Accurate)`);
  console.log(`   Keterangan: ${keterangan}`);
  console.log(`   Items: ${dnData.items.length} SKU`);
  console.log('');
  console.log(`   Silakan import file ini di Accurate Online ${entity}.`);
}

main().catch(err => {
  console.error('Error:', err.message);
});
