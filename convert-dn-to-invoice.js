const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { loadHarga, detectEntity } = require('./load-harga');

// Path ke file template Invoice Accurate
const TEMPLATE_PATH = path.join(__dirname, 'template', 'sales-invoice-import-file-id.xlsx');

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
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.xlsx' || ext === '.xls') {
    return parseDNExcel(filePath);
  } else if (ext === '.pdf') {
    console.log('Error: PDF parsing belum didukung langsung.');
    console.log('Silakan export DN dalam format Excel (.xlsx) dari Accurate.');
    process.exit(1);
  }
  console.log('Error: Format file tidak didukung:', ext);
  process.exit(1);
}

function parseDNExcel(filePath) {
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

    if (row[25] && String(row[25]).startsWith('DN/')) {
      dnNumber = String(row[25]).trim();
    }

    if (row[12] && typeof row[12] === 'string' && /^\d{1,2}\s\w+\s\d{4}$/.test(String(row[12]).trim())) {
      dnDate = String(row[12]).trim();
    }
    if (row[12] instanceof Date) {
      dnDate = formatDate(row[12]);
    }

    if (row[1] && /^(CV|PT)\s/i.test(String(row[1]).trim()) && !inItems) {
      customerName = String(row[1]).trim();
    }

    if (row[12] && String(row[12]).includes('Warehouse')) {
      warehouse = String(row[12]).trim();
    }

    if (row[1] === 'Item Kode' && row[22] === 'Qty') {
      inItems = true;
      continue;
    }

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
  console.log('  ZUMA - Konversi DN (DDD) ke Invoice Import');
  console.log('  Faktur Penjualan untuk Accurate Online DDD');
  console.log('============================================');
  console.log('');
  console.log('Penggunaan:');
  console.log('  node convert-dn-to-invoice.js <file_dn>');
  console.log('  Contoh: node convert-dn-to-invoice.js dn.xlsx');
  console.log('');

  const dnFile = (process.argv[2] || '').replace(/"/g, '').trim();

  if (!dnFile) {
    console.log('Error: Path file DN tidak boleh kosong.');
    console.log('Contoh: node convert-dn-to-invoice.js "C:\\path\\dn.xlsx"');
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
  console.log('');

  // Detect entity from customer name and load harga
  const entity = detectEntity(dnData.customerName);
  let hargaMap = new Map();
  if (entity) {
    console.log(`Entity terdeteksi: ${entity} (dari customer: ${dnData.customerName})`);
    hargaMap = loadHarga(entity);
  } else {
    console.log('Warning: Entity tidak terdeteksi dari customer name. Harga tidak diisi.');
  }

  // --- Build workbook using ExcelJS ---
  const wb = new ExcelJS.Workbook();

  let hasTemplate = false;
  if (fs.existsSync(TEMPLATE_PATH)) {
    await wb.xlsx.readFile(TEMPLATE_PATH);
    hasTemplate = true;
    console.log('Template Invoice Accurate ditemukan, menggunakan format template asli.');
  } else {
    console.log('Template tidak ditemukan, membuat file baru.');
  }

  let ws;
  if (hasTemplate) {
    ws = wb.getWorksheet('Template Faktur Penjualan') || wb.getWorksheet(1);
    const rowCount = ws.rowCount;
    for (let i = rowCount; i >= 1; i--) {
      ws.spliceRows(i, 1);
    }
  } else {
    ws = wb.addWorksheet('Template Faktur Penjualan');
  }

  const COL_COUNT = 52;

  // --- Row 1: HEADER labels ---
  const headerLabels = [
    "HEADER", "No Faktur", "Tgl Faktur", "No Pelanggan", "Alamat Faktur", "Kena PPN",
    "Total Termasuk PPN", "Nomor Faktur Pajak", "Faktur Dimuka", "Diskon Faktur (%)",
    "Diskon Faktur (Rp)", "Keterangan", "Nama Cabang", "No PO", "Pengiriman",
    "Tgl Pengiriman", "FOB", "Syarat Pembayaran", "Bank Pembayaran", "Nilai Pembayaran",
    "Kustom Karakter 1", "Kustom Karakter 2", "Kustom Karakter 3",
    "Kustom Karakter 4", "Kustom Karakter 5", "Kustom Karakter 6",
    "Kustom Karakter 7", "Kustom Karakter 8", "Kustom Karakter 9", "Kustom Karakter 10",
    "Kustom Angka 1", "Kustom Angka 2", "Kustom Angka 3", "Kustom Angka 4",
    "Kustom Angka 5", "Kustom Angka 6", "Kustom Angka 7", "Kustom Angka 8",
    "Kustom Angka 9", "Kustom Angka 10",
    "Kustom Tanggal 1", "Kustom Tanggal 2",
    "Nomor VA", "Nomor Akun Piutang", "Pembayaran Dengan Kode Unik", "Sub Company Code",
    "", "", "", "", "", ""
  ];
  const r1 = ws.addRow(headerLabels);
  applyRowStyle(r1, STYLES.HEADER, COL_COUNT);

  // --- Row 2: ITEM labels ---
  const itemLabels = [
    "ITEM", "Kode Barang", "Nama Barang", "Kuantitas", "Satuan", "Harga Satuan",
    "Diskon Barang (%)", "Diskon Barang (Rp)", "Catatan Barang", "Nama Gudang",
    "ID Salesman", "Nama Dept Barang", "No Proyek Barang",
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
    "Kategori Keuangan 10",
    "No. Pengiriman Pesanan", "No. Pesanan Penjualan"
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
    "No. Pesanan Penjualan",
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
  ];
  const r3 = ws.addRow(expenseLabels);
  applyRowStyle(r3, STYLES.EXPENSE, COL_COUNT);

  // --- Row 4: HEADER data (only column A gets color) ---
  const tglFaktur = convertDateToAccurate(dnData.dnDate);
  const keterangan = `Invoice dari ${dnData.dnNumber}`;

  const headerData = new Array(COL_COUNT).fill('');
  headerData[0] = 'HEADER';
  headerData[1] = '';                // No Faktur (auto)
  headerData[2] = tglFaktur;        // Tgl Faktur
  headerData[3] = '';                // No Pelanggan (diisi manual di Accurate)
  headerData[4] = '';                // Alamat Faktur
  headerData[5] = 'Ya';             // Kena PPN
  headerData[6] = 'Ya';             // Total Termasuk PPN
  headerData[7] = '';                // Nomor Faktur Pajak
  headerData[8] = 'Tidak';          // Faktur Dimuka
  headerData[9] = '';                // Diskon Faktur (%)
  headerData[10] = '';               // Diskon Faktur (Rp)
  headerData[11] = keterangan;       // Keterangan
  headerData[12] = '';               // Nama Cabang
  headerData[13] = '';               // No PO
  headerData[14] = '';               // Pengiriman
  headerData[15] = tglFaktur;       // Tgl Pengiriman = sama dengan Tgl Faktur
  const r4 = ws.addRow(headerData);
  applyCellStyle(r4, 1, STYLES.HEADER);

  // --- Item rows (only column A gets color) ---
  let missingHarga = 0;
  for (const item of dnData.items) {
    const harga = hargaMap.get(item.kode) || 0;
    if (!harga && hargaMap.size > 0) missingHarga++;

    const itemData = new Array(COL_COUNT).fill('');
    itemData[0] = 'ITEM';
    itemData[1] = item.kode;         // Kode Barang
    itemData[2] = item.nama;         // Nama Barang
    itemData[3] = item.qty;          // Kuantitas
    itemData[4] = item.unit;         // Satuan
    itemData[5] = harga || '';       // Harga Satuan (after diskon dari Master Harga)
    itemData[6] = '';                // Diskon Barang (%)
    itemData[7] = '';                // Diskon Barang (Rp)
    itemData[8] = '';                // Catatan Barang
    itemData[9] = dnData.warehouse;  // Nama Gudang (dari DN)
    itemData[10] = '';               // ID Salesman
    itemData[11] = '';               // Nama Dept Barang
    itemData[12] = '';               // No Proyek Barang
    const ri = ws.addRow(itemData);
    applyCellStyle(ri, 1, STYLES.ITEM);
  }

  // --- Output ---
  const dnBaseName = dnData.dnNumber.replace(/\//g, '-') || 'DN';
  const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const timeDetail = new Date().toTimeString().slice(0, 8).replace(/:/g, '');
  const outputFileName = `INV-DDD-dari-${dnBaseName}-${timestamp}-${timeDetail}.xlsx`;
  const outputDir = path.dirname(dnFile);
  const outputPath = path.join(outputDir, outputFileName);

  await wb.xlsx.writeFile(outputPath);

  console.log('');
  console.log('File Invoice berhasil dibuat!');
  console.log(`   File: ${outputPath}`);
  console.log(`   Untuk: DDD (PT. Dream Dare Discover)`);
  console.log(`   Pelanggan: ${dnData.customerName} (${entity || '?'}) - No Pelanggan diisi manual`);
  console.log(`   Gudang: ${dnData.warehouse}`);
  console.log(`   Tgl Pengiriman: ${tglFaktur} (sama dengan Tgl Faktur)`);
  console.log(`   Harga: Master Harga ${entity || '-'} (after diskon)`);
  if (missingHarga > 0) {
    console.log(`   WARNING: ${missingHarga} SKU tidak ditemukan di Master Harga!`);
  }
  console.log(`   Keterangan: ${keterangan}`);
  console.log(`   Items: ${dnData.items.length} SKU`);
  console.log('');
  console.log('   Silakan import file ini di Accurate Online DDD.');
}

main().catch(err => {
  console.error('Error:', err.message);
});
