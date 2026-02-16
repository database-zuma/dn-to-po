const XLSX = require('xlsx');
const path = require('path');

const MASTER_HARGA_PATH = path.join(__dirname, 'template', 'Master Harga.xlsx');

/**
 * Load harga after diskon dari Master Harga
 * @param {string} entity - 'MBB' atau 'UBB'
 * @returns {Map<string, number>} Map kode SKU -> harga after diskon
 */
function loadHarga(entity) {
  const wb = XLSX.readFile(MASTER_HARGA_PATH);
  const sheetName = entity.toUpperCase();

  if (!wb.SheetNames.includes(sheetName)) {
    console.log(`Warning: Sheet '${sheetName}' tidak ditemukan di Master Harga.`);
    console.log('Sheet tersedia:', wb.SheetNames.join(', '));
    return new Map();
  }

  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
  const hargaMap = new Map();

  // Data mulai row 3 (index 3), row 0-2 adalah header
  // Kolom: 0=Kode Variant, 1=Nama, 2=Seri, 3=Harga Price Tag, 4=Diskon%, 5=Diskon Rp, 6=Harga After Diskon
  for (let i = 3; i < data.length; i++) {
    const row = data[i];
    const kode = String(row[0] || '').trim();
    const hargaAfterDiskon = Number(row[6]) || 0;

    if (kode && hargaAfterDiskon > 0) {
      hargaMap.set(kode, hargaAfterDiskon);
    }
  }

  console.log(`Master Harga ${sheetName}: ${hargaMap.size} SKU loaded.`);
  return hargaMap;
}

/**
 * Detect entity dari customer name di DN
 * @param {string} customerName
 * @returns {string|null} 'MBB', 'UBB', atau null
 */
function detectEntity(customerName) {
  const upper = customerName.toUpperCase();
  if (upper.includes('MAKMUR') || upper.includes('MBB')) return 'MBB';
  if (upper.includes('UNTUNG') || upper.includes('UBB')) return 'UBB';
  return null;
}

module.exports = { loadHarga, detectEntity, MASTER_HARGA_PATH };
