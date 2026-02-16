# DN to PO & Invoice Converter - ZUMA

## Deskripsi
Script Node.js untuk mengkonversi file **Delivery Note (DN)** dari entitas **DDD (PT. Dream Dare Discover)** menjadi **2 file output** yang siap diimpor ke **Accurate Online**:

1. **PO (Pesanan Pembelian)** - untuk import ke Accurate **MBB** atau **UBB** (sisi pembeli)
2. **Invoice (Faktur Penjualan)** - untuk import ke Accurate **DDD** (sisi penjual)

> **PENTING:** Setiap DN yang diproses **WAJIB menghasilkan 2 output** (PO + Invoice).

> **FORMAT FILE:** Mendukung **PDF** dan **Excel (.xlsx)** — script auto-detect file extension.

## Latar Belakang
ZUMA adalah perusahaan produksi sandal yang memiliki 4 entitas:
- **LJBB** - CV Lancar Jaya Besar Bersama
- **DDD** - PT. Dream Dare Discover
- **MBB** - CV Makmur Besar Bersama
- **UBB** - CV Untung Besar Bersama

### Alur Transaksi Internal
```
LJBB ──(jual)──> DDD ──(jual)──> MBB / UBB
```
- DDD mengeluarkan **DN (Delivery Note)** saat menjual ke MBB/UBB
- Dari 1 DN, dihasilkan:
  - **Invoice** untuk DDD (mencatat penjualan)
  - **PO** untuk MBB/UBB (mencatat pembelian)

## Instalasi

```bash
cd dn-to-po
npm install
```

## Cara Penggunaan

### Setiap DN harus generate 2 file:

**1. Generate Invoice (untuk DDD)**
```bash
node convert-dn-to-invoice.js <file_DN>
```

**2. Generate PO (untuk MBB/UBB)**
```bash
node convert-dn-to-po.js <file_DN> <entitas>
```

### Parameter

| Script | Parameter | Keterangan |
|--------|-----------|-----------|
| `convert-dn-to-invoice.js` | `file_DN` | Path file DN (**PDF** atau **.xlsx**) dari Accurate DDD |
| `convert-dn-to-po.js` | `file_DN` | Path file DN (**PDF** atau **.xlsx**) dari Accurate DDD |
| `convert-dn-to-po.js` | `entitas` | `MBB` atau `UBB` |

### Contoh Lengkap (1 DN = 2 output)

**Dari Excel:**
```bash
# 1. Invoice untuk DDD
node convert-dn-to-invoice.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx"

# 2. PO untuk MBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx" MBB
```

**Dari PDF:**
```bash
# 1. Invoice untuk DDD
node convert-dn-to-invoice.js "C:\Users\ZUMA\Downloads\dn_pluit.pdf"

# 2. PO untuk MBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\dn_pluit.pdf" MBB
```

**Format file auto-detected** — script membaca extension `.pdf` atau `.xlsx` dan menggunakan parser yang sesuai.

### Output Files
```
INV-DDD-dari-{NO_DN}-{TANGGAL}-{JAM}.xlsx    --> import ke Accurate DDD
PO-{ENTITAS}-dari-{NO_DN}.xlsx                --> import ke Accurate MBB/UBB
```

## Harga

Harga satuan otomatis diambil dari file **Master Harga** (`template/Master Harga.xlsx`):
- Sheet **MBB** untuk transaksi ke MBB
- Sheet **UBB** untuk transaksi ke UBB
- Menggunakan kolom **Harga After Diskon** (harga setelah diskon)
- Entity otomatis terdeteksi dari nama customer di DN
- Jika SKU tidak ditemukan di Master Harga, akan muncul warning

## Mapping Data DN ke Output

### Invoice (Faktur Penjualan DDD)

#### Header
| Field | Sumber | Keterangan |
|-------|--------|------------|
| No Faktur | - | Dikosongkan, auto-generate oleh Accurate |
| Tgl Faktur | Tanggal DN | Format DD/MM/YYYY |
| No Pelanggan | - | Dikosongkan, diisi manual di Accurate |
| Kena PPN | Ya | Transaksi internal kena PPN |
| Total Termasuk PPN | Ya | |
| Keterangan | No DN | `Invoice dari DN/DDD/WHB/2026/II/021` |
| Tgl Pengiriman | Tanggal DN | Sama dengan Tgl Faktur |

#### Item
| Field | Sumber | Keterangan |
|-------|--------|------------|
| Kode Barang | Item Kode DN | Kode SKU produk |
| Nama Barang | Name Article DN | Nama lengkap artikel |
| Kuantitas | Qty DN | Jumlah pasang |
| Satuan | Unit DN | Default: PAIR |
| Harga Satuan | Master Harga | Harga after diskon sesuai entity |
| Nama Gudang | Warehouse DN | Nama gudang dari DN |

### PO (Pesanan Pembelian MBB/UBB)

#### Header
| Field | Sumber | Keterangan |
|-------|--------|------------|
| No Form | - | Dikosongkan, auto-generate oleh Accurate |
| Tgl Pesanan | Tanggal DN | Format DD/MM/YYYY |
| No Pemasok | - | Dikosongkan, diisi manual di Accurate |
| Kena PPN | Ya | Transaksi internal kena PPN |
| Total Termasuk PPN | Ya | |
| Keterangan | No DN | `PO dari DN/DDD/WHB/2026/II/021` |
| Mata Uang | IDR | Default Rupiah |

#### Item
| Field | Sumber | Keterangan |
|-------|--------|------------|
| Kode Barang | Item Kode DN | Kode SKU produk |
| Nama Barang | Name Article DN | Nama lengkap artikel |
| Kuantitas | Qty DN | Jumlah pasang |
| Satuan | Unit DN | Default: PAIR |
| Harga Satuan | Master Harga | Harga after diskon sesuai entity |

## Format File Output

Kedua file output sama persis dengan template Accurate asli:
- 4 Sheet: Template data + 3 sheet Penjelasan Kolom
- Warna: HEADER (hijau), ITEM (biru), EXPENSE (orange)
- Row 1-3: Label kolom dengan warna full row
- Row 4+: Data (hanya kolom A berwarna)

## Dependencies
- [xlsx](https://www.npmjs.com/package/xlsx) - Library untuk baca file Excel (DN & Master Harga)
- [exceljs](https://www.npmjs.com/package/exceljs) - Library untuk tulis file Excel dengan styling/warna
- [pdf-parse](https://www.npmjs.com/package/pdf-parse) - Library untuk ekstrak text dari PDF

## File Template
Disimpan di folder `template/`:
- `purchase-order-import-file.xlsx` - Template PO Accurate
- `sales-invoice-import-file-id.xlsx` - Template Invoice Accurate
- `Master Harga.xlsx` - Data master harga (sheet UBB & MBB)

## Warehouse Coverage
Script mendukung **semua warehouse DDD** (nama gudang diambil langsung dari DN):
- **Warehouse Bali Gatsu - Box**
- **Warehouse Pluit**
- Dan warehouse lainnya sesuai nama di Accurate

Nama gudang di output Invoice/PO akan **exact match** dengan nama di DN (tidak dimodifikasi).

Harga diambil dari **Master Harga** sheet MBB/UBB (coverage 5000+ SKU), tidak bergantung pada warehouse tertentu.

## Catatan
- **Format file:** Mendukung PDF dan Excel (.xlsx) — auto-detect extension
- **PDF parsing:** Script extract DN number, date, warehouse, dan items dari text PDF
- Setiap DN **WAJIB menghasilkan 2 output**: Invoice (DDD) + PO (MBB/UBB)
- Harga satuan otomatis dari **Master Harga** (kolom Harga After Diskon)
- No Pelanggan dan No Pemasok **tidak** diisi otomatis, diisi manual di Accurate
- Nomor DN tercatat di kolom **Keterangan** sebagai referensi
- Tgl Pengiriman di Invoice = sama dengan Tgl Faktur
- **Nama Gudang di Invoice = exact match dari DN** (tidak dimodifikasi)
- **Tested:** DN/DDD/WHB/2026/II/021 (Warehouse Bali Gatsu - Box), DN/DDD/WHJ/2026/II/007 (Warehouse Pluit)
