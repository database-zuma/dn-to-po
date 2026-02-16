# DN to PO Converter - ZUMA

## Deskripsi
Script Node.js untuk mengkonversi file **Delivery Note (DN)** dari entitas **DDD (PT. Dream Dare Discover)** menjadi file **Purchase Order (PO) Import** yang siap diimpor ke **Accurate Online** untuk entitas **MBB (CV Makmur Besar Bersama)** atau **UBB (CV Untung Besar Bersama)**.

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
- DDD mengeluarkan **DN (Delivery Note)** dan **Invoice** saat menjual ke MBB/UBB
- MBB/UBB perlu membuat **PO (Purchase Order)** di Accurate untuk mencatat pembelian dari DDD
- Script ini otomatis mengkonversi DN dari DDD menjadi file PO import untuk MBB/UBB

## Instalasi

```bash
cd dn-to-po
npm install
```

## Cara Penggunaan

```bash
node convert-dn-to-po.js <file_DN> <entitas>
```

### Parameter
| Parameter | Keterangan | Contoh |
|-----------|-----------|--------|
| `file_DN` | Path ke file DN (.xlsx) yang diexport dari Accurate DDD | `"C:\Downloads\dn.xlsx"` |
| `entitas` | Entitas tujuan PO: `MBB` atau `UBB` | `MBB` |

### Contoh

```bash
# Generate PO untuk MBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx" MBB

# Generate PO untuk UBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx" UBB
```

### Output
File PO akan disimpan di folder yang sama dengan file DN input, dengan nama:
```
PO-{ENTITAS}-dari-{NO_DN}.xlsx
```
Contoh: `PO-MBB-dari-DN-DDD-WHB-2026-II-021.xlsx`

## Mapping Data DN ke PO

### Header PO
| Field PO | Sumber | Keterangan |
|----------|--------|------------|
| No Form | - | Dikosongkan, auto-generate oleh Accurate |
| Tgl Pesanan | Tanggal DN | Format DD/MM/YYYY |
| No Pemasok | - | Dikosongkan, diisi manual di Accurate |
| Kena PPN | Ya | Transaksi internal kena PPN |
| Total Termasuk PPN | Ya | |
| Keterangan | No DN | `PO dari DN/DDD/WHB/2026/II/021` |
| Mata Uang | IDR | Default Rupiah |

### Item PO
| Field PO | Sumber DN | Keterangan |
|----------|-----------|------------|
| Kode Barang | Item Kode | Kode SKU produk |
| Nama Barang | Name Article | Nama lengkap artikel |
| Kuantitas | Qty | Jumlah pasang |
| Satuan | Unit | Default: PAIR |
| Harga Satuan | - | Dikosongkan, diisi manual di Accurate |

## Format File

### Input: DN (Delivery Note)
File export "Pengiriman Pesanan" dari Accurate Online DDD dengan struktur:
- Header: nama perusahaan, alamat, tanggal, nomor DN, nama customer, warehouse
- Detail: Item Kode, Name Article, Qty, Unit
- Mendukung multi-halaman (page break)

### Output: PO Import
File Excel sesuai template import "Pesanan Pembelian" Accurate Online (sama persis dengan template asli):
- 4 Sheet: Template Pesanan Pembelian + 3 sheet Penjelasan Kolom
- Warna: HEADER (hijau), ITEM (biru), EXPENSE (orange) - sesuai template Accurate
- Row 1-3: Label kolom (HEADER / ITEM / EXPENSE) dengan warna full row
- Row 4+: Data (hanya kolom A berwarna)

## Dependencies
- [xlsx](https://www.npmjs.com/package/xlsx) - Library untuk baca file Excel (DN)
- [exceljs](https://www.npmjs.com/package/exceljs) - Library untuk tulis file Excel dengan styling/warna

## Workflow Delivery

### 1. Generate PO
```bash
node convert-dn-to-po.js <file_DN> <entitas>
```

### 2. Upload ke Google Drive
```bash
gog drive upload <output_file> --name "PO-{ENTITAS}-dari-{NO_DN}.xlsx" --json
gog drive share <file_id> --email wayan@zuma.id --role writer
gog drive share <file_id> --email database@zuma.id --role writer
gog drive share <file_id> --anyone --role reader
```

### 3. Kirim ke User
**Format standar:**
```
📄 **PO-{ENTITAS}-dari-{NO_DN}**

{NO_DN}
{X} SKU, {Y} pairs
Tanggal: {TANGGAL}

🔗 **Google Sheets:**
{GSHEET_LINK}
```

Kirim bersamaan:
- File Excel (attachment)
- Google Sheets link (di caption)

## Catatan
- Harga satuan **tidak** diisi otomatis dari DN (karena DN tidak mengandung harga). Harga perlu diisi manual di Accurate setelah import.
- No Pemasok (Supplier ID DDD) **tidak** diisi otomatis. Diisi manual di Accurate saat import.
- Nomor DN DDD akan tercatat di kolom **Keterangan** PO sebagai referensi.
- File template Accurate asli disimpan di folder `template/` untuk menjaga format yang konsisten.
- **Filename:** Tanpa timestamp — cukup nomor DN aja (clean & predictable)
