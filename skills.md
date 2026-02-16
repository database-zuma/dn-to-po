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
node convert-dn-to-po.js <file_DN> <entitas> <no_pemasok>
```

### Parameter
| Parameter | Keterangan | Contoh |
|-----------|-----------|--------|
| `file_DN` | Path ke file DN (.xlsx) yang diexport dari Accurate DDD | `"C:\Downloads\dn.xlsx"` |
| `entitas` | Entitas tujuan PO: `MBB` atau `UBB` | `MBB` |
| `no_pemasok` | No Pemasok / Supplier ID DDD di Accurate entitas tujuan | `V.00001` |

### Contoh

```bash
# Generate PO untuk MBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx" MBB V.00001

# Generate PO untuk UBB
node convert-dn-to-po.js "C:\Users\ZUMA\Downloads\pengiriman_pesanan.xlsx" UBB V.00002
```

### Output
File PO akan disimpan di folder yang sama dengan file DN input, dengan nama:
```
PO-{ENTITAS}-dari-{NO_DN}-{TANGGAL}.xlsx
```
Contoh: `PO-MBB-dari-DN-DDD-WHB-2026-II-021-20260216.xlsx`

## Mapping Data DN ke PO

### Header PO
| Field PO | Sumber | Keterangan |
|----------|--------|------------|
| No Form | - | Dikosongkan, auto-generate oleh Accurate |
| Tgl Pesanan | Tanggal DN | Format DD/MM/YYYY |
| No Pemasok | Input user | ID supplier DDD di Accurate MBB/UBB |
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
File Excel sesuai template import "Pesanan Pembelian" Accurate Online:
- Row 0: Label kolom HEADER
- Row 1: Label kolom ITEM
- Row 2: Label kolom EXPENSE
- Row 3+: Data (HEADER diikuti ITEM rows)

## Dependencies
- [xlsx](https://www.npmjs.com/package/xlsx) - Library untuk baca/tulis file Excel

## Catatan
- Harga satuan **tidak** diisi otomatis dari DN (karena DN tidak mengandung harga). Harga perlu diisi manual di Accurate setelah import atau menggunakan daftar harga yang sudah ada.
- No Pemasok DDD **berbeda** di setiap entitas (MBB dan UBB), pastikan menggunakan kode yang benar.
- Nomor DN DDD akan tercatat di kolom **Keterangan** PO sebagai referensi.
