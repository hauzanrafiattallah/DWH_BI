# 🔍 Laporan Temuan & Solusi — ETL Data Inovasi & Kewirausahaan FIF

**Dibuat:** 2026-11-11 21:04  
**Pipeline:** `code/etl_fif_inovasi.py`  
**Analyst:** ETL Pipeline (Automated)

---

## 📋 Ringkasan Eksekutif

Pipeline ETL memproses **3 dataset Excel** dari program inovasi & kewirausahaan Fakultas Informatika (FIF). Ditemukan **12 kategori masalah** data yang semuanya telah ditangani secara otomatis. Berikut laporan lengkapnya.

| Dataset | File Sumber | Baris Awal | Baris Bersih | Output |
|---------|------------|:----------:|:------------:|--------|
| Kekayaan Intelektual | `1_Data_KI_FIF_Des25-Jul26.xlsx` | 147 | 147 | `clean_data_ki_fif.csv` |
| Startup FIF | `2_Data_Startup_FIF_2025-2026.xlsx` | 55 (raw) | 53 | `clean_data_startup_fif.csv` |
| Peserta Inkubasi | `3_Peserta_Inkubasi_FIF_2025.xlsx` | 47 (raw) | 44 | `clean_peserta_inkubasi_fif.csv` |

---

## 🔴 Masalah yang Ditemukan & Solusi

### Dataset 1: Data KI FIF (Kekayaan Intelektual)

#### Masalah 1: Data Leakage pada Kolom `objek_komersialisasi`

| Aspek | Detail |
|-------|--------|
| **Masalah** | Kolom `objek_komersialisasi` seharusnya berisi **deskripsi objek** yang dikomersialisasi, tetapi **36 baris** justru berisi nilai boolean "Ya" (19 baris) dan "Tidak" (17 baris) |
| **Dampak** | Informasi boolean tercampur dengan deskripsi teks, menyebabkan analisis komersialisasi tidak akurat |
| **Solusi** | Nilai "Ya"/"Tidak" dipindah ke kolom `sudah_di_komersialisasikan` (jika masih kosong), lalu kolom `objek_komersialisasi` dibersihkan menjadi `null` |
| **Baris Terpengaruh** | 36 baris |

#### Masalah 2: Error `#NAME?` pada Data

| Aspek | Detail |
|-------|--------|
| **Masalah** | Terdapat **1 baris** yang mengandung `#NAME?` — kemungkinan error formula Excel yang tidak ter-resolve |
| **Dampak** | Data corrupt/tidak valid, bisa mengganggu analisis teks |
| **Solusi** | Semua cell yang mengandung `#NAME?` diubah menjadi `null` |
| **Baris Terpengaruh** | 1 baris |

#### Masalah 3: Format Boolean Tidak Konsisten

| Aspek | Detail |
|-------|--------|
| **Masalah** | Kolom `sudah_di_aplikasikan` dan `sudah_di_komersialisasikan` menggunakan format **string** "Y"/"N" bukan tipe boolean |
| **Dampak** | Tidak bisa langsung digunakan untuk filtering/aggregasi boolean |
| **Solusi** | Konversi "Y" → `True`, "N" → `False`, kosong tetap `null` |
| **Baris Terpengaruh** | 138 baris (9 tetap null) |

#### Masalah 4: Inkonsistensi Casing pada Teks Deskriptif

| Aspek | Detail |
|-------|--------|
| **Masalah** | Variasi penulisan: "Perguruan tinggi lainnya" vs "Perguruan Tinggi lainnya", "Universitas telkom" vs "Universitas Telkom" |
| **Dampak** | Aggregasi/grouping menghasilkan kategori duplikat |
| **Solusi** | Normalisasi ke bentuk standar (Title Case) |
| **Baris Terpengaruh** | ~5 baris |

#### Masalah 5: Newline Characters di Dalam Cell

| Aspek | Detail |
|-------|--------|
| **Masalah** | Beberapa cell `objek_komersialisasi` dan `objek_pengaplikasian` mengandung `\n` (newline) |
| **Dampak** | CSV parsing bisa terganggu, tampilan tidak konsisten |
| **Solusi** | Newline diganti dengan separator ` \| ` |

#### Masalah 6: Validasi Silang — Sudah Diaplikasikan tapi Objek Kosong

| Aspek | Detail |
|-------|--------|
| **Masalah** | **4 baris** memiliki `sudah_di_aplikasikan = True` tapi `objek_pengaplikasian` kosong |
| **Dampak** | Data tidak lengkap — klaim sudah diaplikasikan tapi tidak ada bukti objek |
| **Solusi** | Ditandai dengan flag `flag_aplikasi_tanpa_objek = True` agar bisa ditindaklanjuti manual |
| **Baris Terpengaruh** | 4 baris |

---

### Dataset 2: Data Startup FIF 2025-2026

#### Masalah 7: Header Gabungan Excel (Merged Cells)

| Aspek | Detail |
|-------|--------|
| **Masalah** | Baris pertama adalah judul tabel gabungan ("Jumlah Tim Startup FIF Tahun 2025 - 2026"), bukan header kolom. Header asli ada di baris ke-2 |
| **Dampak** | Pandas membaca judul sebagai nama kolom → `Unnamed: 1`, `Unnamed: 2`, dst. |
| **Solusi** | Auto-detect header row, skip judul, assign nama kolom dari baris yang benar |

#### Masalah 8: Kolom `batch_tahun` Tidak Terstruktur

| Aspek | Detail |
|-------|--------|
| **Masalah** | Kolom `Batch/Tahun` berisi campuran format: "Tahun 2025", "Batch 22 Tahun 2025", "Batch 23 Tahun 2026", "Batch 24 Tahun 2026" |
| **Dampak** | Tidak bisa langsung di-filter per batch atau per tahun |
| **Solusi** | Dipecah menjadi 2 kolom terpisah: `batch` (integer/null) dan `tahun` (integer). Startup tahun 2025 tanpa nomor batch → `batch = null` |

#### Masalah 9: Whitespace Berlebih di Deskripsi

| Aspek | Detail |
|-------|--------|
| **Masalah** | Beberapa deskripsi mengandung spasi ganda atau leading/trailing whitespace |
| **Dampak** | Inkonsistensi teks, masalah saat text matching |
| **Solusi** | Collapse multiple whitespace menjadi single space, strip leading/trailing |

---

### Dataset 3: Peserta Inkubasi FIF 2025

#### Masalah 10: Header Gabungan & Baris Non-Data

| Aspek | Detail |
|-------|--------|
| **Masalah** | 2 baris awal berisi judul ("PESERTA PROGRAM INKUBASI TAHUN 2025") dan sub-judul ("FAKULTAS INFORMATIKA"), bukan data. Plus 1 baris header |
| **Dampak** | 3 baris non-data terbaca sebagai record |
| **Solusi** | Auto-detect header, skip baris non-data. 47 baris raw → 44 baris data valid |

#### Masalah 11: Inkonsistensi Nama Prodi & Universitas

| Aspek | Detail |
|-------|--------|
| **Masalah** | Variasi penulisan: `S1 INFORMATIKA` vs `S1 Informatika`, `S1 Informatika - Informatika` vs `S1 Informatika`, `Telkom University` vs `TelU Surabaya` (tanpa kota) |
| **Dampak** | Grouping/aggregasi per prodi menghasilkan kategori duplikat |
| **Solusi** | Normalisasi ke bentuk standar. "Telkom University" → "Telkom University Bandung" agar konsisten dengan "TelU Surabaya" |
| **Baris Terpengaruh** | 5 baris prodi, semua baris univ |

#### Masalah 12: NIM Duplikat & NIM Kosong

| Aspek | Detail |
|-------|--------|
| **Masalah** | **1 NIM duplikat**: `9000005760` (Mahasiswa 5760) terdaftar di **2 startup** sekaligus (FLICKNFIT & Starticas). **1 NIM kosong**: Mahasiswa 1677 (Redooceit) |
| **Dampak** | Duplikat bisa menyebabkan double-counting. NIM kosong menghambat join data |
| **Solusi** | Ditandai dengan flag `flag_nim_duplikat = True`. NIM kosong dibiarkan `null` (tidak di-generate, perlu verifikasi manual) |

---

### Cross-Dataset: Inkonsistensi Antar File

#### Temuan: Startup Tidak Cocok Antar Dataset

| Aspek | Detail |
|-------|--------|
| **Temuan** | **2 startup** ada di data Inkubasi tapi **tidak ada** di data Startup: `Start Smart`, `AplikaPKL` |
| **Kemungkinan** | Startup tersebut belum didaftarkan di data master, atau nama berbeda |
| **Status** | Dilaporkan — tidak ada perubahan otomatis, perlu konfirmasi manual |

> [!IMPORTANT]
> 37 startup dari data Startup FIF **tidak memiliki data peserta inkubasi**. Ini wajar karena data inkubasi hanya mencakup tahun 2025, sementara data startup mencakup 2025-2026 dan kemungkinan tidak semua startup ikut program inkubasi.

---

## ✅ Ringkasan Tindakan

| # | Masalah | Severity | Status | Tindakan |
|:-:|---------|:--------:|:------:|----------|
| 1 | Data leakage objek_komersialisasi | 🔴 Tinggi | ✅ Fixed | Nilai dipindah ke kolom yang benar |
| 2 | Error `#NAME?` | 🔴 Tinggi | ✅ Fixed | Diganti `null` |
| 3 | Format boolean string | 🟡 Sedang | ✅ Fixed | Konversi ke True/False |
| 4 | Inkonsistensi casing | 🟡 Sedang | ✅ Fixed | Normalisasi |
| 5 | Newline dalam cell | 🟢 Rendah | ✅ Fixed | Diganti separator |
| 6 | Aplikasi tanpa objek | 🟡 Sedang | ⚠️ Flagged | Perlu verifikasi manual |
| 7 | Header gabungan Excel | 🟡 Sedang | ✅ Fixed | Auto-detect & skip |
| 8 | Kolom batch_tahun campur | 🟡 Sedang | ✅ Fixed | Dipecah jadi 2 kolom |
| 9 | Whitespace berlebih | 🟢 Rendah | ✅ Fixed | Collapse & trim |
| 10 | Baris non-data | 🟡 Sedang | ✅ Fixed | Auto-detect & skip |
| 11 | Inkonsistensi prodi/univ | 🟡 Sedang | ✅ Fixed | Normalisasi |
| 12 | NIM duplikat & kosong | 🔴 Tinggi | ⚠️ Flagged | Perlu verifikasi manual |

**Legend:** ✅ Fixed = Otomatis diperbaiki | ⚠️ Flagged = Ditandai, perlu tindak lanjut manual

---

## 📁 Struktur Output

```
ETL_DAWER/
├── output/
│   ├── clean_data_ki_fif.csv          (147 baris, 10 kolom)
│   ├── clean_data_startup_fif.csv     (53 baris, 5 kolom)
│   └── clean_peserta_inkubasi_fif.csv (44 baris, 8 kolom)
├── docs/
│   ├── etl_report_20261111_204855.md  (Laporan proses ETL)
│   └── temuan_dan_solusi_20261111_210419.md  (Dokumen ini)
└── code/
    ├── etl_fif_inovasi.py             (Pipeline ETL)
    └── eda_fif_inovasi.ipynb          (Notebook EDA)
```

