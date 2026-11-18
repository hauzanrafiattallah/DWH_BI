"""
Script pembangun notebook DWH — jalankan sekali untuk generate .ipynb
"""

import json

cells = []
exec_counter = [0]  # mutable counter


def md(text: str):
    """Tambah markdown cell."""
    lines = text.split("\n")
    source = [l + "\n" for l in lines[:-1]] + [lines[-1]]
    cells.append({"cell_type": "markdown", "metadata": {}, "source": source})


def code(lines_or_str, comment_header: str = ""):
    """Tambah code cell dari list baris atau string multiline."""
    exec_counter[0] += 1
    if isinstance(lines_or_str, str):
        raw = lines_or_str.strip().split("\n")
    else:
        raw = lines_or_str
    source = [l + "\n" for l in raw[:-1]] + [raw[-1]]
    cells.append(
        {
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": source,
        }
    )


# ═══════════════════════════════════════════════════════════════════════════
#  COVER
# ═══════════════════════════════════════════════════════════════════════════
md("""\
# 🏗️ Data Warehouse — DWH Inovasi & Kewirausahaan FIF
## Enterprise Constellation Schema: 5 Fact + 9 Dimension Tables

**Tanggal:** 2026-11-18  
**Sumber:** output ETL pipeline (`output/*.csv`)  
**Output:** `output/datawarehouse/`  
**Referensi skema:** Enterprise Constellation Schema DWH Inovasi FIF

---

### 🗺️ Arsitektur

```
output/datawarehouse/
├── dimensions/          ← 9 tabel dimensi (CSV)
│   ├── dim_waktu.csv
│   ├── dim_inovasi.csv
│   ├── dim_program_studi.csv
│   ├── dim_mitra.csv
│   ├── dim_jenis_ki.csv
│   ├── dim_status_implementasi_ki.csv
│   ├── dim_tim_startup.csv
│   ├── dim_status_pendanaan.csv
│   └── dim_jenis_kegiatan.csv
├── facts/               ← 5 tabel fakta (CSV)
│   ├── fact_kekayaan_intelektual.csv
│   ├── fact_status_startup.csv
│   ├── fact_kolaborasi_industri.csv
│   ├── fact_hibah_industri.csv
│   └── fact_penelitian_pkm.csv
└── excel/
    ├── dim_waktu.xlsx
    ├── dim_inovasi.xlsx
    ├── ... (satu file per tabel)
    └── star_schema_fif_inovasi.xlsx  ← semua sheet dalam 1 workbook
```

### 📊 Diagram Skema
```
                    dim_waktu (conformed)
                         │
dim_jenis_ki ──► fact_kekayaan_intelektual ◄── dim_inovasi (conformed)
dim_status_ki ──►                         ◄── dim_mitra (conformed)

dim_tim_startup ──► fact_status_startup ◄── dim_waktu
dim_status_pendanaan ►

dim_mitra ──► fact_kolaborasi_industri ◄── dim_inovasi
              fact_hibah_industri      ◄── dim_program_studi (conformed)
              fact_penelitian_pkm     ◄── dim_jenis_kegiatan
```
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 0 — INISIALISASI
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 📦 Phase 0 — Inisialisasi

### 0.1 Import Library
""")

code("""\
# ── Standard library ─────────────────────────────────────────────────────
import pathlib
import datetime

# ── Data processing ───────────────────────────────────────────────────────
import pandas as pd
import numpy as np

# ── Excel export ──────────────────────────────────────────────────────────
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

print("✅ Library berhasil diimport")
print(f"   pandas  : {pd.__version__}")
print(f"   numpy   : {np.__version__}")
print(f"   openpyxl: {openpyxl.__version__}")
""")

md("### 0.2 Setup Path & Folder Output")

code("""\
# ── Root project ──────────────────────────────────────────────────────────
ROOT = pathlib.Path("..").resolve()          # satu level di atas folder code/
OUT_CLEAN = ROOT / "output"                  # CSV bersih dari ETL pipeline
DWH = ROOT / "output" / "datawarehouse"     # folder DWH utama
DIM_DIR = DWH / "dimensions"                # subfolder dimensi
FACT_DIR = DWH / "facts"                    # subfolder fakta
XL_DIR = DWH / "excel"                      # subfolder excel

# Pastikan semua folder ada
for folder in [DWH, DIM_DIR, FACT_DIR, XL_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

print("📁 Struktur folder output:")
for f in [DWH, DIM_DIR, FACT_DIR, XL_DIR]:
    print(f"   {f.relative_to(ROOT)}")
print("\\n✅ Semua folder siap")
""")

md("### 0.3 Load Data Sumber (CSV Bersih dari ETL Pipeline)")

code("""\
# ── Load 3 dataset bersih ─────────────────────────────────────────────────
df_ki  = pd.read_csv(OUT_CLEAN / "clean_data_ki_fif.csv")
df_st  = pd.read_csv(OUT_CLEAN / "clean_data_startup_fif.csv")
df_ink = pd.read_csv(OUT_CLEAN / "clean_peserta_inkubasi_fif.csv")

# Tampilkan ringkasan
for name, df in [("KI FIF", df_ki), ("Startup FIF", df_st), ("Inkubasi FIF", df_ink)]:
    print(f"📄 {name:15s} → {df.shape[0]:3d} baris × {df.shape[1]} kolom  "
          f"| kolom: {list(df.columns)}")
""")

md("### 0.4 Pra-proses Tipe Data")

code("""\
# ── KI: konversi boolean yang terbaca sebagai string/object ───────────────
bool_map = {"True": True, "False": False, True: True, False: False}
df_ki["sudah_di_aplikasikan"]       = df_ki["sudah_di_aplikasikan"].map(bool_map)
df_ki["sudah_di_komersialisasikan"] = df_ki["sudah_di_komersialisasikan"].map(bool_map)
df_ki["flag_aplikasi_tanpa_objek"]  = df_ki["flag_aplikasi_tanpa_objek"].astype(bool)

# ── Inkubasi: NIM dari float → string bersih ──────────────────────────────
df_ink["nim"] = df_ink["nim"].apply(
    lambda x: str(int(x)) if pd.notna(x) else None
)
df_ink["flag_nim_duplikat"] = df_ink["flag_nim_duplikat"].astype(bool)

# ── Startup: batch dari float → Int64 (nullable integer) ─────────────────
df_st["batch"] = df_st["batch"].astype("Int64")

print("✅ Tipe data sudah dikonversi")
print("\\n--- KI dtypes ---")
print(df_ki.dtypes)
print("\\n--- Startup dtypes ---")
print(df_st.dtypes)
print("\\n--- Inkubasi dtypes ---")
print(df_ink.dtypes)
""")

# ═══════════════════════════════════════════════════════════════════════════
#  HELPER
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 🛠️ Phase 0.5 — Helper Functions

Fungsi bantu untuk: (1) menyimpan CSV, (2) menyimpan Excel per tabel dengan styling, \
(3) format sheet di workbook gabungan.
""")

code("""\
# ═════════════════════════════════════════════════════════════════════════
# HELPER: simpan CSV
# ═════════════════════════════════════════════════════════════════════════
def save_csv(df: pd.DataFrame, folder: pathlib.Path, filename: str) -> pathlib.Path:
    \"\"\"Simpan DataFrame ke CSV dan cetak konfirmasi.\"\"\"
    path = folder / filename
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"   💾 CSV  → {path.relative_to(ROOT)}  ({len(df)} baris)")
    return path


# ═════════════════════════════════════════════════════════════════════════
# HELPER: styling Excel
# ═════════════════════════════════════════════════════════════════════════

# Palet warna tema
THEME = {
    "dim_conformed": "FF7F3E",   # orange  → dimensi conformed
    "dim_local"    : "4CAF50",   # hijau   → dimensi local
    "fact"         : "1E3A5F",   # biru tua → tabel fakta
    "header_text"  : "FFFFFF",   # putih   → teks header
    "alt_row"      : "EAF4FB",   # biru muda → baris ganjil
}

BORDER_THIN = Border(
    left=Side(style="thin", color="BDBDBD"),
    right=Side(style="thin", color="BDBDBD"),
    top=Side(style="thin", color="BDBDBD"),
    bottom=Side(style="thin", color="BDBDBD"),
)


def style_worksheet(ws, header_color: str, freeze: bool = True):
    \"\"\"
    Terapkan styling ke worksheet openpyxl:
    - Header tebal + fill warna + border
    - Baris alternating fill
    - Auto-width kolom
    - Freeze baris pertama
    \"\"\"
    header_fill = PatternFill("solid", fgColor=header_color)
    alt_fill    = PatternFill("solid", fgColor=THEME["alt_row"])
    header_font = Font(bold=True, color=THEME["header_text"], size=11)
    normal_font = Font(size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    col_widths = {}

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            # Border semua cell
            cell.border = BORDER_THIN

            if row_idx == 1:
                # Header row
                cell.fill      = header_fill
                cell.font      = header_font
                cell.alignment = center_align
            else:
                # Data rows — alternating
                if row_idx % 2 == 0:
                    cell.fill = alt_fill
                cell.font      = normal_font
                cell.alignment = left_align

            # Hitung lebar kolom ideal
            col_letter = cell.column_letter
            val_len = len(str(cell.value)) if cell.value is not None else 0
            col_widths[col_letter] = max(col_widths.get(col_letter, 10), min(val_len + 4, 50))

    # Terapkan lebar kolom
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Freeze row pertama (header)
    if freeze:
        ws.freeze_panes = "A2"


def save_excel_single(df: pd.DataFrame, folder: pathlib.Path,
                      filename: str, sheet_name: str,
                      header_color: str, table_type: str = ""):
    \"\"\"Simpan satu DataFrame ke Excel file individual dengan styling.\"\"\"
    path = folder / filename
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Tulis data
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    style_worksheet(ws, header_color)

    # Metadata sheet
    ws_meta = wb.create_sheet("_Info")
    ws_meta["A1"] = "Tabel"
    ws_meta["B1"] = sheet_name
    ws_meta["A2"] = "Tipe"
    ws_meta["B2"] = table_type
    ws_meta["A3"] = "Jumlah Baris"
    ws_meta["B3"] = len(df)
    ws_meta["A4"] = "Jumlah Kolom"
    ws_meta["B4"] = len(df.columns)
    ws_meta["A5"] = "Dibuat"
    ws_meta["B5"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_meta["A6"] = "Kolom"
    ws_meta["B6"] = ", ".join(df.columns)

    wb.save(path)
    print(f"   📊 XLSX → {path.relative_to(ROOT)}  ({len(df)} baris)")
    return path


def add_to_master_wb(wb_master, df: pd.DataFrame, sheet_name: str, header_color: str):
    \"\"\"Tambahkan sheet ke workbook master.\"\"\"
    ws = wb_master.create_sheet(title=sheet_name)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    style_worksheet(ws, header_color)


print("✅ Helper functions siap")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 1 — DIMENSION TABLES
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 🧊 Phase 1 — Dimension Tables

> **Urutan pembuatan:** Conformed dimensions dulu, kemudian local dimensions.  
> Setiap dimensi memiliki surrogate key (`id_xxx`) sebagai Primary Key.

### Conformed Dimensions (berbagi antar Fact Table)
- `dim_waktu` — dimensi waktu akademik
- `dim_inovasi` — judul & kategori inovasi/KI
- `dim_program_studi` — prodi, jenjang, kampus
- `dim_mitra` — institusi mitra/aplikasi KI

### Local Dimensions (spesifik per Business Process)
- `dim_jenis_ki` — jenis KI + dasar hukum
- `dim_status_implementasi_ki` — status DJKI
- `dim_tim_startup` — identitas tim startup
- `dim_status_pendanaan` — batch & status inkubasi
- `dim_jenis_kegiatan` — kategori kegiatan (placeholder)
""")

# 1a — dim_waktu
md("### 1.1 `dim_waktu` — Dimensi Waktu Akademik _(Conformed — BP1 s/d BP5)_")

code("""\
# ── dim_waktu ──────────────────────────────────────────────────────────────
# Sumber: derived dari tahun yang ada di data startup & KI
# Grain  : 1 baris = 1 periode semester akademik (Ganjil/Genap)
# ─────────────────────────────────────────────────────────────────────────

# Tahun akademik yang diketahui dari data
tahun_list = sorted(df_st["tahun"].dropna().unique().tolist())  # [2025, 2026]

rows_waktu = []
for tahun in tahun_list:
    for semester_no, semester_nama in [(1, "Ganjil"), (2, "Genap")]:
        rows_waktu.append({
            "bulan"           : 1 if semester_no == 1 else 7,
            "semester"        : semester_nama,
            "tahun"           : int(tahun),
            "tahun_akademik"  : f"{tahun}/{tahun + 1}",
            "triwulan"        : "TW1" if semester_no == 1 else "TW3",
        })

dim_waktu = pd.DataFrame(rows_waktu)
# Surrogate key
dim_waktu.insert(0, "id_waktu", range(1, len(dim_waktu) + 1))

print("📋 dim_waktu")
print(dim_waktu.to_string(index=False))
print(f"\\nShape: {dim_waktu.shape}")
""")

code("""\
# ── Simpan dim_waktu ──────────────────────────────────────────────────────
print("💾 Menyimpan dim_waktu...")
save_csv(dim_waktu, DIM_DIR, "dim_waktu.csv")
save_excel_single(dim_waktu, XL_DIR, "dim_waktu.xlsx",
                  sheet_name="dim_waktu",
                  header_color=THEME["dim_conformed"],
                  table_type="Conformed Dimension")
print("\\n✅ dim_waktu selesai")
""")

# 1b — dim_inovasi
md("### 1.2 `dim_inovasi` — Dimensi Inovasi/KI _(Conformed — BP1, BP2, BP3)_")

code("""\
# ── dim_inovasi ────────────────────────────────────────────────────────────
# Sumber: df_ki (kolom judul, kategori)
# Grain  : 1 baris = 1 karya inovasi/KI unik
# Bidang inovasi diturunkan dari kategori KI
# ─────────────────────────────────────────────────────────────────────────

# Mapping kategori → bidang inovasi
bidang_map = {
    "HAK CIPTA"      : "Perangkat Lunak & Karya Cipta",
    "PATEN"          : "Teknologi & Rekayasa",
    "DESAIN INDUSTRI": "Desain & Estetika Produk",
    "MEREK"          : "Identitas Usaha & Merek",
}

dim_inovasi = df_ki[["no", "judul", "kategori"]].copy()
dim_inovasi = dim_inovasi.rename(columns={
    "no"      : "id_inovasi",
    "judul"   : "judul_inovasi",
    "kategori": "kategori_inovasi",
})
dim_inovasi["bidang_inovasi"] = dim_inovasi["kategori_inovasi"].map(bidang_map)

# Kolom akhir sesuai skema
dim_inovasi = dim_inovasi[["id_inovasi", "judul_inovasi", "kategori_inovasi", "bidang_inovasi"]]

print("📋 dim_inovasi (5 baris pertama)")
print(dim_inovasi.head(5).to_string(index=False))
print(f"\\nShape: {dim_inovasi.shape}")
print("\\nDistribusi kategori_inovasi:")
print(dim_inovasi["kategori_inovasi"].value_counts().to_string())
""")

code("""\
# ── Simpan dim_inovasi ────────────────────────────────────────────────────
print("💾 Menyimpan dim_inovasi...")
save_csv(dim_inovasi, DIM_DIR, "dim_inovasi.csv")
save_excel_single(dim_inovasi, XL_DIR, "dim_inovasi.xlsx",
                  sheet_name="dim_inovasi",
                  header_color=THEME["dim_conformed"],
                  table_type="Conformed Dimension")
print("\\n✅ dim_inovasi selesai")
""")

# 1c — dim_program_studi
md("### 1.3 `dim_program_studi` — Dimensi Program Studi _(Conformed — BP1, BP3)_")

code("""\
# ── dim_program_studi ─────────────────────────────────────────────────────
# Sumber: df_ink (kolom prodi, universitas, fakultas)
# Grain  : 1 baris = 1 kombinasi unik program studi + kampus
# ─────────────────────────────────────────────────────────────────────────

# Ambil kombinasi unik prodi × kampus
prodi_uniq = (
    df_ink[["prodi", "fakultas", "universitas"]]
    .drop_duplicates()
    .sort_values(["universitas", "prodi"])
    .reset_index(drop=True)
)

# Mapping jenjang dari nama prodi
def parse_jenjang(prodi_nama: str) -> str:
    if prodi_nama.upper().startswith("S1"):
        return "Sarjana (S1)"
    elif prodi_nama.upper().startswith("D4"):
        return "Sarjana Terapan (D4)"
    elif prodi_nama.upper().startswith("S2"):
        return "Magister (S2)"
    return "Lainnya"

# Kode prodi: singkatan dari nama prodi
kode_map = {
    "S1 Informatika"              : "IF",
    "S1 Teknologi Informasi"      : "TI",
    "S1 Rekayasa Perangkat Lunak" : "RPL",
    "S1 Sains Data"               : "SD",
}

dim_program_studi = prodi_uniq.copy()
dim_program_studi.insert(0, "id_program_studi", range(1, len(dim_program_studi) + 1))
dim_program_studi["kode_prodi"]  = dim_program_studi["prodi"].map(kode_map).fillna("XX")
dim_program_studi["jenjang"]     = dim_program_studi["prodi"].apply(parse_jenjang)
dim_program_studi["nama_prodi"]  = dim_program_studi["prodi"]
dim_program_studi["kampus"]      = dim_program_studi["universitas"]

dim_program_studi = dim_program_studi[[
    "id_program_studi", "kode_prodi", "nama_prodi",
    "jenjang", "fakultas", "kampus"
]]

print("📋 dim_program_studi")
print(dim_program_studi.to_string(index=False))
print(f"\\nShape: {dim_program_studi.shape}")
""")

code("""\
# ── Simpan dim_program_studi ──────────────────────────────────────────────
print("💾 Menyimpan dim_program_studi...")
save_csv(dim_program_studi, DIM_DIR, "dim_program_studi.csv")
save_excel_single(dim_program_studi, XL_DIR, "dim_program_studi.xlsx",
                  sheet_name="dim_program_studi",
                  header_color=THEME["dim_conformed"],
                  table_type="Conformed Dimension")
print("\\n✅ dim_program_studi selesai")
""")

# 1d — dim_mitra
md("### 1.4 `dim_mitra` — Dimensi Mitra Industri/Institusi _(Conformed — BP1, BP2)_")

code("""\
# ── dim_mitra ─────────────────────────────────────────────────────────────
# Sumber: df_ki (kolom objek_pengaplikasian) — setiap entri objek = 1 mitra
# Grain  : 1 baris = 1 mitra/institusi unik
# Note   : Data mitra hanya dari KI yang sudah diaplikasikan
# ─────────────────────────────────────────────────────────────────────────

# Mapping sektor dari nama mitra (rule-based sederhana)
def infer_sektor(nama: str) -> str:
    n = nama.lower()
    if any(k in n for k in ["universitas", "telkom", "telu", "mahasiswa", "kampus",
                             "fakultas", "prodi", "akademi", "sekolah tinggi"]):
        return "Pendidikan Tinggi"
    elif any(k in n for k in ["smp", "sd", "sma", "smk", "smk", "sekolah", "tk"]):
        return "Pendidikan Dasar & Menengah"
    elif any(k in n for k in ["pt ", "cv ", "peruri", "ptpn", "palmco", "industri",
                              "manufaktur", "decor", "aiesha", "mobilindo"]):
        return "Industri & Manufaktur"
    elif any(k in n for k in ["pemerintah", "desa", "kelurahan", "kabupaten", "pkk",
                              "bbksda"]):
        return "Pemerintahan & Publik"
    elif any(k in n for k in ["pondok", "pesantren", "yayasan"]):
        return "Sosial & Keagamaan"
    elif any(k in n for k in ["laundry", "warung", "masyarakat", "properti"]):
        return "UMKM & Masyarakat"
    return "Lainnya"

def infer_skala(sektor: str) -> str:
    if sektor in ("Industri & Manufaktur",):
        return "Besar"
    elif sektor in ("Pendidikan Tinggi", "Pemerintahan & Publik"):
        return "Menengah"
    return "Kecil"

def infer_wilayah(nama: str) -> str:
    n = nama.lower()
    if any(k in n for k in ["surabaya", "jatim", "jawa timur"]):
        return "Jawa Timur"
    elif any(k in n for k in ["bandung", "pangandaran", "garut", "nagreg",
                              "cimaragas", "banjarsari", "cipamokolan"]):
        return "Jawa Barat"
    elif any(k in n for k in ["jakarta"]):
        return "DKI Jakarta"
    return "Tidak Diketahui"

# Kumpulkan semua nama mitra dari objek_pengaplikasian (non-null)
mitra_raw = (
    df_ki["objek_pengaplikasian"]
    .dropna()
    .unique()
    .tolist()
)

rows_mitra = []
for nama in sorted(mitra_raw):
    sektor = infer_sektor(nama)
    rows_mitra.append({
        "nama_mitra"    : nama,
        "sektor_industri": sektor,
        "skala_mitra"   : infer_skala(sektor),
        "wilayah"       : infer_wilayah(nama),
    })

dim_mitra = pd.DataFrame(rows_mitra)
dim_mitra.insert(0, "id_mitra", range(1, len(dim_mitra) + 1))

print("📋 dim_mitra")
print(dim_mitra.to_string(index=False))
print(f"\\nShape: {dim_mitra.shape}")
print("\\nDistribusi sektor:")
print(dim_mitra["sektor_industri"].value_counts().to_string())
""")

code("""\
# ── Simpan dim_mitra ──────────────────────────────────────────────────────
print("💾 Menyimpan dim_mitra...")
save_csv(dim_mitra, DIM_DIR, "dim_mitra.csv")
save_excel_single(dim_mitra, XL_DIR, "dim_mitra.xlsx",
                  sheet_name="dim_mitra",
                  header_color=THEME["dim_conformed"],
                  table_type="Conformed Dimension")
print("\\n✅ dim_mitra selesai")
""")

# 1e — dim_jenis_ki
md("### 1.5 `dim_jenis_ki` — Dimensi Jenis KI _(Local — BP4)_")

code("""\
# ── dim_jenis_ki ──────────────────────────────────────────────────────────
# Sumber: statis — 4 jenis KI yang diakui DJKI
# Grain  : 1 baris = 1 jenis KI
# ─────────────────────────────────────────────────────────────────────────

dim_jenis_ki = pd.DataFrame([
    {
        "id_jenis_ki" : 1,
        "kode_jenis"  : "HC",
        "nama_jenis_ki": "HAK CIPTA",
        "dasar_hukum" : "UU No. 28 Tahun 2014 tentang Hak Cipta",
    },
    {
        "id_jenis_ki" : 2,
        "kode_jenis"  : "PAT",
        "nama_jenis_ki": "PATEN",
        "dasar_hukum" : "UU No. 13 Tahun 2016 tentang Paten",
    },
    {
        "id_jenis_ki" : 3,
        "kode_jenis"  : "DI",
        "nama_jenis_ki": "DESAIN INDUSTRI",
        "dasar_hukum" : "UU No. 31 Tahun 2000 tentang Desain Industri",
    },
    {
        "id_jenis_ki" : 4,
        "kode_jenis"  : "MRK",
        "nama_jenis_ki": "MEREK",
        "dasar_hukum" : "UU No. 20 Tahun 2016 tentang Merek dan Indikasi Geografis",
    },
])

print("📋 dim_jenis_ki")
print(dim_jenis_ki.to_string(index=False))
""")

code("""\
# ── Simpan dim_jenis_ki ───────────────────────────────────────────────────
print("💾 Menyimpan dim_jenis_ki...")
save_csv(dim_jenis_ki, DIM_DIR, "dim_jenis_ki.csv")
save_excel_single(dim_jenis_ki, XL_DIR, "dim_jenis_ki.xlsx",
                  sheet_name="dim_jenis_ki",
                  header_color=THEME["dim_local"],
                  table_type="Local Dimension — BP4")
print("\\n✅ dim_jenis_ki selesai")
""")

# 1f — dim_status_implementasi_ki
md("### 1.6 `dim_status_implementasi_ki` — Status Implementasi KI _(Local — BP4)_")

code("""\
# ── dim_status_implementasi_ki ────────────────────────────────────────────
# Sumber: statis — status yang ada di data KI + status aplikasi/komersialisasi
# Grain  : 1 baris = 1 status unik
# ─────────────────────────────────────────────────────────────────────────

dim_status_implementasi_ki = pd.DataFrame([
    {"id_status_implementasi": 1, "kode_status": "SELESAI",
     "nama_status": "Sertifikat Selesai",
     "deskripsi"  : "KI telah mendapatkan sertifikat resmi dari DJKI"},
    {"id_status_implementasi": 2, "kode_status": "PENGAJUAN",
     "nama_status": "Pengajuan DJKI",
     "deskripsi"  : "KI sedang dalam proses pengajuan ke DJKI"},
    {"id_status_implementasi": 3, "kode_status": "APL_Y",
     "nama_status": "Sudah Diaplikasikan",
     "deskripsi"  : "KI telah diimplementasikan pada mitra/industri"},
    {"id_status_implementasi": 4, "kode_status": "APL_N",
     "nama_status": "Belum Diaplikasikan",
     "deskripsi"  : "KI belum diimplementasikan"},
    {"id_status_implementasi": 5, "kode_status": "KOM_Y",
     "nama_status": "Sudah Dikomersialisasikan",
     "deskripsi"  : "KI telah menghasilkan nilai ekonomi"},
    {"id_status_implementasi": 6, "kode_status": "KOM_N",
     "nama_status": "Belum Dikomersialisasikan",
     "deskripsi"  : "KI belum menghasilkan nilai ekonomi"},
])

print("📋 dim_status_implementasi_ki")
print(dim_status_implementasi_ki.to_string(index=False))
""")

code("""\
# ── Simpan dim_status_implementasi_ki ────────────────────────────────────
print("💾 Menyimpan dim_status_implementasi_ki...")
save_csv(dim_status_implementasi_ki, DIM_DIR, "dim_status_implementasi_ki.csv")
save_excel_single(dim_status_implementasi_ki, XL_DIR,
                  "dim_status_implementasi_ki.xlsx",
                  sheet_name="dim_status_impl_ki",
                  header_color=THEME["dim_local"],
                  table_type="Local Dimension — BP4")
print("\\n✅ dim_status_implementasi_ki selesai")
""")

# 1g — dim_tim_startup
md("### 1.7 `dim_tim_startup` — Dimensi Tim Startup _(Local — BP5)_")

code("""\
# ── dim_tim_startup ───────────────────────────────────────────────────────
# Sumber: df_st (nama startup, batch, deskripsi) + df_ink (prodi mayoritas)
# Grain  : 1 baris = 1 tim startup unik
# ─────────────────────────────────────────────────────────────────────────

# Hitung prodi mayoritas per startup dari data inkubasi
prodi_mayoritas = (
    df_ink.groupby("nama_startup")["prodi"]
    .agg(lambda x: x.value_counts().index[0] if len(x) > 0 else None)
    .reset_index()
    .rename(columns={"prodi": "prodi_mayoritas"})
)

# Inferensi bidang usaha dari deskripsi
def infer_bidang_usaha(deskripsi: str) -> str:
    d = str(deskripsi).lower()
    if any(k in d for k in ["ai", "artificial intelligence", "machine learning",
                             "data", "algoritma"]):
        return "Teknologi & AI"
    elif any(k in d for k in ["kesehatan", "klinik", "medis", "health", "dokter"]):
        return "Kesehatan (HealthTech)"
    elif any(k in d for k in ["edukasi", "pendidikan", "belajar", "sekolah",
                              "mahasiswa", "pembelajaran"]):
        return "Pendidikan (EduTech)"
    elif any(k in d for k in ["makanan", "minuman", "food", "kopi", "juice",
                              "meal", "coffee"]):
        return "Food & Beverage"
    elif any(k in d for k in ["pertanian", "sawit", "farm", "agri", "tanam"]):
        return "Agrikultur (AgriTech)"
    elif any(k in d for k in ["lingkungan", "sampah", "energi", "solar", "iot"]):
        return "Lingkungan & IoT"
    elif any(k in d for k in ["disabilitas", "inklusif", "sosial", "ketenagakerjaan"]):
        return "Sosial & Inklusi"
    elif any(k in d for k in ["platform", "marketplace", "aplikasi", "mobile", "web"]):
        return "Platform Digital"
    return "Lainnya"

# Gabung startup + prodi mayoritas
dim_tim_startup = df_st[["nama_startup", "batch", "tahun", "deskripsi"]].copy()
dim_tim_startup = dim_tim_startup.merge(prodi_mayoritas, on="nama_startup", how="left")
dim_tim_startup["bidang_usaha"]    = dim_tim_startup["deskripsi"].apply(infer_bidang_usaha)
dim_tim_startup["batch_pembinaan"] = dim_tim_startup["batch"].apply(
    lambda x: f"Batch {int(x)}" if pd.notna(x) else "Pre-Batch"
)
dim_tim_startup = dim_tim_startup.rename(columns={
    "nama_startup" : "nama_tim",
    "tahun"        : "tahun_terdaftar",
})
dim_tim_startup.insert(0, "id_tim_startup", range(1, len(dim_tim_startup) + 1))
dim_tim_startup = dim_tim_startup[[
    "id_tim_startup", "nama_tim", "bidang_usaha",
    "batch_pembinaan", "tahun_terdaftar", "prodi_mayoritas"
]]

print("📋 dim_tim_startup (10 baris)")
print(dim_tim_startup.head(10).to_string(index=False))
print(f"\\nShape: {dim_tim_startup.shape}")
print("\\nDistribusi bidang_usaha:")
print(dim_tim_startup["bidang_usaha"].value_counts().to_string())
""")

code("""\
# ── Simpan dim_tim_startup ────────────────────────────────────────────────
print("💾 Menyimpan dim_tim_startup...")
save_csv(dim_tim_startup, DIM_DIR, "dim_tim_startup.csv")
save_excel_single(dim_tim_startup, XL_DIR, "dim_tim_startup.xlsx",
                  sheet_name="dim_tim_startup",
                  header_color=THEME["dim_local"],
                  table_type="Local Dimension — BP5")
print("\\n✅ dim_tim_startup selesai")
""")

# 1h — dim_status_pendanaan
md("### 1.8 `dim_status_pendanaan` — Status Pendanaan/Inkubasi _(Local — BP5)_")

code("""\
# ── dim_status_pendanaan ──────────────────────────────────────────────────
# Sumber: statis — status pendanaan/inkubasi startup
# Grain  : 1 baris = 1 status pendanaan unik
# ─────────────────────────────────────────────────────────────────────────

dim_status_pendanaan = pd.DataFrame([
    {"id_status_pendanaan": 1, "kode_status": "INKUBASI_AKTIF",
     "nama_status": "Aktif Inkubasi",
     "deskripsi"  : "Startup sedang aktif mengikuti program inkubasi FIF"},
    {"id_status_pendanaan": 2, "kode_status": "INKUBASI_SELESAI",
     "nama_status": "Selesai Inkubasi",
     "deskripsi"  : "Startup telah menyelesaikan program inkubasi"},
    {"id_status_pendanaan": 3, "kode_status": "MANDIRI",
     "nama_status": "Mandiri (Pre-Batch)",
     "deskripsi"  : "Startup terdaftar sebelum sistem batch, status mandiri"},
    {"id_status_pendanaan": 4, "kode_status": "BATCH_22",
     "nama_status": "Batch 22 (2025)",
     "deskripsi"  : "Startup dalam gelombang Batch 22 tahun 2025"},
    {"id_status_pendanaan": 5, "kode_status": "BATCH_23",
     "nama_status": "Batch 23 (2026)",
     "deskripsi"  : "Startup dalam gelombang Batch 23 tahun 2026"},
    {"id_status_pendanaan": 6, "kode_status": "BATCH_24",
     "nama_status": "Batch 24 (2026)",
     "deskripsi"  : "Startup dalam gelombang Batch 24 tahun 2026"},
])

print("📋 dim_status_pendanaan")
print(dim_status_pendanaan.to_string(index=False))
""")

code("""\
# ── Simpan dim_status_pendanaan ───────────────────────────────────────────
print("💾 Menyimpan dim_status_pendanaan...")
save_csv(dim_status_pendanaan, DIM_DIR, "dim_status_pendanaan.csv")
save_excel_single(dim_status_pendanaan, XL_DIR, "dim_status_pendanaan.xlsx",
                  sheet_name="dim_status_pendanaan",
                  header_color=THEME["dim_local"],
                  table_type="Local Dimension — BP5")
print("\\n✅ dim_status_pendanaan selesai")
""")

# 1i — dim_jenis_kegiatan
md("### 1.9 `dim_jenis_kegiatan` — Dimensi Jenis Kegiatan _(Local — BP3)_")

code("""\
# ── dim_jenis_kegiatan ────────────────────────────────────────────────────
# Sumber: statis — kategori kegiatan penelitian & PKM
# Grain  : 1 baris = 1 jenis kegiatan
# ─────────────────────────────────────────────────────────────────────────

dim_jenis_kegiatan = pd.DataFrame([
    {"id_jenis_kegiatan": 1,  "kode_jenis": "PKM-RE",
     "nama_jenis": "PKM Riset Eksakta"},
    {"id_jenis_kegiatan": 2,  "kode_jenis": "PKM-RSH",
     "nama_jenis": "PKM Riset Sosial Humaniora"},
    {"id_jenis_kegiatan": 3,  "kode_jenis": "PKM-K",
     "nama_jenis": "PKM Kewirausahaan"},
    {"id_jenis_kegiatan": 4,  "kode_jenis": "PKM-PM",
     "nama_jenis": "PKM Pengabdian Masyarakat"},
    {"id_jenis_kegiatan": 5,  "kode_jenis": "PKM-PI",
     "nama_jenis": "PKM Penerapan Iptek"},
    {"id_jenis_kegiatan": 6,  "kode_jenis": "PKM-KC",
     "nama_jenis": "PKM Karsa Cipta"},
    {"id_jenis_kegiatan": 7,  "kode_jenis": "RISET-INT",
     "nama_jenis": "Penelitian Internal (DRTPM)"},
    {"id_jenis_kegiatan": 8,  "kode_jenis": "RISET-EKS",
     "nama_jenis": "Penelitian Eksternal / Hibah"},
    {"id_jenis_kegiatan": 9,  "kode_jenis": "INKUBASI",
     "nama_jenis": "Program Inkubasi Startup"},
    {"id_jenis_kegiatan": 10, "kode_jenis": "KOMPETISI",
     "nama_jenis": "Kompetisi Inovasi"},
])

print("📋 dim_jenis_kegiatan")
print(dim_jenis_kegiatan.to_string(index=False))
""")

code("""\
# ── Simpan dim_jenis_kegiatan ─────────────────────────────────────────────
print("💾 Menyimpan dim_jenis_kegiatan...")
save_csv(dim_jenis_kegiatan, DIM_DIR, "dim_jenis_kegiatan.csv")
save_excel_single(dim_jenis_kegiatan, XL_DIR, "dim_jenis_kegiatan.xlsx",
                  sheet_name="dim_jenis_kegiatan",
                  header_color=THEME["dim_local"],
                  table_type="Local Dimension — BP3")
print("\\n✅ dim_jenis_kegiatan selesai")
""")

md("### 1.10 Ringkasan Semua Dimensi")

code("""\
# ── Ringkasan dimensi yang telah dibuat ───────────────────────────────────
print("=" * 65)
print(f"{'TABEL':<35} {'TIPE':<15} {'BARIS':>5} {'KOLOM':>5}")
print("=" * 65)

all_dims = [
    (dim_waktu,                  "Conformed"),
    (dim_inovasi,                "Conformed"),
    (dim_program_studi,          "Conformed"),
    (dim_mitra,                  "Conformed"),
    (dim_jenis_ki,               "Local BP4"),
    (dim_status_implementasi_ki, "Local BP4"),
    (dim_tim_startup,            "Local BP5"),
    (dim_status_pendanaan,       "Local BP5"),
    (dim_jenis_kegiatan,         "Local BP3"),
]

dim_names = [
    "dim_waktu", "dim_inovasi", "dim_program_studi", "dim_mitra",
    "dim_jenis_ki", "dim_status_implementasi_ki",
    "dim_tim_startup", "dim_status_pendanaan", "dim_jenis_kegiatan",
]

for name, (df, tipe) in zip(dim_names, all_dims):
    print(f"  {name:<33} {tipe:<15} {len(df):>5} {len(df.columns):>5}")

print("=" * 65)
print(f"  {'TOTAL':<33} {'':15} {sum(len(d) for d, _ in all_dims):>5}")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 2 — FACT TABLES
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 📊 Phase 2 — Fact Tables

> Setiap fact table berisi Foreign Key ke dimensi terkait + measure.  
> Semua PK fact table menggunakan surrogate key `id_fakta_xxx`.

| Fact Table | Business Process | Tipe Grain |
|---|---|---|
| `fact_kekayaan_intelektual` | BP4 | Accumulating Snapshot |
| `fact_status_startup` | BP5 | Periodic Snapshot |
| `fact_kolaborasi_industri` | BP1 | Transaction |
| `fact_hibah_industri` | BP2 | Transaction |
| `fact_penelitian_pkm` | BP3 | Transaction |
""")

# Lookup tables
md("### 2.0 Buat Lookup Table (FK Resolver)")

code("""\
# ── Lookup dict untuk resolve FK ─────────────────────────────────────────
# Mapping dari nilai natural key → surrogate key

# dim_waktu: (tahun, semester) → id_waktu
waktu_lookup = {
    (row["tahun"], row["semester"]): row["id_waktu"]
    for _, row in dim_waktu.iterrows()
}

# dim_inovasi: id_KI asli (kolom 'no') → id_inovasi
inovasi_lookup = dict(zip(dim_inovasi["id_inovasi"], dim_inovasi["id_inovasi"]))
# (id_inovasi = id asli dari KI, sudah sama)

# dim_jenis_ki: nama kategori → id_jenis_ki
jenis_ki_lookup = dict(zip(dim_jenis_ki["nama_jenis_ki"], dim_jenis_ki["id_jenis_ki"]))

# dim_status_implementasi_ki: nama_status → id_status
status_ki_lookup = dict(zip(
    dim_status_implementasi_ki["nama_status"],
    dim_status_implementasi_ki["id_status_implementasi"]
))

# dim_mitra: nama_mitra → id_mitra
mitra_lookup = dict(zip(dim_mitra["nama_mitra"], dim_mitra["id_mitra"]))

# dim_tim_startup: nama_tim → id_tim_startup
tim_lookup = dict(zip(dim_tim_startup["nama_tim"], dim_tim_startup["id_tim_startup"]))

# dim_status_pendanaan: batch + tahun → id_status_pendanaan
def resolve_status_pendanaan(batch, tahun) -> int:
    if pd.isna(batch):
        return 3   # MANDIRI
    b = int(batch)
    if b == 22:  return 4
    elif b == 23: return 5
    elif b == 24: return 6
    return 3

# dim_program_studi: (prodi, kampus) → id_program_studi
prodi_lookup = {
    (row["nama_prodi"], row["kampus"]): row["id_program_studi"]
    for _, row in dim_program_studi.iterrows()
}

print("✅ Semua lookup table siap")
print(f"   waktu_lookup     : {len(waktu_lookup)} entry")
print(f"   jenis_ki_lookup  : {len(jenis_ki_lookup)} entry")
print(f"   mitra_lookup     : {len(mitra_lookup)} entry")
print(f"   tim_lookup       : {len(tim_lookup)} entry")
print(f"   prodi_lookup     : {len(prodi_lookup)} entry")
""")

# fact_kekayaan_intelektual
md("""\
### 2.1 `fact_kekayaan_intelektual` — Accumulating Snapshot _(BP4)_

Grain: **1 baris = 1 karya KI** dengan snapshot status terkini  
Measure: `jumlah_ki` (always=1), `indikator_terimplementasi`, `selang_waktu_impl_hari`
""")

code("""\
# ── fact_kekayaan_intelektual ─────────────────────────────────────────────
# Sumber: df_ki (semua baris)
# Kolom FK: id_waktu_pendaftaran, id_waktu_implementasi, id_jenis_ki,
#            id_status_implementasi, id_inovasi (opsional), id_mitra (opsional)
# Measure : jumlah_ki, indikator_terimplementasi, selang_waktu_impl_hari
# ─────────────────────────────────────────────────────────────────────────

rows_fact_ki = []

# Waktu default: semua data KI diasumsikan dari periode Des 2025 – Jul 2026
# (berdasarkan nama file: Data_KI_FIF_Des25-Jul26)
# Semester Ganjil 2025 = id_waktu 1 (tahun=2025, semester=Ganjil)
DEFAULT_WAKTU_ID = waktu_lookup.get((2025, "Ganjil"), 1)

for _, row in df_ki.iterrows():
    # FK — jenis KI
    id_jenis = jenis_ki_lookup.get(row["kategori"], None)

    # FK — status implementasi (dari kolom 'status' = Sertifikat/Pengajuan)
    status_map = {
        "Sertifikat Selesai": status_ki_lookup.get("Sertifikat Selesai", 1),
        "Pengajuan DJKI"    : status_ki_lookup.get("Pengajuan DJKI", 2),
    }
    id_status_impl = status_map.get(str(row["status"]), None)

    # FK — mitra (dari objek_pengaplikasian) — opsional
    objek_apl = row.get("objek_pengaplikasian")
    id_mitra_fk = mitra_lookup.get(str(objek_apl), None) if pd.notna(objek_apl) else None

    # Measure: indikator_terimplementasi (1 jika sudah aplikasi DAN komersialisasi)
    apl = row["sudah_di_aplikasikan"]
    kom = row["sudah_di_komersialisasikan"]

    if apl is True and kom is True:
        indikator = "Implementasi Penuh"
    elif apl is True:
        indikator = "Diaplikasikan (belum komersialisasi)"
    elif kom is True:
        indikator = "Dikomersialisasikan (belum diaplikasikan)"
    else:
        indikator = "Belum Terimplementasi"

    rows_fact_ki.append({
        "id_waktu_pendaftaran"     : DEFAULT_WAKTU_ID,
        "id_waktu_implementasi"    : DEFAULT_WAKTU_ID if (apl or kom) else None,
        "id_jenis_ki"              : id_jenis,
        "id_status_implementasi"   : id_status_impl,
        "id_inovasi"               : int(row["no"]),
        "id_mitra"                 : id_mitra_fk,
        "jumlah_ki"                : 1,
        "indikator_terimplementasi": indikator,
        "selang_waktu_impl_hari"   : None,  # tidak tersedia di data sumber
        "flag_aplikasi_tanpa_objek": bool(row["flag_aplikasi_tanpa_objek"]),
    })

fact_ki = pd.DataFrame(rows_fact_ki)
fact_ki.insert(0, "id_fakta_ki", range(1, len(fact_ki) + 1))

print("📋 fact_kekayaan_intelektual (5 baris)")
print(fact_ki.head(5).to_string(index=False))
print(f"\\nShape: {fact_ki.shape}")
print("\\nDistribusi indikator_terimplementasi:")
print(fact_ki["indikator_terimplementasi"].value_counts().to_string())
print(f"\\nNull id_mitra (KI belum diaplikasikan): {fact_ki['id_mitra'].isna().sum()}")
""")

code("""\
# ── Simpan fact_kekayaan_intelektual ─────────────────────────────────────
print("💾 Menyimpan fact_kekayaan_intelektual...")
save_csv(fact_ki, FACT_DIR, "fact_kekayaan_intelektual.csv")
save_excel_single(fact_ki, XL_DIR, "fact_kekayaan_intelektual.xlsx",
                  sheet_name="fact_ki",
                  header_color=THEME["fact"],
                  table_type="Accumulating Snapshot — BP4")
print("\\n✅ fact_kekayaan_intelektual selesai")
""")

# fact_status_startup
md("""\
### 2.2 `fact_status_startup` — Periodic Snapshot _(BP5)_

Grain: **1 baris = 1 anggota tim startup pada 1 periode**  
Measure: `jumlah_tim` (count anggota per startup), `is_nim_duplikat`
""")

code("""\
# ── fact_status_startup ───────────────────────────────────────────────────
# Sumber: df_ink (peserta inkubasi) + df_st (data startup master)
# Grain  : 1 baris = 1 peserta × 1 periode inkubasi
# FK     : id_waktu_periode, id_tim_startup, id_status_pendanaan, id_program_studi
# Measure: jumlah_tim (denormalized per startup), is_nim_duplikat
# ─────────────────────────────────────────────────────────────────────────

# Hitung jumlah anggota per startup
tim_size = df_ink.groupby("nama_startup").size().rename("jumlah_tim").reset_index()

# Periode inkubasi: Ganjil 2025 (berdasarkan data)
PERIODE_INKUBASI = waktu_lookup.get((2025, "Ganjil"), 1)

rows_fact_st = []
for _, row in df_ink.iterrows():
    # FK — tim startup
    id_tim = tim_lookup.get(row["nama_startup"], None)

    # FK — status pendanaan (dari data startup master)
    startup_row = df_st[df_st["nama_startup"] == row["nama_startup"]]
    if len(startup_row) > 0:
        batch_val = startup_row.iloc[0]["batch"]
        tahun_val = startup_row.iloc[0]["tahun"]
        id_status_pendanaan = resolve_status_pendanaan(batch_val, tahun_val)
    else:
        id_status_pendanaan = 1   # INKUBASI_AKTIF default

    # FK — program studi
    id_prodi = prodi_lookup.get((row["prodi"], row["universitas"]), None)

    # Measure — jumlah tim dari lookup
    j_tim = tim_size[tim_size["nama_startup"] == row["nama_startup"]]["jumlah_tim"].values
    jumlah_tim_val = int(j_tim[0]) if len(j_tim) > 0 else 1

    rows_fact_st.append({
        "id_waktu_periode"   : PERIODE_INKUBASI,
        "id_tim_startup"     : id_tim,
        "id_status_pendanaan": id_status_pendanaan,
        "id_program_studi"   : id_prodi,
        "nama_peserta"       : row["nama"],
        "nim_peserta"        : row["nim"],
        "jumlah_tim"         : jumlah_tim_val,
        "is_nim_duplikat"    : bool(row["flag_nim_duplikat"]),
    })

fact_startup = pd.DataFrame(rows_fact_st)
fact_startup.insert(0, "id_fakta_startup", range(1, len(fact_startup) + 1))

print("📋 fact_status_startup (5 baris)")
print(fact_startup.head(5).to_string(index=False))
print(f"\\nShape: {fact_startup.shape}")
print(f"\\nNIM duplikat: {fact_startup['is_nim_duplikat'].sum()} baris")
print("\\nDistribusi id_status_pendanaan:")
print(fact_startup["id_status_pendanaan"].value_counts().to_string())
""")

code("""\
# ── Simpan fact_status_startup ────────────────────────────────────────────
print("💾 Menyimpan fact_status_startup...")
save_csv(fact_startup, FACT_DIR, "fact_status_startup.csv")
save_excel_single(fact_startup, XL_DIR, "fact_status_startup.xlsx",
                  sheet_name="fact_startup",
                  header_color=THEME["fact"],
                  table_type="Periodic Snapshot — BP5")
print("\\n✅ fact_status_startup selesai")
""")

# fact_kolaborasi_industri
md("""\
### 2.3 `fact_kolaborasi_industri` — Transaction Fact _(BP1)_

Grain: **1 baris = 1 event kolaborasi** antara FIF dengan mitra industri  
Sumber: KI yang sudah diaplikasikan → setiap KI + mitra = 1 transaksi kolaborasi
""")

code("""\
# ── fact_kolaborasi_industri ──────────────────────────────────────────────
# Sumber: df_ki yang sudah_di_aplikasikan = True
# Grain  : 1 baris = 1 kolaborasi (1 KI × 1 mitra)
# FK     : id_waktu_mulai, id_mitra, id_inovasi, id_program_studi
# Measure: jumlah_kolaborasi (=1), indikator_terkomersialisasi
# ─────────────────────────────────────────────────────────────────────────

# Ambil KI yang sudah diaplikasikan & memiliki data mitra
ki_applied = df_ki[
    (df_ki["sudah_di_aplikasikan"] == True) &
    (df_ki["objek_pengaplikasian"].notna())
].copy()

rows_kolab = []
for _, row in ki_applied.iterrows():
    objek = str(row["objek_pengaplikasian"])
    id_mitra_fk = mitra_lookup.get(objek, None)

    # Prodi kontributor — tidak langsung tersedia, default None
    id_prodi_fk = None

    rows_kolab.append({
        "id_waktu_mulai"            : waktu_lookup.get((2025, "Ganjil"), 1),
        "id_mitra"                  : id_mitra_fk,
        "id_inovasi"                : int(row["no"]),
        "id_program_studi"          : id_prodi_fk,
        "jumlah_kolaborasi"         : 1,
        "indikator_terkomersialisasi": bool(row["sudah_di_komersialisasikan"])
                                       if pd.notna(row["sudah_di_komersialisasikan"])
                                       else False,
        "nama_kontributor"          : row["nama"],
        "nama_mitra_text"           : objek,
    })

fact_kolab = pd.DataFrame(rows_kolab)
fact_kolab.insert(0, "id_fakta_kolaborasi", range(1, len(fact_kolab) + 1))

print("📋 fact_kolaborasi_industri")
print(fact_kolab.to_string(index=False))
print(f"\\nShape: {fact_kolab.shape}")
print(f"\\nKolaborasi terkomersialisasi: {fact_kolab['indikator_terkomersialisasi'].sum()}")
""")

code("""\
# ── Simpan fact_kolaborasi_industri ───────────────────────────────────────
print("💾 Menyimpan fact_kolaborasi_industri...")
save_csv(fact_kolab, FACT_DIR, "fact_kolaborasi_industri.csv")
save_excel_single(fact_kolab, XL_DIR, "fact_kolaborasi_industri.xlsx",
                  sheet_name="fact_kolaborasi",
                  header_color=THEME["fact"],
                  table_type="Transaction Fact — BP1")
print("\\n✅ fact_kolaborasi_industri selesai")
""")

# fact_hibah_industri
md("""\
### 2.4 `fact_hibah_industri` — Transaction Fact _(BP2)_

Grain: **1 baris = 1 KI yang dikomersialisasikan** dengan nilai estimasi  
Sumber: KI yang sudah_di_komersialisasikan = True
""")

code("""\
# ── fact_hibah_industri ───────────────────────────────────────────────────
# Sumber: df_ki yang sudah_di_komersialisasikan = True
# Grain  : 1 baris = 1 KI yang telah menghasilkan nilai ekonomi / hibah
# FK     : id_waktu_penerimaan, id_mitra, id_inovasi
# Measure: nilai_hibah (NULL — tidak ada di data), jumlah_hibah (=1)
# Note   : nilai_hibah tidak tersedia di sumber data, diisi None/0
# ─────────────────────────────────────────────────────────────────────────

ki_komersialisasi = df_ki[df_ki["sudah_di_komersialisasikan"] == True].copy()

rows_hibah = []
for _, row in ki_komersialisasi.iterrows():
    objek_kom = row.get("objek_komersialisasi")
    id_mitra_fk = mitra_lookup.get(str(objek_kom), None) if pd.notna(objek_kom) else None

    rows_hibah.append({
        "id_waktu_penerimaan": waktu_lookup.get((2025, "Ganjil"), 1),
        "id_mitra"           : id_mitra_fk,
        "id_inovasi"         : int(row["no"]),
        "jumlah_hibah"       : 1,
        "nilai_hibah"        : None,         # tidak ada di data sumber
        "objek_komersialisasi_text": str(objek_kom) if pd.notna(objek_kom) else None,
        "nama_pemilik_ki"    : row["nama"],
        "kategori_ki"        : row["kategori"],
    })

fact_hibah = pd.DataFrame(rows_hibah)
fact_hibah.insert(0, "id_fakta_hibah", range(1, len(fact_hibah) + 1))

print("📋 fact_hibah_industri")
print(fact_hibah.head(10).to_string(index=False))
print(f"\\nShape: {fact_hibah.shape}")
print("\\nDistribusi kategori_ki pada KI yang dikomersialisasikan:")
print(fact_hibah["kategori_ki"].value_counts().to_string())
""")

code("""\
# ── Simpan fact_hibah_industri ────────────────────────────────────────────
print("💾 Menyimpan fact_hibah_industri...")
save_csv(fact_hibah, FACT_DIR, "fact_hibah_industri.csv")
save_excel_single(fact_hibah, XL_DIR, "fact_hibah_industri.xlsx",
                  sheet_name="fact_hibah",
                  header_color=THEME["fact"],
                  table_type="Transaction Fact — BP2")
print("\\n✅ fact_hibah_industri selesai")
""")

# fact_penelitian_pkm
md("""\
### 2.5 `fact_penelitian_pkm` — Transaction Fact _(BP3)_

Grain: **1 baris = 1 kegiatan penelitian/PKM** per tim/startup  
Sumber: df_st (startup dengan prodi) + dim_jenis_kegiatan (PKM-K = kewirausahaan)
""")

code("""\
# ── fact_penelitian_pkm ───────────────────────────────────────────────────
# Sumber: df_st (semua startup = representasi kegiatan PKM-K / inkubasi)
# Grain  : 1 baris = 1 startup × 1 jenis kegiatan (PKM-K atau INKUBASI)
# FK     : id_waktu, id_program_studi, id_jenis_kegiatan, id_inovasi (opsional)
# Measure: jumlah_kegiatan (=1), indikator_bermuara_startup,
#           jumlah_pelaksana_mahasiswa
# ─────────────────────────────────────────────────────────────────────────

# Jumlah mahasiswa per startup dari inkubasi
anggota_per_startup = (
    df_ink.groupby("nama_startup")
    .agg(
        jumlah_mahasiswa=("nim", "count"),
        prodi_pertama=("prodi", "first"),
        kampus_pertama=("universitas", "first"),
    )
    .reset_index()
)

# Jenis kegiatan: semua startup dikategorikan sebagai PKM-K atau INKUBASI
PKM_K_ID   = 3   # PKM Kewirausahaan
INKUBASI_ID = 9  # Program Inkubasi

rows_pkm = []
for _, row in df_st.iterrows():
    # Ambil data prodi dari inkubasi (jika ada)
    info = anggota_per_startup[anggota_per_startup["nama_startup"] == row["nama_startup"]]
    if len(info) > 0:
        prodi_val  = info.iloc[0]["prodi_pertama"]
        kampus_val = info.iloc[0]["kampus_pertama"]
        jml_mhs    = int(info.iloc[0]["jumlah_mahasiswa"])
        id_jenis   = INKUBASI_ID
        bermuara   = True   # startup ini sudah ada di program inkubasi
    else:
        prodi_val  = None
        kampus_val = None
        jml_mhs    = 0
        id_jenis   = PKM_K_ID
        bermuara   = False   # startup tidak ada di data inkubasi

    id_prodi_fk = prodi_lookup.get((prodi_val, kampus_val), None) if prodi_val else None
    id_tim_fk   = tim_lookup.get(row["nama_startup"], None)

    # Waktu: sesuai tahun startup
    semester = "Ganjil"
    id_waktu_fk = waktu_lookup.get((int(row["tahun"]), semester),
                                    waktu_lookup.get((2025, "Ganjil"), 1))

    rows_pkm.append({
        "id_waktu"                       : id_waktu_fk,
        "id_program_studi"               : id_prodi_fk,
        "id_jenis_kegiatan"              : id_jenis,
        "id_inovasi"                     : None,     # PKM tidak terikat 1 KI spesifik
        "id_tim_startup"                 : id_tim_fk,
        "jumlah_kegiatan"                : 1,
        "indikator_bermuara_startup"     : bermuara,
        "jumlah_pelaksana_mahasiswa"     : jml_mhs,
        "jumlah_pelaksana_dosen"         : None,     # tidak ada di data
        "nama_startup"                   : row["nama_startup"],
    })

fact_pkm = pd.DataFrame(rows_pkm)
fact_pkm.insert(0, "id_fakta_penelitian", range(1, len(fact_pkm) + 1))

print("📋 fact_penelitian_pkm (8 baris pertama)")
print(fact_pkm.head(8).to_string(index=False))
print(f"\\nShape: {fact_pkm.shape}")
print(f"\\nBermuara startup: {fact_pkm['indikator_bermuara_startup'].sum()} / {len(fact_pkm)}")
print(f"Rata-rata mahasiswa per kegiatan: {fact_pkm['jumlah_pelaksana_mahasiswa'].mean():.1f}")
""")

code("""\
# ── Simpan fact_penelitian_pkm ────────────────────────────────────────────
print("💾 Menyimpan fact_penelitian_pkm...")
save_csv(fact_pkm, FACT_DIR, "fact_penelitian_pkm.csv")
save_excel_single(fact_pkm, XL_DIR, "fact_penelitian_pkm.xlsx",
                  sheet_name="fact_pkm",
                  header_color=THEME["fact"],
                  table_type="Transaction Fact — BP3")
print("\\n✅ fact_penelitian_pkm selesai")
""")

md("### 2.6 Ringkasan Semua Fact Tables")

code("""\
# ── Ringkasan fact tables ─────────────────────────────────────────────────
print("=" * 65)
print(f"{'TABEL':<35} {'TIPE':<20} {'BARIS':>5} {'KOLOM':>5}")
print("=" * 65)

all_facts = [
    ("fact_kekayaan_intelektual", fact_ki,     "Accumulating — BP4"),
    ("fact_status_startup",       fact_startup, "Periodic Snap — BP5"),
    ("fact_kolaborasi_industri",  fact_kolab,   "Transaction — BP1"),
    ("fact_hibah_industri",       fact_hibah,   "Transaction — BP2"),
    ("fact_penelitian_pkm",       fact_pkm,     "Transaction — BP3"),
]

for name, df, tipe in all_facts:
    print(f"  {name:<33} {tipe:<20} {len(df):>5} {len(df.columns):>5}")

print("=" * 65)
print(f"  {'TOTAL BARIS':<33} {'':20} {sum(len(d) for _, d, _ in all_facts):>5}")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 3 — STAR SCHEMA VALIDATION
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 🔗 Phase 3 — Validasi Referential Integrity (Star Schema)

Pastikan semua FK di fact table merujuk ke PK yang valid di dimensi.
""")

code("""\
# ── Fungsi validasi FK ────────────────────────────────────────────────────
def check_fk(fact_df, fk_col, dim_df, pk_col, fact_name, dim_name,
             allow_null=True):
    \"\"\"
    Cek apakah FK di fact_df (fk_col) semuanya ada di dim_df (pk_col).
    Null FK hanya diterima jika allow_null=True.
    \"\"\"
    valid_pks = set(dim_df[pk_col].dropna().unique())
    fk_vals   = fact_df[fk_col]

    if allow_null:
        fk_non_null = fk_vals.dropna()
    else:
        null_count = fk_vals.isna().sum()
        if null_count > 0:
            print(f"  ⚠️  {fact_name}.{fk_col} → {null_count} null FK (not allowed)")
        fk_non_null = fk_vals.dropna()

    invalid = fk_non_null[~fk_non_null.astype(str).isin(
        set(str(v) for v in valid_pks)
    )]

    null_count = fk_vals.isna().sum()
    status = "✅" if len(invalid) == 0 else "❌"
    print(f"  {status} {fact_name}.{fk_col:<35} → {dim_name}.{pk_col}"
          f"  | null={null_count}, invalid={len(invalid)}")

print("=" * 70)
print("  REFERENTIAL INTEGRITY CHECK")
print("=" * 70)

print("\\n📊 fact_kekayaan_intelektual (BP4)")
check_fk(fact_ki, "id_waktu_pendaftaran",    dim_waktu, "id_waktu",
         "fact_ki", "dim_waktu", allow_null=False)
check_fk(fact_ki, "id_jenis_ki",             dim_jenis_ki, "id_jenis_ki",
         "fact_ki", "dim_jenis_ki")
check_fk(fact_ki, "id_status_implementasi",  dim_status_implementasi_ki, "id_status_implementasi",
         "fact_ki", "dim_status_ki")
check_fk(fact_ki, "id_inovasi",              dim_inovasi, "id_inovasi",
         "fact_ki", "dim_inovasi", allow_null=False)
check_fk(fact_ki, "id_mitra",               dim_mitra, "id_mitra",
         "fact_ki", "dim_mitra", allow_null=True)

print("\\n📊 fact_status_startup (BP5)")
check_fk(fact_startup, "id_waktu_periode",    dim_waktu, "id_waktu",
         "fact_startup", "dim_waktu", allow_null=False)
check_fk(fact_startup, "id_tim_startup",      dim_tim_startup, "id_tim_startup",
         "fact_startup", "dim_tim", allow_null=True)
check_fk(fact_startup, "id_status_pendanaan", dim_status_pendanaan, "id_status_pendanaan",
         "fact_startup", "dim_status_pend", allow_null=False)
check_fk(fact_startup, "id_program_studi",    dim_program_studi, "id_program_studi",
         "fact_startup", "dim_prodi", allow_null=True)

print("\\n📊 fact_kolaborasi_industri (BP1)")
check_fk(fact_kolab, "id_waktu_mulai", dim_waktu, "id_waktu",
         "fact_kolab", "dim_waktu", allow_null=False)
check_fk(fact_kolab, "id_mitra",       dim_mitra, "id_mitra",
         "fact_kolab", "dim_mitra", allow_null=True)
check_fk(fact_kolab, "id_inovasi",     dim_inovasi, "id_inovasi",
         "fact_kolab", "dim_inovasi", allow_null=False)

print("\\n📊 fact_hibah_industri (BP2)")
check_fk(fact_hibah, "id_waktu_penerimaan", dim_waktu, "id_waktu",
         "fact_hibah", "dim_waktu", allow_null=False)
check_fk(fact_hibah, "id_mitra",           dim_mitra, "id_mitra",
         "fact_hibah", "dim_mitra", allow_null=True)
check_fk(fact_hibah, "id_inovasi",         dim_inovasi, "id_inovasi",
         "fact_hibah", "dim_inovasi", allow_null=False)

print("\\n📊 fact_penelitian_pkm (BP3)")
check_fk(fact_pkm, "id_waktu",         dim_waktu, "id_waktu",
         "fact_pkm", "dim_waktu", allow_null=False)
check_fk(fact_pkm, "id_program_studi", dim_program_studi, "id_program_studi",
         "fact_pkm", "dim_prodi", allow_null=True)
check_fk(fact_pkm, "id_jenis_kegiatan",dim_jenis_kegiatan, "id_jenis_kegiatan",
         "fact_pkm", "dim_jenis_keg", allow_null=False)
check_fk(fact_pkm, "id_tim_startup",   dim_tim_startup, "id_tim_startup",
         "fact_pkm", "dim_tim", allow_null=True)

print("\\n" + "=" * 70)
print("  Validasi selesai ✅")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 4 — EXPORT MASTER EXCEL
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 📤 Phase 4 — Export Master Excel (All-in-One Workbook)

Semua tabel (9 dimensi + 5 fakta) digabung dalam 1 file Excel dengan  
sheet terpisah, styling warna per tipe tabel, dan sheet indeks.
""")

code("""\
# ── Buat workbook master ──────────────────────────────────────────────────
print("📊 Membuat star_schema_fif_inovasi.xlsx...")
wb_master = openpyxl.Workbook()

# Hapus sheet default kosong
wb_master.remove(wb_master.active)

# ── INDEX SHEET ───────────────────────────────────────────────────────────
ws_idx = wb_master.create_sheet("📋 INDEX", 0)
ws_idx.column_dimensions["A"].width = 35
ws_idx.column_dimensions["B"].width = 25
ws_idx.column_dimensions["C"].width = 10
ws_idx.column_dimensions["D"].width = 10
ws_idx.column_dimensions["E"].width = 55

idx_header = ["Nama Tabel", "Tipe", "Baris", "Kolom", "Deskripsi"]
ws_idx.append(idx_header)
for cell in ws_idx[1]:
    cell.fill = PatternFill("solid", fgColor="2C3E50")
    cell.font = Font(bold=True, color="FFFFFF", size=11)
    cell.alignment = Alignment(horizontal="center")

idx_data = [
    # Dimensi Conformed
    ("dim_waktu",                "Conformed Dim",    len(dim_waktu),
     len(dim_waktu.columns),     "Dimensi waktu akademik (semester, tahun)"),
    ("dim_inovasi",              "Conformed Dim",    len(dim_inovasi),
     len(dim_inovasi.columns),   "Dimensi inovasi / kekayaan intelektual"),
    ("dim_program_studi",        "Conformed Dim",    len(dim_program_studi),
     len(dim_program_studi.columns), "Dimensi program studi + kampus"),
    ("dim_mitra",                "Conformed Dim",    len(dim_mitra),
     len(dim_mitra.columns),     "Dimensi mitra industri / institusi"),
    # Dimensi Local
    ("dim_jenis_ki",             "Local Dim BP4",    len(dim_jenis_ki),
     len(dim_jenis_ki.columns),  "Jenis KI (Hak Cipta, Paten, dll)"),
    ("dim_status_impl_ki",       "Local Dim BP4",    len(dim_status_implementasi_ki),
     len(dim_status_implementasi_ki.columns), "Status implementasi KI"),
    ("dim_tim_startup",          "Local Dim BP5",    len(dim_tim_startup),
     len(dim_tim_startup.columns), "Identitas tim startup"),
    ("dim_status_pendanaan",     "Local Dim BP5",    len(dim_status_pendanaan),
     len(dim_status_pendanaan.columns), "Status pendanaan / batch inkubasi"),
    ("dim_jenis_kegiatan",       "Local Dim BP3",    len(dim_jenis_kegiatan),
     len(dim_jenis_kegiatan.columns), "Kategori kegiatan (PKM, Penelitian, dll)"),
    # Fact Tables
    ("fact_kekayaan_intelektual","Fact BP4",         len(fact_ki),
     len(fact_ki.columns),       "KI accumulating snapshot — status aplikasi & komersialisasi"),
    ("fact_status_startup",      "Fact BP5",         len(fact_startup),
     len(fact_startup.columns),  "Peserta inkubasi periodic snapshot"),
    ("fact_kolaborasi_industri", "Fact BP1",         len(fact_kolab),
     len(fact_kolab.columns),    "Kolaborasi KI dengan mitra industri"),
    ("fact_hibah_industri",      "Fact BP2",         len(fact_hibah),
     len(fact_hibah.columns),    "KI yang dikomersialisasikan / hibah"),
    ("fact_penelitian_pkm",      "Fact BP3",         len(fact_pkm),
     len(fact_pkm.columns),      "Kegiatan PKM / penelitian per startup"),
]

tipe_color = {
    "Conformed Dim": "FFF3E0",
    "Local Dim BP3": "E8F5E9",
    "Local Dim BP4": "E8F5E9",
    "Local Dim BP5": "E8F5E9",
    "Fact BP1": "E3F2FD",
    "Fact BP2": "E3F2FD",
    "Fact BP3": "E3F2FD",
    "Fact BP4": "E3F2FD",
    "Fact BP5": "E3F2FD",
}

for row_data in idx_data:
    ws_idx.append(list(row_data))
    tipe = row_data[1]
    color = tipe_color.get(tipe, "FFFFFF")
    for cell in ws_idx[ws_idx.max_row]:
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(vertical="center")
        cell.border = BORDER_THIN

ws_idx.freeze_panes = "A2"

print("  ✅ Sheet INDEX dibuat")
""")

code("""\
# ── Tambahkan semua sheet dimensi & fakta ─────────────────────────────────
sheets_to_add = [
    # (DataFrame, nama_sheet, warna_header)
    (dim_waktu,                  "dim_waktu",          THEME["dim_conformed"]),
    (dim_inovasi,                "dim_inovasi",         THEME["dim_conformed"]),
    (dim_program_studi,          "dim_program_studi",   THEME["dim_conformed"]),
    (dim_mitra,                  "dim_mitra",           THEME["dim_conformed"]),
    (dim_jenis_ki,               "dim_jenis_ki",        THEME["dim_local"]),
    (dim_status_implementasi_ki, "dim_status_impl_ki",  THEME["dim_local"]),
    (dim_tim_startup,            "dim_tim_startup",     THEME["dim_local"]),
    (dim_status_pendanaan,       "dim_status_pendanaan",THEME["dim_local"]),
    (dim_jenis_kegiatan,         "dim_jenis_kegiatan",  THEME["dim_local"]),
    (fact_ki,                    "fact_ki",             THEME["fact"]),
    (fact_startup,               "fact_startup",        THEME["fact"]),
    (fact_kolab,                 "fact_kolaborasi",     THEME["fact"]),
    (fact_hibah,                 "fact_hibah",          THEME["fact"]),
    (fact_pkm,                   "fact_pkm",            THEME["fact"]),
]

for df, sheet_name, color in sheets_to_add:
    add_to_master_wb(wb_master, df, sheet_name, color)
    print(f"  ✅ Sheet '{sheet_name}' ditambahkan ({len(df)} baris)")

# Simpan
master_path = XL_DIR / "star_schema_fif_inovasi.xlsx"
wb_master.save(master_path)

print(f"\\n💾 Master workbook disimpan:")
print(f"   {master_path.relative_to(ROOT)}")
print(f"   Total sheets: {len(wb_master.sheetnames)}")
print(f"   Sheets: {wb_master.sheetnames}")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 5 — SAMPLE QUERIES / ANALISIS
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## 🔍 Phase 5 — Contoh Query Analitik (JOIN Dimensi × Fakta)

Demonstrasi bahwa skema bintang dapat digunakan untuk menjawab  
pertanyaan bisnis nyata.
""")

code("""\
# ── Query 1: KI per Kategori × Status Implementasi ────────────────────────
print("=" * 60)
print("  QUERY 1: Jumlah KI per Kategori × Status Implementasi")
print("=" * 60)

q1 = (
    fact_ki
    .merge(dim_jenis_ki[["id_jenis_ki", "nama_jenis_ki"]],
           on="id_jenis_ki", how="left")
    .merge(dim_status_implementasi_ki[
               ["id_status_implementasi", "nama_status"]],
           on="id_status_implementasi", how="left")
    .groupby(["nama_jenis_ki", "indikator_terimplementasi"])
    .agg(jumlah_ki=("jumlah_ki", "sum"))
    .reset_index()
    .sort_values(["nama_jenis_ki", "jumlah_ki"], ascending=[True, False])
)

print(q1.to_string(index=False))
""")

code("""\
# ── Query 2: Startup per Bidang Usaha × Batch Pembinaan ──────────────────
print("=" * 60)
print("  QUERY 2: Startup per Bidang Usaha × Batch Pembinaan")
print("=" * 60)

q2 = (
    fact_pkm
    .merge(dim_tim_startup[["id_tim_startup", "nama_tim",
                             "bidang_usaha", "batch_pembinaan"]],
           on="id_tim_startup", how="left")
    .groupby(["bidang_usaha", "batch_pembinaan"])
    .agg(jumlah_startup=("jumlah_kegiatan", "sum"),
         total_mahasiswa=("jumlah_pelaksana_mahasiswa", "sum"))
    .reset_index()
    .sort_values("jumlah_startup", ascending=False)
)

print(q2.to_string(index=False))
""")

code("""\
# ── Query 3: Peserta Inkubasi per Prodi × Kampus ─────────────────────────
print("=" * 60)
print("  QUERY 3: Peserta Inkubasi per Program Studi × Kampus")
print("=" * 60)

q3 = (
    fact_startup
    .merge(dim_program_studi[["id_program_studi", "kode_prodi",
                               "nama_prodi", "kampus"]],
           on="id_program_studi", how="left")
    .groupby(["nama_prodi", "kampus"])
    .agg(jumlah_peserta=("id_fakta_startup", "count"))
    .reset_index()
    .sort_values("jumlah_peserta", ascending=False)
)

print(q3.to_string(index=False))
""")

code("""\
# ── Query 4: Top 5 Mitra Kolaborasi ──────────────────────────────────────
print("=" * 60)
print("  QUERY 4: Top 5 Mitra Kolaborasi")
print("=" * 60)

q4 = (
    fact_kolab
    .merge(dim_mitra[["id_mitra", "nama_mitra", "sektor_industri"]],
           on="id_mitra", how="left")
    .groupby(["nama_mitra", "sektor_industri"])
    .agg(jumlah_kolaborasi=("jumlah_kolaborasi", "sum"),
         pct_terkomersialisasi=("indikator_terkomersialisasi", "mean"))
    .reset_index()
    .sort_values("jumlah_kolaborasi", ascending=False)
    .head(5)
)

q4["pct_terkomersialisasi"] = (q4["pct_terkomersialisasi"] * 100).round(0).astype(int)
print(q4.to_string(index=False))
""")

code("""\
# ── Query 5: Ringkasan Star Schema ───────────────────────────────────────
print("=" * 60)
print("  RINGKASAN STAR SCHEMA")
print("=" * 60)

print(f"\\n  📐 Arsitektur  : Enterprise Constellation Schema")
print(f"  🔵 Fact Tables : 5 tabel")
print(f"  🟠 Dim Conformed: 4 tabel (Waktu, Inovasi, Prodi, Mitra)")
print(f"  🟢 Dim Local   : 5 tabel")
print(f"  📊 Total Dimensi: 9 tabel")

total_baris = (
    len(fact_ki) + len(fact_startup) + len(fact_kolab) +
    len(fact_hibah) + len(fact_pkm)
)
print(f"\\n  Total baris fact tables : {total_baris}")
print(f"  Total baris dim tables  : "
      f"{sum(len(d) for d in [dim_waktu, dim_inovasi, dim_program_studi, dim_mitra, dim_jenis_ki, dim_status_implementasi_ki, dim_tim_startup, dim_status_pendanaan, dim_jenis_kegiatan])}")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  PHASE 6 — FINAL REPORT
# ═══════════════════════════════════════════════════════════════════════════
md("""\
---
## ✅ Phase 6 — Laporan Akhir & Verifikasi File
""")

code("""\
# ── Verifikasi file output ────────────────────────────────────────────────
import os

print("=" * 65)
print("  VERIFIKASI FILE OUTPUT")
print("=" * 65)

print(f"\\n📁 {DWH.relative_to(ROOT)}/")

for subfolder in [DIM_DIR, FACT_DIR, XL_DIR]:
    sub_name = subfolder.name
    files = sorted(subfolder.iterdir())
    total_size = sum(f.stat().st_size for f in files)
    print(f"\\n  📂 {sub_name}/  ({len(files)} file, "
          f"total {total_size/1024:.1f} KB)")
    for f in files:
        size_kb = f.stat().st_size / 1024
        print(f"     {'📊' if f.suffix == '.xlsx' else '📄'} "
              f"{f.name:<45} {size_kb:>7.1f} KB")

print("\\n" + "=" * 65)
print("  ✅ SEMUA FILE BERHASIL DIBUAT")
print("=" * 65)
print(f"\\n  Notebook selesai dieksekusi pada:")
print(f"  {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
""")

# ═══════════════════════════════════════════════════════════════════════════
#  SAVE NOTEBOOK
# ═══════════════════════════════════════════════════════════════════════════
# Validate all code cells
errors = 0
for i, cell in enumerate(cells):
    if cell["cell_type"] == "code":
        src = "".join(cell["source"])
        try:
            compile(src, f"cell_{i}", "exec")
        except SyntaxError as e:
            print(f"❌ SyntaxError cell {i}: {e}")
            errors += 1

if errors == 0:
    print(f"✅ Semua {sum(1 for c in cells if c['cell_type']=='code')} "
          f"code cells valid (tidak ada SyntaxError)")

import pathlib as _pathlib
import json as _json

notebook = {
    "nbformat": 4,
    "nbformat_minor": 5,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "name": "python",
            "version": "3.11.0",
        },
    },
    "cells": cells,
}

out_path = _pathlib.Path(__file__).parent / "dwh_star_schema_fif.ipynb"
with open(out_path, "w", encoding="utf-8") as f:
    _json.dump(notebook, f, ensure_ascii=False, indent=1)

print(f"✅ Notebook disimpan → {out_path}")
print(f"   Total cells: {len(cells)}")
print(f"   Code cells : {sum(1 for c in cells if c['cell_type']=='code')}")
print(f"   MD cells   : {sum(1 for c in cells if c['cell_type']=='markdown')}")
