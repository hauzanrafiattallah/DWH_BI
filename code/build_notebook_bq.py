"""
Script untuk membuat notebook import_bigquery_fif_inovasi.ipynb
"""

import json
import pathlib

cells = []

def md(text: str):
    lines = text.strip().split("\n")
    source = [l + "\n" for l in lines[:-1]] + [lines[-1]]
    cells.append({
        "cell_type": "markdown",
        "metadata": {},
        "source": source
    })

def code(code_text: str):
    lines = code_text.strip().split("\n")
    source = [l + "\n" for l in lines[:-1]] + [lines[-1]]
    cells.append({
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": source
    })

# --- Header Notebook ---
md("""# Pipeline Import Star Schema ke Google BigQuery
Notebook ini memuat data warehouse Star Schema dari file Excel ke Google BigQuery menggunakan Service Account.

Spesifikasi:
- Project ID: tubes-dwh
- Dataset ID: tubes-dwh.fif_inovasi (fif_inovasi)
- Lokasi Dataset: US
- File Input: star_schema_fif_inovasi.xlsx (14 tabel: 9 dimensi, 5 fakta)
- Kredensial: Service Account JSON""")

# --- 1. Install Library ---
md("""## 1. Install library
Menginstall paket yang dibutuhkan untuk koneksi Google Cloud BigQuery, manipulasi data, dan pembacaan file Excel.""")

code("""# Install library pendukung
!pip -q install --upgrade google-cloud-bigquery pandas==2.2.3 openpyxl pyarrow db-dtypes

import pandas as pd
print("Pandas version:", pd.__version__)""")

# --- 2. Autentikasi Service Account ---
md("""## 2. Autentikasi Menggunakan Service Account
Koneksi ke BigQuery menggunakan kredensial Service Account. File JSON dapat dibaca langsung dari path lokal atau diunggah jika dijalankan pada Google Colab.""")

code("""import os
import json
import pathlib
from google.cloud import bigquery
from google.oauth2 import service_account

# Path default untuk lingkungan lokal dan Google Colab
POSSIBLE_KEY_PATHS = [
    "env/tubes-dwh-d535407fe46e.json",
    "../env/tubes-dwh-d535407fe46e.json",
    "tubes-dwh-d535407fe46e.json"
]

SERVICE_ACCOUNT_FILE = None
for p in POSSIBLE_KEY_PATHS:
    if os.path.exists(p):
        SERVICE_ACCOUNT_FILE = p
        break

# Jika dijalankan di Google Colab dan file belum tersedia, tampilkan prompt upload
if not SERVICE_ACCOUNT_FILE:
    try:
        from google.colab import files
        print("Upload file JSON Service Account:")
        uploaded_key = files.upload()
        if uploaded_key:
            SERVICE_ACCOUNT_FILE = next(iter(uploaded_key.keys()))
    except ImportError:
        pass

if not SERVICE_ACCOUNT_FILE or not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise RuntimeError("File JSON Service Account tidak ditemukan.")

# Verifikasi isi file JSON Service Account
with open(SERVICE_ACCOUNT_FILE, "r", encoding="utf-8") as f:
    sa_info = json.load(f)

if sa_info.get("type") != "service_account":
    raise ValueError("File yang dimuat bukan JSON Service Account yang valid.")

# Konfigurasi BigQuery
PROJECT_ID = "tubes-dwh"
DATASET_ID = "fif_inovasi"
LOCATION = "US"

# Inisialisasi kredensial dan BigQuery Client
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=["https://www.googleapis.com/auth/cloud-platform"]
)

client = bigquery.Client(
    project=PROJECT_ID,
    credentials=credentials,
    location=LOCATION
)

print("Autentikasi Service Account berhasil.")
print("Service Account:", sa_info.get("client_email"))
print("Project tujuan :", PROJECT_ID)
print("Dataset tujuan :", f"{PROJECT_ID}.{DATASET_ID}")
print("Location       :", LOCATION)""")

# --- 3. Upload / Baca File Excel Hasil Transformasi ---
md("""## 3. Upload Excel hasil transformasi
Membaca file Excel Star Schema yang memuat 14 tabel (9 tabel dimensi dan 5 tabel fakta).""")

code("""# Path default file Excel
POSSIBLE_EXCEL_PATHS = [
    "output/datawarehouse/excel/star_schema_fif_inovasi.xlsx",
    "../output/datawarehouse/excel/star_schema_fif_inovasi.xlsx",
    "star_schema_fif_inovasi.xlsx"
]

EXCEL_FILE = None
for p in POSSIBLE_EXCEL_PATHS:
    if os.path.exists(p):
        EXCEL_FILE = p
        break

if not EXCEL_FILE:
    try:
        from google.colab import files
        print("Upload file Excel Star Schema (star_schema_fif_inovasi.xlsx):")
        uploaded_excel = files.upload()
        if uploaded_excel:
            EXCEL_FILE = next(iter(uploaded_excel.keys()))
    except ImportError:
        pass

if not EXCEL_FILE or not os.path.exists(EXCEL_FILE):
    raise RuntimeError("File Excel Star Schema tidak ditemukan.")

print("File yang digunakan:", EXCEL_FILE)

# Daftar 14 tabel yang harus tersedia dalam workbook
expected_sheets = [
    "dim_waktu",
    "dim_inovasi",
    "dim_program_studi",
    "dim_mitra",
    "dim_jenis_ki",
    "dim_status_impl_ki",
    "dim_tim_startup",
    "dim_status_pendanaan",
    "dim_jenis_kegiatan",
    "fact_ki",
    "fact_startup",
    "fact_kolaborasi",
    "fact_hibah",
    "fact_pkm"
]

excel = pd.ExcelFile(EXCEL_FILE)
missing_sheets = set(expected_sheets) - set(excel.sheet_names)
if missing_sheets:
    raise ValueError(f"Sheet berikut tidak ditemukan: {sorted(missing_sheets)}")

dataframes = {
    sheet: pd.read_excel(EXCEL_FILE, sheet_name=sheet)
    for sheet in expected_sheets
}

print("Daftar tabel yang berhasil dimuat:")
for name, df in dataframes.items():
    print(f"- {name:<26}: {df.shape[0]:>3} baris, {df.shape[1]:>2} kolom")""")

# --- 4. Periksa Struktur Kolom dan Kualitas Data ---
md("""## 4. Periksa struktur kolom dan kualitas data
Tahap ini memastikan file yang dimuat sesuai dengan struktur Star Schema, surrogate key bersifat unik dan tidak null, serta foreign key pada fact table memiliki referensi yang valid pada tabel dimensi.""")

code("""# Definisi kolom yang diharapkan per tabel
expected_columns = {
    "dim_waktu": ["id_waktu", "bulan", "semester", "tahun", "tahun_akademik", "triwulan"],
    "dim_inovasi": ["id_inovasi", "judul_inovasi", "kategori_inovasi", "bidang_inovasi"],
    "dim_program_studi": ["id_program_studi", "kode_prodi", "nama_prodi", "jenjang", "fakultas", "kampus"],
    "dim_mitra": ["id_mitra", "nama_mitra", "sektor_industri", "skala_mitra", "wilayah"],
    "dim_jenis_ki": ["id_jenis_ki", "kode_jenis", "nama_jenis_ki", "dasar_hukum"],
    "dim_status_impl_ki": ["id_status_implementasi", "kode_status", "nama_status", "deskripsi"],
    "dim_tim_startup": ["id_tim_startup", "nama_tim", "bidang_usaha", "batch_pembinaan", "tahun_terdaftar", "prodi_mayoritas"],
    "dim_status_pendanaan": ["id_status_pendanaan", "kode_status", "nama_status", "deskripsi"],
    "dim_jenis_kegiatan": ["id_jenis_kegiatan", "kode_jenis", "nama_jenis"],
    "fact_ki": [
        "id_fakta_ki", "id_waktu_pendaftaran", "id_waktu_implementasi", "id_jenis_ki",
        "id_status_implementasi", "id_inovasi", "id_mitra", "jumlah_ki",
        "indikator_terimplementasi", "selang_waktu_impl_hari", "flag_aplikasi_tanpa_objek"
    ],
    "fact_startup": [
        "id_fakta_startup", "id_waktu_periode", "id_tim_startup", "id_status_pendanaan",
        "id_program_studi", "nama_peserta", "nim_peserta", "jumlah_tim", "is_nim_duplikat"
    ],
    "fact_kolaborasi": [
        "id_fakta_kolaborasi", "id_waktu_mulai", "id_mitra", "id_inovasi",
        "id_program_studi", "jumlah_kolaborasi", "indikator_terkomersialisasi",
        "nama_kontributor", "nama_mitra_text"
    ],
    "fact_hibah": [
        "id_fakta_hibah", "id_waktu_penerimaan", "id_mitra", "id_inovasi",
        "jumlah_hibah", "nilai_hibah", "objek_komersialisasi_text", "nama_pemilik_ki", "kategori_ki"
    ],
    "fact_pkm": [
        "id_fakta_penelitian", "id_waktu", "id_program_studi", "id_jenis_kegiatan",
        "id_inovasi", "id_tim_startup", "jumlah_kegiatan", "indikator_bermuara_startup",
        "jumlah_pelaksana_mahasiswa", "jumlah_pelaksana_dosen", "nama_startup"
    ]
}

# 1. Validasi kesesuaian kolom
for table, cols in expected_columns.items():
    actual = list(dataframes[table].columns)
    if actual != cols:
        raise ValueError(
            f"Struktur kolom {table} tidak sesuai.\\n"
            f"Seharusnya: {cols}\\n"
            f"Ditemukan : {actual}"
        )

# 2. Validasi surrogate key unik dan tidak null
surrogate_keys = [
    ("dim_waktu", "id_waktu"),
    ("dim_inovasi", "id_inovasi"),
    ("dim_program_studi", "id_program_studi"),
    ("dim_mitra", "id_mitra"),
    ("dim_jenis_ki", "id_jenis_ki"),
    ("dim_status_impl_ki", "id_status_implementasi"),
    ("dim_tim_startup", "id_tim_startup"),
    ("dim_status_pendanaan", "id_status_pendanaan"),
    ("dim_jenis_kegiatan", "id_jenis_kegiatan"),
    ("fact_ki", "id_fakta_ki"),
    ("fact_startup", "id_fakta_startup"),
    ("fact_kolaborasi", "id_fakta_kolaborasi"),
    ("fact_hibah", "id_fakta_hibah"),
    ("fact_pkm", "id_fakta_penelitian")
]

for table, key in surrogate_keys:
    df = dataframes[table]
    if df[key].isna().any() or df[key].duplicated().any():
        raise ValueError(f"{table}.{key} harus unik dan tidak boleh kosong.")

# 3. Validasi foreign key referensial integrity
fk_validations = [
    ("fact_ki", "id_waktu_pendaftaran", "dim_waktu", "id_waktu"),
    ("fact_ki", "id_jenis_ki", "dim_jenis_ki", "id_jenis_ki"),
    ("fact_ki", "id_status_implementasi", "dim_status_impl_ki", "id_status_implementasi"),
    ("fact_ki", "id_inovasi", "dim_inovasi", "id_inovasi"),
    ("fact_startup", "id_waktu_periode", "dim_waktu", "id_waktu"),
    ("fact_startup", "id_status_pendanaan", "dim_status_pendanaan", "id_status_pendanaan"),
    ("fact_kolaborasi", "id_waktu_mulai", "dim_waktu", "id_waktu"),
    ("fact_kolaborasi", "id_inovasi", "dim_inovasi", "id_inovasi"),
    ("fact_hibah", "id_waktu_penerimaan", "dim_waktu", "id_waktu"),
    ("fact_hibah", "id_inovasi", "dim_inovasi", "id_inovasi"),
    ("fact_pkm", "id_waktu", "dim_waktu", "id_waktu"),
    ("fact_pkm", "id_jenis_kegiatan", "dim_jenis_kegiatan", "id_jenis_kegiatan")
]

for f_tbl, f_col, d_tbl, d_col in fk_validations:
    valid_keys = set(dataframes[d_tbl][d_col].dropna())
    f_vals = dataframes[f_tbl][f_col].dropna()
    invalid = set(f_vals) - valid_keys
    if invalid:
        raise ValueError(f"Foreign key {f_tbl}.{f_col} tidak valid terhadap {d_tbl}.{d_col}: {invalid}")

print("Validasi struktur, surrogate key, dan foreign key berhasil.")""")

# --- 5. Buat atau Verifikasi Dataset BigQuery ---
md("""## 5. Buat atau verifikasi dataset BigQuery
Memastikan dataset target sudah tersedia pada project BigQuery dengan lokasi yang sesuai.""")

code("""dataset_ref = bigquery.Dataset(f"{PROJECT_ID}.{DATASET_ID}")
dataset_ref.location = LOCATION

try:
    client.get_dataset(dataset_ref)
    print(f"Dataset sudah tersedia: {PROJECT_ID}.{DATASET_ID}")
except Exception:
    try:
        client.create_dataset(dataset_ref, exists_ok=True)
        print(f"Dataset berhasil dibuat: {PROJECT_ID}.{DATASET_ID}")
    except Exception as exc:
        raise RuntimeError(
            f"Gagal mengakses atau membuat dataset {PROJECT_ID}.{DATASET_ID}. "
            "Periksa Project ID, BigQuery API, izin akun, dan LOCATION."
        ) from exc""")

# --- 6. Tentukan Skema BigQuery ---
md("""## 6. Tentukan skema BigQuery
Skema didefinisikan secara eksplisit untuk menjamin konsistensi tipe data di BigQuery tanpa bergantung pada inferensi otomatis.""")

code("""SCHEMAS = {
    "dim_waktu": [
        bigquery.SchemaField("id_waktu", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("bulan", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("semester", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("tahun", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("tahun_akademik", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("triwulan", "STRING", mode="NULLABLE"),
    ],
    "dim_inovasi": [
        bigquery.SchemaField("id_inovasi", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("judul_inovasi", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("kategori_inovasi", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("bidang_inovasi", "STRING", mode="NULLABLE"),
    ],
    "dim_program_studi": [
        bigquery.SchemaField("id_program_studi", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("kode_prodi", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_prodi", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("jenjang", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("fakultas", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("kampus", "STRING", mode="NULLABLE"),
    ],
    "dim_mitra": [
        bigquery.SchemaField("id_mitra", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("nama_mitra", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("sektor_industri", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("skala_mitra", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("wilayah", "STRING", mode="NULLABLE"),
    ],
    "dim_jenis_ki": [
        bigquery.SchemaField("id_jenis_ki", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("kode_jenis", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_jenis_ki", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("dasar_hukum", "STRING", mode="NULLABLE"),
    ],
    "dim_status_impl_ki": [
        bigquery.SchemaField("id_status_implementasi", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("kode_status", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_status", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("deskripsi", "STRING", mode="NULLABLE"),
    ],
    "dim_tim_startup": [
        bigquery.SchemaField("id_tim_startup", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("nama_tim", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("bidang_usaha", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("batch_pembinaan", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("tahun_terdaftar", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("prodi_mayoritas", "STRING", mode="NULLABLE"),
    ],
    "dim_status_pendanaan": [
        bigquery.SchemaField("id_status_pendanaan", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("kode_status", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_status", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("deskripsi", "STRING", mode="NULLABLE"),
    ],
    "dim_jenis_kegiatan": [
        bigquery.SchemaField("id_jenis_kegiatan", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("kode_jenis", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_jenis", "STRING", mode="NULLABLE"),
    ],
    "fact_ki": [
        bigquery.SchemaField("id_fakta_ki", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("id_waktu_pendaftaran", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_waktu_implementasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_jenis_ki", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_status_implementasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_inovasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_mitra", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_ki", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("indikator_terimplementasi", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("selang_waktu_impl_hari", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("flag_aplikasi_tanpa_objek", "BOOL", mode="NULLABLE"),
    ],
    "fact_startup": [
        bigquery.SchemaField("id_fakta_startup", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("id_waktu_periode", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_tim_startup", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_status_pendanaan", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_program_studi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("nama_peserta", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nim_peserta", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_tim", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("is_nim_duplikat", "BOOL", mode="NULLABLE"),
    ],
    "fact_kolaborasi": [
        bigquery.SchemaField("id_fakta_kolaborasi", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("id_waktu_mulai", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_mitra", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_inovasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_program_studi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_kolaborasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("indikator_terkomersialisasi", "BOOL", mode="NULLABLE"),
        bigquery.SchemaField("nama_kontributor", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_mitra_text", "STRING", mode="NULLABLE"),
    ],
    "fact_hibah": [
        bigquery.SchemaField("id_fakta_hibah", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("id_waktu_penerimaan", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_mitra", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_inovasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_hibah", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("nilai_hibah", "FLOAT64", mode="NULLABLE"),
        bigquery.SchemaField("objek_komersialisasi_text", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("nama_pemilik_ki", "STRING", mode="NULLABLE"),
        bigquery.SchemaField("kategori_ki", "STRING", mode="NULLABLE"),
    ],
    "fact_pkm": [
        bigquery.SchemaField("id_fakta_penelitian", "INT64", mode="REQUIRED"),
        bigquery.SchemaField("id_waktu", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_program_studi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_jenis_kegiatan", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_inovasi", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("id_tim_startup", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_kegiatan", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("indikator_bermuara_startup", "BOOL", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_pelaksana_mahasiswa", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("jumlah_pelaksana_dosen", "INT64", mode="NULLABLE"),
        bigquery.SchemaField("nama_startup", "STRING", mode="NULLABLE"),
    ]
}

print(f"Skema eksplisit telah didefinisikan untuk {len(SCHEMAS)} tabel.")""")

# --- 7. Upload Tabel ke BigQuery ---
md("""## 7. Upload tabel ke BigQuery
Memuat tabel dimensi terlebih dahulu, kemudian tabel fakta. Operasi menggunakan WRITE_TRUNCATE untuk mengganti isi tabel jika sudah ada.""")

code("""LOAD_ORDER = [
    # Dimensi
    "dim_waktu",
    "dim_inovasi",
    "dim_program_studi",
    "dim_mitra",
    "dim_jenis_ki",
    "dim_status_impl_ki",
    "dim_tim_startup",
    "dim_status_pendanaan",
    "dim_jenis_kegiatan",
    # Fakta
    "fact_ki",
    "fact_startup",
    "fact_kolaborasi",
    "fact_hibah",
    "fact_pkm"
]

load_results = []

for table_name in LOAD_ORDER:
    table_id = f"{PROJECT_ID}.{DATASET_ID}.{table_name}"
    df = dataframes[table_name].copy()
    
    # Normalisasi tipe data agar selaras dengan skema BigQuery
    schema_fields = {f.name: f.field_type for f in SCHEMAS[table_name]}
    for col in df.columns:
        ftype = schema_fields.get(col)
        if ftype == "INT64":
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
        elif ftype == "STRING":
            df[col] = df[col].apply(
                lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else (str(x) if pd.notna(x) else None)
            )
        elif ftype == "BOOL":
            df[col] = df[col].astype("boolean")
        elif ftype == "FLOAT64":
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("float64")
            
    job_config = bigquery.LoadJobConfig(
        write_disposition="WRITE_TRUNCATE",
        schema=SCHEMAS[table_name]
    )
    
    load_job = client.load_table_from_dataframe(df, table_id, job_config=job_config)
    load_job.result()
    
    destination_table = client.get_table(table_id)
    load_results.append({
        "nama_tabel": table_name,
        "baris_dimuat": destination_table.num_rows,
        "status": "Sukses"
    })
    print(f"Berhasil memuat {table_name:<26}: {destination_table.num_rows:>3} baris")

summary_df = pd.DataFrame(load_results)
print("\\nProses pemuatan seluruh tabel selesai.")""")

# --- 8. Verifikasi Tabel di BigQuery ---
md("""## 8. Verifikasi tabel di BigQuery
Memeriksa jumlah baris yang tersimpan pada masing-masing tabel langsung melalui query BigQuery.""")

code("""union_queries = [
    f"SELECT '{tbl}' AS nama_tabel, COUNT(*) AS jumlah_record FROM `{PROJECT_ID}.{DATASET_ID}.{tbl}`"
    for tbl in LOAD_ORDER
]
verification_sql = "\\nUNION ALL\\n".join(union_queries) + "\\nORDER BY nama_tabel"

verification_df = client.query(verification_sql).to_dataframe(create_bqstorage_client=False)
display(verification_df)""")

# --- 9. Contoh Query Analitik JOIN Star Schema ---
md("""## 9. Contoh query JOIN Star Schema
Pengujian query analitik yang menggabungkan tabel fakta dengan tabel dimensi yang bersesuaian.""")

code("""# Query 1: Analisis Kekayaan Intelektual berdasarkan Kategori dan Status Implementasi
query_ki = f\"\"\"
SELECT
    j.nama_jenis_ki AS jenis_ki,
    s.nama_status AS status_djki,
    f.indikator_terimplementasi,
    COUNT(f.id_fakta_ki) AS jumlah_ki
FROM `{PROJECT_ID}.{DATASET_ID}.fact_ki` AS f
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_jenis_ki` AS j
    ON f.id_jenis_ki = j.id_jenis_ki
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_status_impl_ki` AS s
    ON f.id_status_implementasi = s.id_status_implementasi
GROUP BY jenis_ki, status_djki, f.indikator_terimplementasi
ORDER BY jenis_ki, jumlah_ki DESC
\"\"\"

print("Hasil Analisis Kekayaan Intelektual:")
df_res_ki = client.query(query_ki).to_dataframe(create_bqstorage_client=False)
display(df_res_ki)""")

code("""# Query 2: Analisis Startup Inkubasi berdasarkan Program Studi dan Kampus
query_startup = f\"\"\"
SELECT
    p.nama_prodi,
    p.kampus,
    t.batch_pembinaan,
    COUNT(f.id_fakta_startup) AS jumlah_peserta,
    COUNT(DISTINCT f.id_tim_startup) AS jumlah_tim_unik
FROM `{PROJECT_ID}.{DATASET_ID}.fact_startup` AS f
LEFT JOIN `{PROJECT_ID}.{DATASET_ID}.dim_tim_startup` AS t
    ON f.id_tim_startup = t.id_tim_startup
LEFT JOIN `{PROJECT_ID}.{DATASET_ID}.dim_program_studi` AS p
    ON f.id_program_studi = p.id_program_studi
GROUP BY p.nama_prodi, p.kampus, t.batch_pembinaan
ORDER BY jumlah_peserta DESC
\"\"\"

print("Hasil Analisis Startup Inkubasi:")
df_res_startup = client.query(query_startup).to_dataframe(create_bqstorage_client=False)
display(df_res_startup)""")

code("""# Query 3: Analisis Kolaborasi Industri dan Tingkat Komersialisasi
query_kolaborasi = f\"\"\"
SELECT
    m.sektor_industri,
    COUNT(k.id_fakta_kolaborasi) AS total_kolaborasi,
    COUNTIF(k.indikator_terkomersialisasi = TRUE) AS kolaborasi_terkomersialisasi,
    ROUND(COUNTIF(k.indikator_terkomersialisasi = TRUE) / COUNT(k.id_fakta_kolaborasi) * 100, 2) AS persentase_komersialisasi
FROM `{PROJECT_ID}.{DATASET_ID}.fact_kolaborasi` AS k
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_mitra` AS m
    ON k.id_mitra = m.id_mitra
GROUP BY m.sektor_industri
ORDER BY total_kolaborasi DESC
\"\"\"

print("Hasil Analisis Kolaborasi Industri:")
df_res_kolaborasi = client.query(query_kolaborasi).to_dataframe(create_bqstorage_client=False)
display(df_res_kolaborasi)""")

# --- 10. Pengukuran dan Analisis 5 KPI Utama ---
md("""## 10. Pengukuran dan Analisis 5 KPI Utama
Bagian ini memuat implementasi query SQL BigQuery untuk menjawab 5 Key Performance Indicators (KPI) utama sesuai rancangan proses bisnis (PB1 - PB5) Enterprise Constellation Schema DWH Inovasi FIF:
- K1 (PB1): Kolaborasi Industri (per mitra, per inovasi komersial)
- K2 (PB2): Jumlah Pendanaan Hibah Industri (per mitra)
- K3 (PB3): Jumlah Penelitian dan PkM yang Memiliki Startup
- K4 (PB4): Jumlah Kekayaan Intelektual (per jenis KI, per status implementasi)
- K5 (PB5): Jumlah Tim Startup (per status pendanaan)""")

# KPI 1
md("""### 10.1 KPI 1 (PB1) - Kolaborasi Industri
Tujuan: Menghitung total kolaborasi industri per mitra dan rasio inovasi yang berhasil dikomersialisasikan.
Tabel fakta: fact_kolaborasi (PB1 - Transaction Fact)
Ukuran (measures): jumlah_kolaborasi, jumlah inovasi dikomersialisasikan""")

code("""query_kpi1 = f\"\"\"
SELECT
    COALESCE(m.nama_mitra, 'Mitra Tidak Teridentifikasi') AS nama_mitra,
    m.sektor_industri,
    COUNT(k.id_fakta_kolaborasi) AS total_kolaborasi,
    COUNT(DISTINCT k.id_inovasi) AS total_inovasi_dilibatkan,
    COUNTIF(k.indikator_terkomersialisasi = TRUE) AS jumlah_inovasi_dikomersialisasikan,
    ROUND(COUNTIF(k.indikator_terkomersialisasi = TRUE) / COUNT(k.id_fakta_kolaborasi) * 100, 2) AS persentase_komersialisasi
FROM `{PROJECT_ID}.{DATASET_ID}.fact_kolaborasi` AS k
LEFT JOIN `{PROJECT_ID}.{DATASET_ID}.dim_mitra` AS m
    ON k.id_mitra = m.id_mitra
GROUP BY nama_mitra, m.sektor_industri
ORDER BY total_kolaborasi DESC, jumlah_inovasi_dikomersialisasikan DESC
LIMIT 15;
\"\"\"

print("KPI 1 - Kolaborasi Industri per Mitra:")
df_kpi1 = client.query(query_kpi1).to_dataframe(create_bqstorage_client=False)
display(df_kpi1)""")

# KPI 2
md("""### 10.2 KPI 2 (PB2) - Jumlah Pendanaan Hibah Industri
Tujuan: Menghitung total pendanaan hibah dan frekuensi penerimaan hibah industri per mitra.
Tabel fakta: fact_hibah (PB2 - Transaction Fact)
Ukuran (measures): total pendanaan hibah, frekuensi penerimaan hibah
Catatan: Nominal pendanaan saat ini berstatus belum tersedia (menunggu persetujuan data CoE MREC), namun catatan transaksi penerimaan hibah per mitra/objek telah terekam.""")

code("""query_kpi2 = f\"\"\"
SELECT
    COALESCE(m.nama_mitra, h.objek_komersialisasi_text, 'Mitra Belum Teridentifikasi') AS nama_mitra,
    m.sektor_industri,
    COUNT(h.id_fakta_hibah) AS frekuensi_penerimaan_hibah,
    COUNT(DISTINCT h.id_inovasi) AS jumlah_inovasi_komersil,
    COALESCE(SUM(h.nilai_hibah), 0) AS total_nominal_hibah,
    CASE 
        WHEN SUM(h.nilai_hibah) IS NULL THEN 'Data Nominal Belum Tersedia (Menunggu CoE MREC)'
        ELSE 'Tersedia'
    END AS status_data_pendanaan
FROM `{PROJECT_ID}.{DATASET_ID}.fact_hibah` AS h
LEFT JOIN `{PROJECT_ID}.{DATASET_ID}.dim_mitra` AS m
    ON h.id_mitra = m.id_mitra
GROUP BY nama_mitra, m.sektor_industri, h.objek_komersialisasi_text
ORDER BY frekuensi_penerimaan_hibah DESC
LIMIT 15;
\"\"\"

print("KPI 2 - Pendanaan Hibah Industri per Mitra:")
df_kpi2 = client.query(query_kpi2).to_dataframe(create_bqstorage_client=False)
display(df_kpi2)""")

# KPI 3
md("""### 10.3 KPI 3 (PB3) - Jumlah Penelitian dan PkM yang Memiliki Startup
Tujuan: Menghitung jumlah kegiatan penelitian dan pengabdian masyarakat (PkM) yang bermuara menjadi startup mahasiswa.
Tabel fakta: fact_pkm (PB3 - Transaction Fact)
Ukuran (measures): jumlah penelitian dan PkM bermuara startup, rasio keberhasilan hilirisasi startup""")

code("""query_kpi3 = f\"\"\"
SELECT
    jk.nama_jenis AS jenis_kegiatan,
    COUNT(p.id_fakta_penelitian) AS total_kegiatan,
    COUNTIF(p.indikator_bermuara_startup = TRUE) AS jumlah_bermuara_startup,
    COUNTIF(p.indikator_bermuara_startup = FALSE) AS jumlah_belum_bermuara_startup,
    ROUND(COUNTIF(p.indikator_bermuara_startup = TRUE) / COUNT(p.id_fakta_penelitian) * 100, 2) AS rasio_bermuara_startup_persen,
    SUM(p.jumlah_pelaksana_mahasiswa) AS total_mahasiswa_terlibat
FROM `{PROJECT_ID}.{DATASET_ID}.fact_pkm` AS p
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_jenis_kegiatan` AS jk
    ON p.id_jenis_kegiatan = jk.id_jenis_kegiatan
GROUP BY jk.nama_jenis
ORDER BY total_kegiatan DESC;
\"\"\"

print("KPI 3 - Penelitian dan PkM yang Memiliki Startup:")
df_kpi3 = client.query(query_kpi3).to_dataframe(create_bqstorage_client=False)
display(df_kpi3)""")

# KPI 4
md("""### 10.4 KPI 4 (PB4) - Jumlah Kekayaan Intelektual
Tujuan: Menghitung jumlah pendaftaran KI menurut jenis KI, status penetapan DJKI, dan status implementasi di industri/masyarakat.
Tabel fakta: fact_ki (PB4 - Accumulating Snapshot Fact)
Ukuran (measures): jumlah KI, jumlah KI terimplementasi, persentase implementasi""")

code("""query_kpi4 = f\"\"\"
SELECT
    j.nama_jenis_ki AS jenis_ki,
    s.nama_status AS status_djki,
    COUNT(f.id_fakta_ki) AS total_ki,
    COUNTIF(f.indikator_terimplementasi LIKE 'Implementasi Penuh%' OR f.indikator_terimplementasi LIKE 'Diaplikasikan%') AS jumlah_ki_terimplementasi,
    COUNTIF(f.indikator_terimplementasi LIKE 'Implementasi Penuh%' OR f.indikator_terimplementasi LIKE 'Dikomersialisasikan%') AS jumlah_ki_dikomersialisasikan,
    COUNTIF(f.indikator_terimplementasi = 'Belum Terimplementasi') AS jumlah_ki_belum_terimplementasi,
    ROUND(COUNTIF(f.indikator_terimplementasi != 'Belum Terimplementasi') / COUNT(f.id_fakta_ki) * 100, 2) AS persentase_implementasi
FROM `{PROJECT_ID}.{DATASET_ID}.fact_ki` AS f
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_jenis_ki` AS j
    ON f.id_jenis_ki = j.id_jenis_ki
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_status_impl_ki` AS s
    ON f.id_status_implementasi = s.id_status_implementasi
GROUP BY jenis_ki, status_djki
ORDER BY total_ki DESC;
\"\"\"

print("KPI 4 - Kekayaan Intelektual per Jenis dan Status Implementasi:")
df_kpi4 = client.query(query_kpi4).to_dataframe(create_bqstorage_client=False)
display(df_kpi4)""")

# KPI 5
md("""### 10.5 KPI 5 (PB5) - Jumlah Tim Startup per Status Pendanaan
Tujuan: Menghitung sebaran jumlah tim startup dan keterlibatan mahasiswa berdasarkan status pendanaan dan batch pembinaan inkubasi.
Tabel fakta: fact_startup (PB5 - Periodic Snapshot Fact)
Ukuran (measures): jumlah tim startup, jumlah tim didanai/tidak didanai, rata-rata anggota per tim""")

code("""query_kpi5 = f\"\"\"
SELECT
    sp.nama_status AS status_pendanaan_inkubasi,
    sp.kode_status,
    COUNT(DISTINCT f.id_tim_startup) AS jumlah_tim_startup,
    COUNT(f.id_fakta_startup) AS jumlah_anggota_mahasiswa,
    ROUND(COUNT(f.id_fakta_startup) / NULLIF(COUNT(DISTINCT f.id_tim_startup), 0), 1) AS rata_rata_anggota_per_tim
FROM `{PROJECT_ID}.{DATASET_ID}.fact_startup` AS f
JOIN `{PROJECT_ID}.{DATASET_ID}.dim_status_pendanaan` AS sp
    ON f.id_status_pendanaan = sp.id_status_pendanaan
GROUP BY status_pendanaan_inkubasi, sp.kode_status
ORDER BY jumlah_tim_startup DESC;
\"\"\"

print("KPI 5 - Jumlah Tim Startup per Status Pendanaan:")
df_kpi5 = client.query(query_kpi5).to_dataframe(create_bqstorage_client=False)
display(df_kpi5)""")

# --- 11. Catatan Penggunaan, Izin, dan Keamanan ---
md("""## 11. Catatan penggunaan, izin, dan keamanan
- Project ID: Gunakan Project ID 'tubes-dwh' (bukan display name).
- Identitas eksekusi: Operasi BigQuery dijalankan menggunakan Service Account 'colab-bigquery-loader@tubes-dwh.iam.gserviceaccount.com'.
- Izin minimum yang disarankan: Peran BigQuery Job User pada level project dan BigQuery Data Editor pada dataset 'fif_inovasi'.
- Lokasi dataset: LOCATION disetel ke 'US' sesuai konfigurasi dataset pada Google Cloud.
- Penimpaan data: WRITE_TRUNCATE mengganti seluruh isi tabel target saat dimuat ulang.
- Surrogate key: Surrogate key yang sudah dibentuk pada tahap transformasi dimuat langsung ke BigQuery tanpa auto-increment ulang.
- Keamanan kunci privat: File JSON kredensial tidak boleh diunggah ke repositori publik (tambahkan folder env/ dan *.json ke .gitignore).""")

# Validate notebook syntax
for i, cell in enumerate(cells):
    if cell["cell_type"] == "code":
        code_str = "".join(cell["source"])
        # Remove shell lines like !pip for syntax check
        clean_code = "\n".join([line for line in code_str.split("\n") if not line.strip().startswith("!")])
        try:
            compile(clean_code, f"cell_{i}", "exec")
        except SyntaxError as e:
            print(f"Syntax error in cell {i}: {e}")

notebook = {
    "nbformat": 4,
    "nbformat_minor": 5,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3"
        },
        "language_info": {
            "name": "python",
            "version": "3.11.0"
        }
    },
    "cells": cells
}

notebook_path = pathlib.Path(__file__).resolve().parent / "import_bigquery_fif_inovasi.ipynb"
with open(notebook_path, "w", encoding="utf-8") as f:
    json.dump(notebook, f, ensure_ascii=False, indent=1)

print(f"Notebook berhasil disimpan di {notebook_path}")
print(f"Total cell: {len(cells)}")
