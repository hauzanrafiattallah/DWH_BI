"""ETL Pipeline: Data Inovasi & Kewirausahaan FIF
=================================================
Sumber data: 3 file Excel dari folder dataset/
  1. Data KI FIF (Kekayaan Intelektual)
  2. Data Startup FIF 2025-2026
  3. Peserta Inkubasi FIF 2025

Output: File CSV bersih di folder output/ + dokumen ringkasan di folder docs/

Jalankan:
  python code/etl_fif_inovasi.py
"""

import re
import pathlib
import datetime

import pandas as pd

# ── Path setup ───────────────────────────────────────────────────────────────
ROOT = pathlib.Path(__file__).resolve().parent.parent
DATASET = ROOT / "dataset"
OUT = ROOT / "output"
DOCS = ROOT / "docs"

OUT.mkdir(exist_ok=True)
DOCS.mkdir(exist_ok=True)

TIMESTAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# ── Mapping normalisasi ─────────────────────────────────────────────────────
PRODI_MAP = {
    "S1 Informatika - Informatika": "S1 Informatika",
    "S1 INFORMATIKA": "S1 Informatika",
    "s1 informatika": "S1 Informatika",
}

UNIV_MAP = {
    "Telkom University": "Telkom University Bandung",
}

BOOL_MAP = {"Y": True, "N": False, "Ya": True, "Tidak": False}


# ══════════════════════════════════════════════════════════════════════════════
#  PIPELINE 1: Data KI FIF (Kekayaan Intelektual)
# ══════════════════════════════════════════════════════════════════════════════
def pipeline_ki():
    """Membersihkan data Kekayaan Intelektual FIF."""
    src = DATASET / "1_Data_KI_FIF_Des25-Jul26.xlsx"
    df = pd.read_excel(src)
    log = {"file": src.name, "baris_awal": len(df), "langkah": []}

    # 1. Strip whitespace di semua kolom string
    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None, "None": None})
    log["langkah"].append("Strip whitespace pada semua kolom string")

    # 2. Ganti #NAME? error → None
    name_error_mask = df.apply(lambda row: row.astype(str).str.contains("#NAME\\?", regex=True).any(), axis=1)
    name_error_count = name_error_mask.sum()
    for col in df.columns:
        df[col] = df[col].apply(lambda x: None if isinstance(x, str) and "#NAME?" in x else x)
    log["langkah"].append(f"Hapus #NAME? error → None ({name_error_count} baris terpengaruh)")

    # 3. Normalisasi kolom boolean: sudah_di_aplikasikan & sudah_di_komersialisasikan
    for col in ["sudah_di_aplikasikan", "sudah_di_komersialisasikan"]:
        df[col] = df[col].map(BOOL_MAP)
    log["langkah"].append("Konversi Y/N → True/False pada kolom aplikasi & komersialisasi")

    # 4. Perbaiki data leakage di objek_komersialisasi:
    #    Kolom ini seharusnya berisi deskripsi objek, bukan "Ya"/"Tidak"
    #    Jika isi = "Ya" atau "Tidak", pindahkan ke sudah_di_komersialisasikan jika kosong
    leak_ya = df["objek_komersialisasi"] == "Ya"
    leak_tidak = df["objek_komersialisasi"] == "Tidak"
    # Set sudah_di_komersialisasikan dari leak jika NaN
    df.loc[leak_ya & df["sudah_di_komersialisasikan"].isna(), "sudah_di_komersialisasikan"] = True
    df.loc[leak_tidak & df["sudah_di_komersialisasikan"].isna(), "sudah_di_komersialisasikan"] = False
    # Bersihkan objek_komersialisasi dari Ya/Tidak
    df.loc[leak_ya | leak_tidak, "objek_komersialisasi"] = None
    log["langkah"].append(
        f"Fix data leakage: 'Ya' ({leak_ya.sum()}) & 'Tidak' ({leak_tidak.sum()}) "
        "di objek_komersialisasi → dipindah ke sudah_di_komersialisasikan"
    )

    # 5. Normalisasi casing inkonsisten di objek_komersialisasi
    case_fixes = {
        "Perguruan tinggi lainnya": "Perguruan Tinggi Lainnya",
        "Perguruan Tinggi lainnya": "Perguruan Tinggi Lainnya",
        "Universitas lain": "Universitas Lain",
        "universitas lain": "Universitas Lain",
        "sekolah": "Sekolah",
    }
    df["objek_komersialisasi"] = df["objek_komersialisasi"].replace(case_fixes)
    log["langkah"].append("Normalisasi casing inkonsisten di objek_komersialisasi")

    # 6. Normalisasi casing di objek_pengaplikasian
    case_fixes_apl = {
        "Universitas telkom": "Universitas Telkom",
    }
    df["objek_pengaplikasian"] = df["objek_pengaplikasian"].replace(case_fixes_apl)
    log["langkah"].append("Normalisasi casing di objek_pengaplikasian")

    # 7. Tambah flag validasi: applied=True tapi objek kosong
    df["flag_aplikasi_tanpa_objek"] = df["sudah_di_aplikasikan"].eq(True) & df["objek_pengaplikasian"].isna()
    count_flag = df["flag_aplikasi_tanpa_objek"].sum()
    log["langkah"].append(f"Tambah flag: sudah_di_aplikasikan=True tapi objek kosong ({count_flag} baris)")

    # 8. Pastikan tipe data 'no' adalah integer
    df["no"] = df["no"].astype(int)
    log["langkah"].append("Pastikan kolom 'no' bertipe integer")

    # 9. Newlines di objek_komersialisasi → spasi (untuk CSV)
    for col in ["objek_komersialisasi", "objek_pengaplikasian"]:
        df[col] = df[col].apply(lambda x: re.sub(r"\s*\n\s*", " | ", x) if isinstance(x, str) else x)
    log["langkah"].append("Normalisasi newline di kolom objek → separator ' | '")

    log["baris_akhir"] = len(df)
    log["null_summary"] = df.isnull().sum().to_dict()
    log["kategori_dist"] = df["kategori"].value_counts().to_dict()
    log["status_dist"] = df["status"].value_counts().to_dict()

    out_path = OUT / "clean_data_ki_fif.csv"
    df.to_csv(out_path, index=False)
    print(f"✅ KI: {len(df)} baris → {out_path.name}")

    return df, log


# ══════════════════════════════════════════════════════════════════════════════
#  PIPELINE 2: Data Startup FIF 2025-2026
# ══════════════════════════════════════════════════════════════════════════════
def pipeline_startup():
    """Membersihkan data Startup FIF."""
    src = DATASET / "2_Data_Startup_FIF_2025-2026.xlsx"

    # File punya header gabungan di baris pertama; baris ke-2 adalah header asli
    df_raw = pd.read_excel(src, header=None)
    log = {"file": src.name, "baris_awal_raw": len(df_raw), "langkah": []}

    # Temukan baris header
    header_idx = None
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "no" in vals and "deskripsi" in vals:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header row tidak ditemukan di file startup!")

    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_idx].values]
    df = df.reset_index(drop=True)
    log["langkah"].append(f"Skip header gabungan, header asli ditemukan di baris {header_idx}")

    # 1. Strip whitespace
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None, "None": None})
    log["langkah"].append("Strip whitespace pada semua kolom string")

    # 2. Rename kolom agar konsisten (snake_case)
    df.columns = ["no", "batch_tahun", "nama_startup", "deskripsi"]
    log["langkah"].append("Rename kolom ke snake_case: no, batch_tahun, nama_startup, deskripsi")

    # 3. Konversi 'no' ke integer
    df["no"] = pd.to_numeric(df["no"], errors="coerce").astype("Int64")
    log["langkah"].append("Konversi kolom 'no' ke integer")

    # 4. Normalisasi batch_tahun → pisah batch & tahun
    def parse_batch_tahun(val):
        if val is None:
            return None, None
        val = str(val).strip()
        # Pola: "Batch 22 Tahun 2025", "Tahun 2025", "Batch 23 Tahun 2026"
        m_batch = re.search(r"Batch\s+(\d+)", val)
        m_tahun = re.search(r"(?:Tahun\s+)?(\d{4})", val)
        batch = int(m_batch.group(1)) if m_batch else None
        tahun = int(m_tahun.group(1)) if m_tahun else None
        return batch, tahun

    parsed = df["batch_tahun"].apply(parse_batch_tahun)
    df["batch"] = parsed.apply(lambda x: x[0]).astype("Int64")
    df["tahun"] = parsed.apply(lambda x: x[1]).astype("Int64")
    log["langkah"].append("Pisah batch_tahun → kolom 'batch' (Int) & 'tahun' (Int)")

    # 5. Normalisasi nama startup (title case, strip extra space)
    df["nama_startup"] = df["nama_startup"].apply(
        lambda x: re.sub(r"\s+", " ", x).strip() if isinstance(x, str) else x
    )
    log["langkah"].append("Normalisasi spasi ganda di nama_startup")

    # 6. Bersihkan deskripsi dari whitespace berlebih
    df["deskripsi"] = df["deskripsi"].apply(
        lambda x: re.sub(r"\s+", " ", x).strip() if isinstance(x, str) else x
    )
    log["langkah"].append("Bersihkan whitespace berlebih di deskripsi")

    # 7. Hapus baris yang sepenuhnya kosong
    null_rows = df[["nama_startup", "deskripsi"]].isnull().all(axis=1)
    df = df[~null_rows].reset_index(drop=True)
    log["langkah"].append(f"Hapus baris kosong ({null_rows.sum()} baris)")

    # 8. Susun kolom final
    df = df[["no", "nama_startup", "batch", "tahun", "deskripsi"]]

    log["baris_akhir"] = len(df)
    log["tahun_dist"] = df["tahun"].value_counts(dropna=False).to_dict()
    log["batch_dist"] = df["batch"].value_counts(dropna=False).to_dict()

    out_path = OUT / "clean_data_startup_fif.csv"
    df.to_csv(out_path, index=False)
    print(f"✅ Startup: {len(df)} baris → {out_path.name}")

    return df, log


# ══════════════════════════════════════════════════════════════════════════════
#  PIPELINE 3: Peserta Inkubasi FIF 2025
# ══════════════════════════════════════════════════════════════════════════════
def pipeline_inkubasi():
    """Membersihkan data Peserta Inkubasi FIF."""
    src = DATASET / "3_Peserta_Inkubasi_FIF_2025.xlsx"
    df_raw = pd.read_excel(src, header=None)
    log = {"file": src.name, "baris_awal_raw": len(df_raw), "langkah": []}

    # Temukan baris header (cari baris yang mengandung 'No' dan 'NIM')
    header_idx = None
    for i, row in df_raw.iterrows():
        vals = [str(v).strip() for v in row.values]
        if "No" in vals and "NIM" in vals:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Header row tidak ditemukan di file inkubasi!")

    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_idx].values]
    df = df.reset_index(drop=True)
    log["langkah"].append(f"Skip header gabungan, header asli ditemukan di baris {header_idx}")

    # 1. Rename kolom ke snake_case
    df.columns = ["no", "nama_startup", "nama", "nim", "prodi", "fakultas", "universitas"]
    log["langkah"].append("Rename kolom ke snake_case")

    # 2. Strip whitespace
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip().replace({"nan": None, "": None, "None": None})
    log["langkah"].append("Strip whitespace pada semua kolom string")

    # 3. Handle nomor urut yang hilang (merged cell) → forward fill berdasarkan startup
    #    Tapi 'no' NaN artinya baris masih bagian dari startup sebelumnya
    no_null_count = df["no"].isna().sum()
    log["langkah"].append(f"Ditemukan {no_null_count} baris dengan 'no' kosong (merged cell)")

    # 4. Forward fill nama_startup untuk merged cells
    df["nama_startup"] = df["nama_startup"].ffill()
    log["langkah"].append("Forward fill nama_startup untuk merged cells")

    # 5. Konversi NIM ke string (tanpa decimal), handle NaN
    def clean_nim(val):
        if val is None or str(val).strip().lower() in ("nan", "none", ""):
            return None
        val = str(val).strip()
        # Hapus .0 jika ada (dari float conversion)
        val = re.sub(r"\.0$", "", val)
        return val

    df["nim"] = df["nim"].apply(clean_nim)
    nim_null = df["nim"].isna().sum()
    log["langkah"].append(f"Bersihkan NIM (hapus .0, handle NaN) → {nim_null} NIM kosong")

    # 6. Normalisasi prodi
    df["prodi"] = df["prodi"].replace(PRODI_MAP)
    prodi_fixes = df["prodi"].value_counts().to_dict()
    log["langkah"].append(f"Normalisasi prodi (hapus inkonsistensi casing). Distribusi: {prodi_fixes}")

    # 7. Normalisasi universitas
    df["universitas"] = df["universitas"].replace(UNIV_MAP)
    log["langkah"].append("Normalisasi nama universitas (Telkom University → Telkom University Bandung)")

    # 8. Deteksi NIM duplikat (anggota yang masuk di lebih dari 1 startup)
    nim_valid = df[df["nim"].notna()]
    dup_nim = nim_valid[nim_valid.duplicated("nim", keep=False)]
    df["flag_nim_duplikat"] = df["nim"].notna() & df["nim"].isin(dup_nim["nim"].unique())
    dup_count = df["flag_nim_duplikat"].sum()
    log["langkah"].append(f"Deteksi NIM duplikat (anggota di >1 startup): {dup_count} baris")
    if dup_count > 0:
        log["nim_duplikat_detail"] = (
            dup_nim.groupby("nim")["nama_startup"]
            .apply(list)
            .to_dict()
        )

    # 9. Nomor urut: isi ulang secara berurutan
    df["no"] = range(1, len(df) + 1)
    log["langkah"].append("Re-number kolom 'no' secara berurutan (1..N)")

    # 10. Hapus baris yang benar-benar kosong (tidak ada nama & nim)
    empty_mask = df["nama"].isna() & df["nim"].isna()
    df = df[~empty_mask].reset_index(drop=True)
    df["no"] = range(1, len(df) + 1)
    log["langkah"].append(f"Hapus baris kosong ({empty_mask.sum()} baris)")

    log["baris_akhir"] = len(df)
    log["null_summary"] = df.isnull().sum().to_dict()
    log["startup_count"] = df["nama_startup"].nunique()
    log["prodi_dist"] = df["prodi"].value_counts().to_dict()
    log["univ_dist"] = df["universitas"].value_counts().to_dict()

    out_path = OUT / "clean_peserta_inkubasi_fif.csv"
    df.to_csv(out_path, index=False)
    print(f"✅ Inkubasi: {len(df)} baris → {out_path.name}")

    return df, log


# ══════════════════════════════════════════════════════════════════════════════
#  Buat Dokumen Ringkasan (.md)
# ══════════════════════════════════════════════════════════════════════════════
def generate_docs(logs: dict, dfs: dict):
    """Generate markdown documentation dengan ringkasan proses ETL."""
    doc_path = DOCS / f"etl_report_{TIMESTAMP}.md"

    lines = [
        f"# 📊 Laporan ETL Pipeline - Data Inovasi & Kewirausahaan FIF",
        f"",
        f"**Dibuat:** {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"**Pipeline:** `code/etl_fif_inovasi.py`",
        f"",
        f"---",
        f"",
    ]

    # ── Ringkasan umum
    lines += [
        "## 📋 Ringkasan Umum",
        "",
        "| Dataset | File Sumber | Baris Awal | Baris Akhir | Output |",
        "|---------|------------|:----------:|:-----------:|--------|",
    ]

    file_map = {
        "ki": ("Data KI FIF", "clean_data_ki_fif.csv"),
        "startup": ("Data Startup FIF", "clean_data_startup_fif.csv"),
        "inkubasi": ("Peserta Inkubasi FIF", "clean_peserta_inkubasi_fif.csv"),
    }

    for key, (label, outfile) in file_map.items():
        log = logs[key]
        awal = log.get("baris_awal", log.get("baris_awal_raw", "?"))
        akhir = log.get("baris_akhir", "?")
        lines.append(f"| {label} | `{log['file']}` | {awal} | {akhir} | `{outfile}` |")

    lines += ["", "---", ""]

    # ── Detail per pipeline
    # Pipeline 1: KI
    log_ki = logs["ki"]
    lines += [
        "## 1️⃣ Data KI FIF (Kekayaan Intelektual)",
        "",
        f"**Sumber:** `{log_ki['file']}`  ",
        f"**Baris awal:** {log_ki['baris_awal']} → **Baris akhir:** {log_ki['baris_akhir']}",
        "",
        "### Langkah Pembersihan",
        "",
    ]
    for i, step in enumerate(log_ki["langkah"], 1):
        lines.append(f"{i}. {step}")

    lines += [
        "",
        "### Distribusi Kategori KI",
        "",
        "| Kategori | Jumlah |",
        "|----------|:------:|",
    ]
    for k, v in log_ki.get("kategori_dist", {}).items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "### Distribusi Status",
        "",
        "| Status | Jumlah |",
        "|--------|:------:|",
    ]
    for k, v in log_ki.get("status_dist", {}).items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "### Null Values Setelah Cleaning",
        "",
        "| Kolom | Jumlah Null |",
        "|-------|:-----------:|",
    ]
    for k, v in log_ki.get("null_summary", {}).items():
        lines.append(f"| `{k}` | {v} |")

    lines += ["", "---", ""]

    # Pipeline 2: Startup
    log_st = logs["startup"]
    lines += [
        "## 2️⃣ Data Startup FIF 2025-2026",
        "",
        f"**Sumber:** `{log_st['file']}`  ",
        f"**Baris awal (raw):** {log_st['baris_awal_raw']} → **Baris akhir:** {log_st['baris_akhir']}",
        "",
        "### Langkah Pembersihan",
        "",
    ]
    for i, step in enumerate(log_st["langkah"], 1):
        lines.append(f"{i}. {step}")

    lines += [
        "",
        "### Distribusi per Tahun",
        "",
        "| Tahun | Jumlah Startup |",
        "|:-----:|:--------------:|",
    ]
    for k, v in log_st.get("tahun_dist", {}).items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "### Distribusi per Batch",
        "",
        "| Batch | Jumlah |",
        "|:-----:|:------:|",
    ]
    for k, v in log_st.get("batch_dist", {}).items():
        label = f"Batch {k}" if pd.notna(k) else "Tanpa Batch (Tahun 2025)"
        lines.append(f"| {label} | {v} |")

    lines += ["", "---", ""]

    # Pipeline 3: Inkubasi
    log_ink = logs["inkubasi"]
    lines += [
        "## 3️⃣ Peserta Inkubasi FIF 2025",
        "",
        f"**Sumber:** `{log_ink['file']}`  ",
        f"**Baris awal (raw):** {log_ink['baris_awal_raw']} → **Baris akhir:** {log_ink['baris_akhir']}",
        "",
        "### Langkah Pembersihan",
        "",
    ]
    for i, step in enumerate(log_ink["langkah"], 1):
        lines.append(f"{i}. {step}")

    lines += [
        "",
        f"### Jumlah Startup Unik: {log_ink.get('startup_count', '?')}",
        "",
        "### Distribusi Prodi",
        "",
        "| Program Studi | Jumlah |",
        "|--------------|:------:|",
    ]
    for k, v in log_ink.get("prodi_dist", {}).items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "### Distribusi Universitas",
        "",
        "| Universitas | Jumlah |",
        "|------------|:------:|",
    ]
    for k, v in log_ink.get("univ_dist", {}).items():
        lines.append(f"| {k} | {v} |")

    lines += [
        "",
        "### Null Values Setelah Cleaning",
        "",
        "| Kolom | Jumlah Null |",
        "|-------|:-----------:|",
    ]
    for k, v in log_ink.get("null_summary", {}).items():
        lines.append(f"| `{k}` | {v} |")

    if "nim_duplikat_detail" in log_ink:
        lines += [
            "",
            "### ⚠️ NIM Duplikat (Anggota di Lebih Dari 1 Startup)",
            "",
            "| NIM | Startup |",
            "|-----|---------|",
        ]
        for nim, startups in log_ink["nim_duplikat_detail"].items():
            lines.append(f"| `{nim}` | {', '.join(startups)} |")

    lines += [
        "",
        "---",
        "",
        "## ✅ Catatan",
        "",
        "- Semua output berada di folder `output/`",
        "- File output menggunakan format CSV (UTF-8)",
        "- Kolom boolean menggunakan `True`/`False` (bukan `Y`/`N`)",
        "- `None`/kosong pada CSV direpresentasikan sebagai field kosong",
        f"- Dokumen ini di-generate otomatis pada {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]

    doc_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"📄 Dokumentasi → {doc_path.relative_to(ROOT)}")
    return doc_path


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("  ETL Pipeline: Data Inovasi & Kewirausahaan FIF")
    print("=" * 60)
    print()

    df_ki, log_ki = pipeline_ki()
    df_st, log_st = pipeline_startup()
    df_ink, log_ink = pipeline_inkubasi()

    logs = {"ki": log_ki, "startup": log_st, "inkubasi": log_ink}
    dfs = {"ki": df_ki, "startup": df_st, "inkubasi": df_ink}

    doc_path = generate_docs(logs, dfs)

    print()
    print("=" * 60)
    print("  Pipeline selesai! ✅")
    print(f"  Output   → output/")
    print(f"  Docs     → {doc_path.relative_to(ROOT)}")
    print("=" * 60)
