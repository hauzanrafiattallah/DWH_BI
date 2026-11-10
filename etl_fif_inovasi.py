"""ETL awal data inovasi dan kewirausahaan FIF (ekstraksi, pembersihan, pemuatan ke CSV).
Jalankan: python etl_fif_inovasi.py <folder_pdf> <folder_keluaran>"""
import re, sys, pathlib
import pdfplumber, pandas as pd

SRC = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else ".")
OUT = pathlib.Path(sys.argv[2] if len(sys.argv) > 2 else "out"); OUT.mkdir(exist_ok=True)
F_KI, F_ST, F_TIM = (next(SRC.glob(p)) for p in ("All_Data_KI*.pdf", "Data_Startup*.pdf", "Tim_Inkubasi*.pdf"))

PRODI = {r"^(s1 )?(teknik )?informatika( - informatika)?$": "S1 Informatika",
         r"^(s1 |prodi )?(sains data|data science|data sains)$": "S1 Sains Data",
         r"^s1 (mbti|manajemen bisnis telekomunikasi.*)$": "S1 MBTI",
         r"^s1 (terapan )?teknologi rekayasa multimedia$": "D4 Teknologi Rekayasa Multimedia",
         r"^s1 sistem informasi$": "S1 Sistem Informasi"}
KAMPUS = {"telu purwokerto": "TelU Purwokerto"}
BOOL = {"y": True, "ya": True, "n": False, "tidak": False}

def norm_prodi(v):
    v = re.sub(r"\s+", " ", str(v or "")).strip()
    for pat, ganti in PRODI.items():
        if re.match(pat, v.lower()): return ganti
    return v or None

def baca_tabel(path):
    baris = []
    with pdfplumber.open(path) as pdf:
        for hal in pdf.pages:
            for tb in hal.extract_tables(): baris += tb
    return baris

# ---- 1. Tim inkubasi (satu baris = satu anggota)
tim = [r for r in baca_tabel(F_TIM) if r[2] and r[2].startswith("Mahasiswa")]
tim = pd.DataFrame([r[:7] for r in tim], columns=["no", "startup", "anggota", "nim", "prodi", "fakultas", "kampus"])
tim["startup"] = tim["startup"].replace("", None).ffill().str.strip()          # sel gabungan
tim["nim"] = tim["nim"].replace("", None)                                       # NIM kosong -> NULL
tim["prodi"] = tim["prodi"].map(norm_prodi)
tim["startup"] = tim["startup"].str.upper().replace({"FLICKNFIT": "FlicknFit"})
tim["dup_nim"] = tim["nim"].notna() & tim.duplicated("nim", keep=False)

# ---- 2. Data startup (satu baris = satu anggota). Tata letak tabel berbeda antarhalaman,
# sehingga baris anggota dibaca dari teks dan startup dihitung dari nomor urut entri.
teks_st = ""
with pdfplumber.open(F_ST) as pdf:
    for hal in pdf.pages: teks_st += (hal.extract_text() or "") + "\n"
FAK = r"(FIF|FRI|FEB|FIT|FIK|FTE|FKS)"
st = []
for ln in teks_st.split("\n"):
    m = re.match(r"Mahasiswa (\d{4})\s*(.*)", ln)
    if not m: continue
    sisa = m.group(2)
    f = re.search(r"\b" + FAK + r"\b", sisa)
    k = re.search(r"(Telkom University|Tel[uU] (?:Bandung|Surabaya|Purwokerto|Jakarta))", sisa)
    p = sisa[:f.start()] if f else sisa
    st.append(dict(anggota=m.group(1), prodi=norm_prodi(p) if f else None,
                   fakultas=f.group(1) if f else None,
                   kampus=KAMPUS.get(k.group(1).lower(), k.group(1)) if k else None))
semua_id = re.findall(r"Mahasiswa (\d{4})", teks_st)   # termasuk anggota pada baris gabungan
id_hitung = pd.Series(semua_id).value_counts()
st = pd.DataFrame(st)
st["id_ganda"] = st.duplicated("anggota", keep=False)
st_bersih = st.drop_duplicates(["anggota", "prodi", "fakultas", "kampus"])

# ---- 3. KI (tabel tidak terdeteksi; dibaca dari teks dan dipecah per nomor urut)
teks = ""
with pdfplumber.open(F_KI) as pdf:
    for hal in pdf.pages: teks += (hal.extract_text() or "") + "\n"
ki = []
for rec in re.split(r"(?m)^(?=\d{1,3} [A-Z])", teks):
    m = re.match(r"(\d{1,3}) ", rec)
    s = re.search(r"(Sertifikat Selesai|Pengajuan DJKI)\s*(.*)", rec, re.S)
    if not (m and s): continue
    kat = next((k for k in ("HAK CIPTA", "PATEN", "DESAIN INDUSTRI", "MEREK") if k in rec), None)
    baris = s.group(2).strip().split("\n")[0]
    ap = kom = None
    if baris[:1] in ("Y", "N") and not baris[1:2].islower():
        ap = baris[0]
        k = re.search(r"(?<![A-Za-z])([YN])(?=\s|$)|(?<=[a-z])([YN])(?=\s|$)", baris[1:])
        kom = (k.group(1) or k.group(2)) if k else None
    ki.append(dict(no=int(m.group(1)), kategori=kat, status=s.group(1),
                   sudah_diaplikasikan=BOOL.get((ap or "").lower()),
                   sudah_dikomersialisasikan=BOOL.get((kom or "").lower()),
                   ada_nama_error="#NAME?" in rec))
ki = pd.DataFrame(ki).drop_duplicates("no")

# ---- Keluaran dan ringkasan profil
ki.to_csv(OUT / "stg_ki.csv", index=False); st_bersih.to_csv(OUT / "stg_startup.csv", index=False)
tim.to_csv(OUT / "stg_tim_inkubasi.csv", index=False)
print("KI baris", len(ki)); print(ki["kategori"].value_counts().to_dict(), ki["status"].value_counts().to_dict())
print("KI aplikasi:", ki["sudah_diaplikasikan"].value_counts(dropna=False).to_dict())
print("KI komersial:", ki["sudah_dikomersialisasikan"].value_counts(dropna=False).to_dict())
print("Startup: kemunculan anggota", len(semua_id), "| ID unik", id_hitung.size, "| ID berulang", int((id_hitung > 1).sum()), "| baris terbaca utuh", len(st),
      "| prodi kosong", int(st["prodi"].isna().sum()), "| fakultas kosong", int(st["fakultas"].isna().sum()), "| kampus kosong", int(st["kampus"].isna().sum()))
print("Tim baris", len(tim), "| tim", tim["startup"].nunique(), "| NIM kosong", tim["nim"].isna().sum(),
      "| NIM ganda", tim.loc[tim["dup_nim"], "nim"].unique().tolist())