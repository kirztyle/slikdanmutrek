import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz, process
import io
import re

st.set_page_config(page_title="Fuzzy Matching Nama – DHN", layout="wide")

st.title("🔍 Fuzzy Matching Nama – DHN Screening")

# =====================
# 🔧 Helper Functions
# =====================

def clean_numeric_string(value):
    """
    Bersihkan nilai numeric string: hapus .0, handle NaN, pastikan string.
    """
    if pd.isna(value) or value == 'nan':
        return ''
    str_val = str(value).strip()
    if str_val.endswith('.0') and str_val[:-2].replace('.', '').isdigit():
        return str_val[:-2]
    return str_val


def format_columns_as_text(df, columns):
    """
    Format kolom tertentu sebagai text clean untuk export Excel.
    """
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = df[col].apply(clean_numeric_string)
    return df


def parse_txt_line(line):
    """
    Parse satu baris dari file TXT tanpa header sesuai formula Excel:
    - Nama: =MID(A1;80;40)  → posisi 80, panjang 40 (Excel 1-based → Python 79)
    - NPWP: =MID(A1;140;17) → posisi 140, panjang 17 (Excel 1-based → Python 139)
    """
    try:
        # Excel MID menggunakan 1-based indexing, Python string 0-based
        nama = line[79:79+40].strip() if len(line) >= 79 else ''
        npwp = line[139:139+17].strip() if len(line) >= 139 else ''
        return {"Nama": nama, "NPWP": npwp}
    except Exception:
        return {"Nama": '', "NPWP": ''}


def read_txt_sumber(file):
    """
    Baca file TXT tanpa header dan parse sesuai formula MID.
    Return DataFrame dengan kolom: Nama, NPWP
    """
    data = []
    # Decode bytes ke string
    content = file.getvalue().decode('utf-8', errors='ignore')
    lines = content.splitlines()
    
    for line in lines:
        if line.strip():  # Skip baris kosong
            parsed = parse_txt_line(line)
            if parsed["Nama"] or parsed["NPWP"]:  # Hanya tambahkan jika ada data
                data.append(parsed)
    
    return pd.DataFrame(data, columns=["Nama", "NPWP"])


# =====================
# Upload File
# =====================
st.subheader("📤 Upload File")

# ✅ Label disesuaikan dengan AML
file_nama_dicari = st.file_uploader(
    "Upload data nasabah dari DWH yang mau dicari",
    type=["xlsx"]
)

# ✅ MODIFIKASI: Multiple TXT file upload untuk sumber data
files_sumber_txt = st.file_uploader(
    "Upload File Sumber Data (format .txt tanpa header)",
    type=["txt"],
    accept_multiple_files=True  # ✅ Izinkan upload banyak file TXT
)

if not file_nama_dicari or not files_sumber_txt:
    st.info("Silakan upload file 'Nama Dicari' (Excel) dan minimal 1 file 'Sumber Data' (TXT) untuk memulai proses.")
    st.stop()

# =====================
# Baca File Nama Dicari (Excel)
# =====================
# ✅ Baca dengan dtype=str untuk kolom identifier agar tidak jadi float
df_nama = pd.read_excel(
    file_nama_dicari,
    dtype={
        'CIF': str,
        'NAMA': str, 
        'STATUS': str,
        'NIK': str,
        'NPWP': str
    }
)

# =====================
# ✅ MODIFIKASI: Compile multiple TXT files jadi satu DataFrame
# =====================
st.subheader("📦 Kompilasi Sumber Data (TXT)")
df_sumber_list = []

for file in files_sumber_txt:
    try:
        df_temp = read_txt_sumber(file)
        
        # Validasi kolom
        if "Nama" not in df_temp.columns or "NPWP" not in df_temp.columns:
            st.warning(f"⚠️ File '{file.name}' tidak memiliki kolom required (Nama | NPWP). Dilewati.")
            continue
        
        df_temp["_source_file"] = file.name  # Tandai asal file (opsional)
        df_sumber_list.append(df_temp)
        st.success(f"✅ Loaded: {file.name} ({len(df_temp)} baris)")
    except Exception as e:
        st.error(f"❌ Gagal membaca file '{file.name}': {str(e)}")

if not df_sumber_list:
    st.error("Tidak ada file sumber data TXT yang valid. Proses dihentikan.")
    st.stop()

# Gabungkan semua DataFrame sumber
df_sumber = pd.concat(df_sumber_list, ignore_index=True)
st.info(f"📊 Total data sumber setelah dikompilasi: **{len(df_sumber):,}** baris dari {len(files_sumber_txt)} file")

# Bersihkan data sumber: hapus baris dengan Nama kosong
df_sumber = df_sumber[df_sumber["Nama"].str.strip().astype(bool)].reset_index(drop=True)

# =====================
# Validasi Kolom
# =====================
# ✅ Kolom disesuaikan dengan AML (uppercase)
required_nama = ["CIF", "NAMA", "STATUS", "NIK", "NPWP"]
required_sumber = ["Nama", "NPWP"]

if not all(col in df_nama.columns for col in required_nama):
    st.error(f"File Nama Dicari harus memiliki kolom: {required_nama}")
    st.stop()

if not all(col in df_sumber.columns for col in required_sumber):
    st.error(f"Data sumber terkompilasi harus memiliki kolom: {required_sumber}")
    st.stop()

# =====================
# Proses Fuzzy Matching
# =====================
st.subheader("⚙️ Proses Matching")

threshold = st.slider(
    "Minimum Persentase Kemiripan",
    min_value=50,
    max_value=100,
    value=70
)

total_data = len(df_nama)
progress = st.progress(0)
log_box = st.empty()

log_box.info("📌 Menyiapkan data sumber terkompilasi...")
sumber_nama_list = df_sumber["Nama"].dropna().unique().tolist()

log_box.info(f"🗂️ Unique nama di sumber data: {len(sumber_nama_list):,}")

hasil = []

log_box.info(f"🚀 Memulai fuzzy matching untuk {total_data} data...")

for idx, row in df_nama.iterrows():
    nama_dicari = row["NAMA"]

    if idx % 10 == 0:  # Update log tiap 10 data agar tidak terlalu spam
        log_box.write(
            f"🔍 ({idx + 1}/{total_data}) Memproses: **{nama_dicari}**"
        )

    match_result = process.extractOne(
        nama_dicari,
        sumber_nama_list,
        scorer=fuzz.token_sort_ratio
    )

    if match_result:
        match, score = match_result[0], match_result[1]
        # Ambil NPWP dari sumber - jika ada duplikat nama, ambil yang pertama
        mask = df_sumber["Nama"] == match
        npwp_sumber = df_sumber.loc[mask, "NPWP"].iloc[0] if mask.any() else "-"
    else:
        match, score, npwp_sumber = None, 0, "-"

    hasil.append({
        "CIF": row["CIF"],
        "Nama Dicari": nama_dicari,
        "Status": row["STATUS"],
        "Nama di Sumber Data": match,
        "NPWP Sumber": npwp_sumber,
        "NIK": row["NIK"],
        "NPWP Dicari": row["NPWP"],
        "Persentase Kemiripan Nama": score
    })

    progress.progress((idx + 1) / total_data)

log_box.success("✅ Proses fuzzy matching selesai.")

df_hasil = pd.DataFrame(hasil)

# ✅ Filter berdasarkan threshold (opsional, bisa di-toggle)
if st.checkbox("✅ Hanya tampilkan hasil ≥ threshold", value=True):
    df_hasil = df_hasil[df_hasil["Persentase Kemiripan Nama"] >= threshold]

df_hasil = df_hasil.sort_values(
    "Persentase Kemiripan Nama", ascending=False
).reset_index(drop=True)

# =====================
# Tampilkan Hasil
# =====================
st.subheader("📊 Hasil Fuzzy Matching")

col1, col2 = st.columns(2)
with col1:
    st.metric("Total Data Dicari", total_data)
with col2:
    st.metric(f"Match (≥{threshold}%)", len(df_hasil))

st.dataframe(df_hasil, use_container_width=True)

# =====================
# Input Nama Report (Opsional)
# =====================
st.subheader("⚙️ Pengaturan Export")

report_suffix = st.text_input(
    "📝 Masukkan keterangan untuk nama file (opsional)",
    placeholder="Contoh: DHN-Q1-2026, Cabang-Surabaya, dll",
    key="dhn_report_suffix_input"
)

# Generate filename dinamis - null safe
base_filename = "Report DHN Screening"
suffix_clean = (report_suffix or "").strip()
if suffix_clean:
    # Sanitasi karakter yang tidak diizinkan di filename
    suffix_safe = re.sub(r'[<>:"/\\|?*]', '_', suffix_clean)
    final_filename = f"{base_filename} - {suffix_safe}.xlsx"
else:
    final_filename = f"{base_filename}.xlsx"

st.caption(f"📄 Nama file akan: `{final_filename}`")

# =====================
# Download Excel
# =====================
# ✅ Pastikan NIK dan NPWP bertipe text clean
df_export = df_hasil.copy()
df_export = format_columns_as_text(df_export, ["NIK", "NPWP Dicari", "NPWP Sumber", "CIF"])

buffer = io.BytesIO()  # ✅ FIX: Tambahkan () untuk instantiate object

with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_export.to_excel(writer, index=False)
    
    # ✅ Format sel sebagai Text di Excel (mencegah scientific notation)
    worksheet = writer.sheets["Sheet1"]
    for col in ["NIK", "NPWP Dicari", "NPWP Sumber", "CIF"]:
        if col in df_export.columns:
            col_idx = df_export.columns.get_loc(col) + 1  # openpyxl 1-based index
            for row in range(2, len(df_export) + 2):  # skip header row
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = "@"  # Text format di Excel

buffer.seek(0)

st.download_button(
    label="⬇️ Download Hasil (Excel)",
    data=buffer.getvalue(),
    file_name=final_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    key="dhn_download_excel_btn"
)

# =====================
# Informasi Tambahan
# =====================
with st.expander("ℹ️ Informasi Kolom & Formula Parsing TXT"):
    st.markdown("""
    **Kolom Hasil:**
    1. **CIF**: Customer Identification File
    2. **Nama Dicari**: Nama dari file DWH (Excel)
    3. **Status**: Status dari file DWH
    4. **Nama di Sumber Data**: Hasil matching dari file TXT sumber
    5. **NPWP Sumber**: NPWP dari file TXT sumber yang match
    6. **NIK**: NIK dari file DWH
    7. **NPWP Dicari**: NPWP dari file DWH
    8. **Persentase Kemiripan Nama**: Score fuzzy matching (0-100%)
    
    **📋 Formula Parsing File TXT (tanpa header):**
    | Field | Formula Excel | Posisi Python | Panjang |
    |-------|--------------|---------------|---------|
    | Nama | `=MID(A1;80;40)` | `[79:119]` | 40 karakter |
    | NPWP | `=MID(A1;140;17)` | `[139:156]` | 17 karakter |
    
    > ⚠️ Pastikan file TXT memiliki format fixed-width sesuai posisi di atas.
    """)
