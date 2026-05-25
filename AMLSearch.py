import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz, process
import io

st.set_page_config(page_title="Fuzzy Matching Nama", layout="wide")

st.title("🔍 Fuzzy Matching Nama – AML Screening")

# =====================
# Upload File
# =====================
st.subheader("📤 Upload File")

file_nama_dicari = st.file_uploader(
    "Upload data nasabah dari DWH yang mau dicari ",
    type=["xlsx"]
)

# ✅ MODIFIKASI: Multiple file upload untuk sumber data
files_sumber_data = st.file_uploader(
    "Upload File AML",
    type=["xlsx"],
    accept_multiple_files=True  # ✅ Izinkan upload banyak file
)

if not file_nama_dicari or not files_sumber_data:
    st.info("Silakan upload file 'Nama Dicari' dan minimal 1 file 'Sumber Data' untuk memulai proses.")
    st.stop()

# =====================
# Baca File Nama Dicari
# =====================
df_nama = pd.read_excel(file_nama_dicari)

# =====================
# ✅ MODIFIKASI: Compile multiple source files jadi satu DataFrame
# =====================
st.subheader("📦 Kompilasi Sumber Data")
df_sumber_list = []

for file in files_sumber_data:
    try:
        df_temp = pd.read_excel(file)
        # Validasi kolom per file
        if "Watchlist Name" not in df_temp.columns or "Upload Type" not in df_temp.columns:
            st.warning(f"⚠️ File '{file.name}' tidak memiliki kolom required (Watchlist Name | Upload Type). Dilewati.")
            continue
        df_temp["_source_file"] = file.name  # Tandai asal file (opsional, untuk tracking)
        df_sumber_list.append(df_temp)
        st.success(f"✅ Loaded: {file.name} ({len(df_temp)} baris)")
    except Exception as e:
        st.error(f"❌ Gagal membaca file '{file.name}': {str(e)}")

if not df_sumber_list:
    st.error("Tidak ada file sumber data yang valid. Proses dihentikan.")
    st.stop()

# Gabungkan semua DataFrame sumber
df_sumber = pd.concat(df_sumber_list, ignore_index=True)
st.info(f"📊 Total data sumber setelah dikompilasi: **{len(df_sumber):,}** baris dari {len(files_sumber_data)} file")

# =====================
# Validasi Kolom (setelah compile)
# =====================
required_nama = ["CIF", "NAMA", "STATUS", "NIK", "NPWP"]
required_sumber = ["Watchlist Name", "Upload Type"]

if not all(col in df_nama.columns for col in required_nama):
    st.error(f"File Nama Dicari harus memiliki kolom: {required_nama}")
    st.stop()

# Kolom sumber sudah divalidasi saat loop compile, tapi double-check:
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

# ✅ Opsional: Filter hanya data dengan score >= threshold di akhir, 
# atau tampilkan semua dengan highlight

total_data = len(df_nama)
progress = st.progress(0)
log_box = st.empty()

log_box.info("📌 Menyiapkan data sumber terkompilasi...")
sumber_nama_list = df_sumber["Watchlist Name"].dropna().unique().tolist()

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
        # Ambil UploadType - jika ada duplikat nama, ambil yang pertama
        upload_type = df_sumber.loc[df_sumber["Watchlist Name"] == match, "Upload Type"].iloc[0] if not df_sumber[df_sumber["Watchlist Name"] == match].empty else "-"
    else:
        match, score, upload_type = None, 0, "-"

    hasil.append({
        "CIF": row["CIF"],
        "Nama Dicari": nama_dicari,
        "Status": row["STATUS"],
        "Nama di Sumber Data": match,
        "Upload Type": upload_type,
        "NIK": row["NIK"],
        "NPWP": row["NPWP"],
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
st.dataframe(df_hasil, use_container_width=True)

# =====================
# Download Excel
# =====================
# ✅ Perbaiki: to_excel() tidak return bytes, perlu buffer
buffer = io.BytesIO
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    df_hasil.to_excel(writer, index=False)

st.download_button(
    label="⬇️ Download Hasil (Excel)",
    data=buffer.getvalue(),
    file_name="hasil_fuzzy_matching_streamlit.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
