import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

def clean_amount_to_text(amount_str):
    """
    Bersihkan angka dari PDF dan kembalikan sebagai TEXT (string).
    Tidak melakukan konversi ke float, hanya membersihkan format dasar.
    """
    if not amount_str or str(amount_str).strip() in ['', 'None', 'NaN']:
        return ''
    
    amount_str = str(amount_str).strip()
    
    # Handle tanda minus di akhir (format bank: 1234.56-)
    if amount_str.endswith('-'):
        amount_str = '-' + amount_str[:-1].strip()
    
    return amount_str

def format_date(date_str):
    """Format tanggal dari berbagai format ke DD-MMM-YYYY"""
    if not date_str or str(date_str).strip() in ['', 'None', 'NaN']:
        return ''
    
    date_str = str(date_str).strip()
    
    date_patterns = [
        (r'(\d{2})\s*([A-Za-z]{3})\s*(\d{4})', r'\1-\2-\3'),
        (r'(\d{2})([A-Za-z]{3})(\d{4})', r'\1-\2-\3'),
    ]
    
    for pattern, replacement in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            day = match.group(1)
            month = match.group(2).upper()
            year = match.group(3)
            return f"{day}-{month}-{year}"
    
    return date_str

def extract_from_pdf(pdf_file):
    """
    Ekstrak tabel dari PDF dan mapping kolom sesuai request:
    
    Original → New Column Name:
    - Reference Number → Posting Date (text)
    - Credit → Value Date (text)
    - Debit → Transaction Branch (TEXT, bukan numeric)
    - Balance → Reference Number (text)
    - Branch → Description (text)
    - Description → Debit (text, NO REFORMAT)
    - Posting Date → Credit (text, NO REFORMAT)
    - Value Date → Balance (text, NO REFORMAT)
    """
    all_transactions = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            
            if tables:
                for table in tables:
                    for row in table:
                        if row and len(row) >= 4:
                            # Cek apakah row mengandung data transaksi
                            if any(cell and str(cell).strip() and str(cell)[0].isdigit() for cell in row):
                                # === COLUMN MAPPING (SEMUA TEXT, NO REFORMAT) ===
                                trans = {
                                    'Posting Date': str(row[0]).strip() if len(row) > 0 and row[0] else '',              # Reference Number → Posting Date
                                    'Value Date': str(row[1]).strip() if len(row) > 1 and row[1] else '',                # Credit → Value Date
                                    'Transaction Branch': clean_amount_to_text(row[2]) if len(row) > 2 else '',          # Debit → Transaction Branch (TEXT)
                                    'Reference Number': str(row[3]).strip() if len(row) > 3 and row[3] else '',          # Balance → Reference Number (TEXT)
                                    'Description': str(row[4]).strip() if len(row) > 4 and row[4] else '',               # Branch → Description
                                    'Debit': str(row[5]).strip() if len(row) > 5 and row[5] else '',                     # Description → Debit (NO REFORMAT)
                                    'Credit': str(row[6]).strip() if len(row) > 6 and row[6] else '',                    # Posting Date → Credit (NO REFORMAT)
                                    'Balance': str(row[7]).strip() if len(row) > 7 and row[7] else '',                   # Value Date → Balance (NO REFORMAT)
                                }
                                
                                # Tambahkan jika Transaction Branch tidak kosong
                                if trans['Transaction Branch'] and trans['Transaction Branch'].strip():
                                    all_transactions.append(trans)
    
    return pd.DataFrame(all_transactions)

def process_multiple_pdfs(uploaded_files):
    """Proses multiple PDF files dan gabungkan hasilnya"""
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            df = extract_from_pdf(uploaded_file)
            if not df.empty:
                # Tambahkan kolom sumber file untuk tracking
                df['Source File'] = uploaded_file.name
                all_dfs.append(df)
        except Exception as e:
            st.warning(f"⚠️ Gagal memproses {uploaded_file.name}: {str(e)}")
            continue
    
    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)
    return pd.DataFrame()

def main():
    st.set_page_config(page_title="PDF ke Excel - Multi File", layout="wide")
    st.title("📄 Ekstrak Transaksi Bank dari PDF ke Excel")
    st.markdown("""
    ### 📋 Fitur:
    - ✅ Upload **multiple PDF** sekaligus
    - ✅ Ekstraksi tabel otomatis dari PDF teks
    - ✅ **Tanpa reformat** angka di kolom Debit, Credit, Balance (sesuai asli PDF)
    - ✅ Kolom **Transaction Branch** & **Reference Number** sebagai TEXT
    - ✅ Download hasil sebagai Excel dengan format siap pakai
    
    ### 🔀 Mapping Kolom:
    | Kolom Asli | Kolom Baru | Tipe Data | Keterangan |
    |------------|------------|-----------|------------|
    | Reference Number | Posting Date | Text | 12-15 digit reference |
    | Credit | Value Date | Text | Nilai credit asli |
    | Debit | **Transaction Branch** | **Text** | ✅ Tidak diformat, sesuai PDF |
    | Balance | Reference Number | **Text** | ✅ Dipaksa text |
    | Branch | Description | Text | Nama branch/cabang |
    | Description | Debit | Text | ✅ Tidak diformat, sesuai PDF |
    | Posting Date | Credit | Text | ✅ Tidak diformat, sesuai PDF |
    | Value Date | Balance | Text | ✅ Tidak diformat, sesuai PDF |
    """)
    
    uploaded_files = st.file_uploader(
        "Pilih file PDF (bisa multiple)", 
        type="pdf",
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.info(f"📌 Memproses {len(uploaded_files)} file PDF...")
        
        with st.spinner("Memproses PDF..."):
            try:
                df = process_multiple_pdfs(uploaded_files)
                
                if df.empty:
                    st.error("❌ Tidak ada transaksi yang berhasil diekstrak dari semua file.")
                    st.markdown("""
                    ### Kemungkinan penyebab:
                    1. **Format PDF tidak sesuai** - Pastikan PDF memiliki format teks (bukan hasil scan)
                    2. **Struktur tabel berbeda** - Pastikan ada tabel dengan minimal 4 kolom
                    3. **PDF terkunci/password protected**
                    """)
                else:
                    st.success(f"✅ Berhasil mengekstrak **{len(df)} transaksi** dari {len(uploaded_files)} file!")
                    
                    # === PREPARE DATA FOR DISPLAY ===
                    df_display = df.copy()
                    
                    # Pastikan semua kolom yang diminta sebagai text
                    text_columns = ['Transaction Branch', 'Reference Number', 'Debit', 'Credit', 'Balance']
                    for col in text_columns:
                        if col in df_display.columns:
                            df_display[col] = df_display[col].astype(str).replace('nan', '')
                    
                    # Tampilkan data
                    st.subheader("📊 Preview Data Transaksi")
                    st.dataframe(df_display.head(20), use_container_width=True)
                    
                    # Statistik sederhana (count only, karena data text)
                    st.subheader("📈 Ringkasan")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Transaksi", len(df))
                    col2.metric("File Diproses", len(uploaded_files))
                    col3.metric("Rata-rata per File", f"{len(df)/len(uploaded_files):.1f}" if uploaded_files else "0")
                    
                    # === PREPARE EXCEL EXPORT ===
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_export = df.copy()
                        
                        # Pastikan kolom text tetap sebagai string di Excel
                        for col in ['Transaction Branch', 'Reference Number', 'Debit', 'Credit', 'Balance', 'Posting Date', 'Value Date', 'Description']:
                            if col in df_export.columns:
                                df_export[col] = df_export[col].astype(str).replace('nan', '')
                        
                        # Tambahkan prefix apostrophe untuk Reference Number agar Excel benar-benar treat sebagai text
                        if 'Reference Number' in df_export.columns:
                            df_export['Reference Number'] = "'" + df_export['Reference Number']
                        
                        # Tulis ke sheet Transactions
                        df_export.to_excel(writer, index=False, sheet_name='Transactions')
                        
                        # Sheet Summary
                        stats_df = pd.DataFrame({
                            'Metric': ['Jumlah Transaksi', 'File Diproses', 'Kolom Text'],
                            'Value': [len(df), len(uploaded_files), 'Debit, Credit, Balance, Transaction Branch']
                        })
                        stats_df.to_excel(writer, sheet_name='Summary', index=False)
                        
                        # Sheet Column Info
                        col_info = pd.DataFrame({
                            'New Column Name': ['Posting Date', 'Value Date', 'Transaction Branch', 'Reference Number', 
                                               'Description', 'Debit', 'Credit', 'Balance'],
                            'Original Column': ['Reference Number', 'Credit', 'Debit', 'Balance',
                                               'Branch', 'Description', 'Posting Date', 'Value Date'],
                            'Data Type': ['Text', 'Text', 'Text', 'Text (Forced)',
                                         'Text', 'Text (No Reformat)', 'Text (No Reformat)', 'Text (No Reformat)'],
                            'Format': ['Asli dari PDF', 'Asli dari PDF', 'Asli dari PDF', "Text + prefix '",
                                      'Asli dari PDF', '✅ Tidak diubah', '✅ Tidak diubah', '✅ Tidak diubah']
                        })
                        col_info.to_excel(writer, sheet_name='Column_Info', index=False)
                    
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 Download Excel",
                        data=excel_data,
                        file_name=f"bank_transactions_{len(uploaded_files)}files.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Info tambahan
                    with st.expander("ℹ️ Info Format & Tipe Data"):
                        st.markdown("""
                        **Kolom Tanpa Reformat:**
                        - Debit, Credit, Balance: Ditampilkan persis seperti di PDF (tanpa konversi format angka)
                        - Berguna jika format asli PDF sudah sesuai kebutuhan laporan Anda
                        
                        **Kolom Text:**
                        - Transaction Branch: Disimpan sebagai text (bukan numeric)
                        - Reference Number: Dipaksa text + prefix `'` agar Excel tidak ubah jadi format ilmiah
                        
                        **Multiple File:**
                        - Data dari semua PDF digabung dalam satu sheet `Transactions`
                        - Kolom `Source File` menunjukkan asal masing-masing transaksi
                        """)
                        
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {str(e)}")
                st.info("Silakan periksa format file PDF atau coba file lainnya.")
    else:
        st.info("👆 Silakan upload file PDF untuk mulai ekstraksi")

if __name__ == "__main__":
    main()
