import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io  # Modul baru untuk membuat file Excel yang bisa didownload

# 1. Tampilan Header Website
st.title("🛠️ Pembersih Master Data PO")
st.write("Upload file PO dari user untuk disamakan dengan Master Data Purchasing.")

# 2. Kolom Upload File
col1, col2 = st.columns(2)
with col1:
    file_master = st.file_uploader("Upload Master Data (Excel)", type=["xlsx"])
with col2:
    file_po = st.file_uploader("Upload Data PO Kotor (Excel)", type=["xlsx"])

# 3. Logika Pemrosesan Data
if file_master and file_po:
    df_master = pd.read_excel(file_master)
    df_po = pd.read_excel(file_po)
    
    st.write("### 📄 Data PO Asli dari User:")
    st.dataframe(df_po.head())
    
    st.write("---")
    st.write("Memproses pencocokan data... ⚙️")
    
    # Ambil list nama baku
    list_baku = df_master['Nama Baku'].tolist()
    
    hasil_nama = []
    hasil_skor = []
    
    # ⚠️ PERHATIAN: JIKA JUDUL KOLOM DI EXCEL ASLI ANDA BUKAN INI, SILAKAN DIGANTI LAGI YA
    kolom_kotor = 'Nama Item User' 
    
    for nama_kotor in df_po[kolom_kotor]:
        match = process.extractOne(str(nama_kotor), list_baku, scorer=fuzz.token_set_ratio)
        
        if match:
            skor = round(match[1], 2)
            if skor >= 80:
                hasil_nama.append(match[0])
            else:
                hasil_nama.append(f"⚠️ Cek Manual (Maksudnya: {match[0]}?)")
            hasil_skor.append(skor)
        else:
            hasil_nama.append("Tidak Ditemukan")
            hasil_skor.append(0)
            
    df_po['Nama Baku (Hasil Mapping)'] = hasil_nama
    df_po['Akurasi (%)'] = hasil_skor
    
    st.write("### ✨ Hasil Pembersihan Data:")
    st.dataframe(df_po)

    # 4. FITUR DOWNLOAD EXCEL (BARU)
    st.write("---")
    
    # Mengubah tabel hasil ke format Excel di belakang layar
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_po.to_excel(writer, index=False, sheet_name='Hasil_Mapping')
    hasil_excel = output.getvalue()
    
    # Menampilkan tombol download
    st.download_button(
        label="📥 Download Hasil (Excel)",
        data=hasil_excel,
        file_name="Hasil_Pembersihan_PO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )