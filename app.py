import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

# 1. Tampilan Header Website
st.title("🛠️ Pembersih Master Data PO")
st.write("Standarisasi nama barang dari user agar sesuai dengan Master Data Purchasing.")

# Master Data wajib diupload di awal sebagai kamus acuan
file_master = st.file_uploader("1️⃣ Upload Master Data Purchasing (Excel)", type=["xlsx"])

if file_master:
    df_master = pd.read_excel(file_master)
    list_baku = df_master['Nama Baku'].tolist()
    
    st.write("---")
    st.write("### 2️⃣ Masukkan Data PO Kotor")
    
    # MEMBUAT 2 TAB UNTUK PILIHAN INPUT
    tab1, tab2 = st.tabs(["📁 Upload File Excel", "📋 Copy-Paste Teks"])
    
    df_po = None
    kolom_kotor = 'Nama Item User'
    
    with tab1:
        st.info("Gunakan opsi ini jika data Anda sudah dalam bentuk file Excel.")
        file_po = st.file_uploader("Upload Data PO Kotor (Excel)", type=["xlsx"])
        if file_po:
            df_po = pd.read_excel(file_po)
            # Cek keamanan jika judul kolom di excel salah
            if kolom_kotor not in df_po.columns:
                st.error(f"Error: Pastikan ada kolom bernama tepat '{kolom_kotor}' di Excel Anda.")
                df_po = None
                
    with tab2:
        st.info("Gunakan opsi ini untuk cara cepat. Copy nama barang dari mana saja dan Paste di bawah.")
        teks_po = st.text_area("Paste daftar nama barang kotor di sini (Satu barang per baris):", height=200)
        
        # Tombol khusus untuk memproses teks
        if st.button("🚀 Proses Teks Copy-Paste"):
            if teks_po.strip():
                # Memecah teks per baris (enter) dan membuang baris kosong
                daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
                # Menyulap teks menjadi tabel DataFrame persis seperti Excel
                df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])
            else:
                st.warning("Kotak teks masih kosong! Silakan paste data terlebih dahulu.")

    # 3. Logika Pemrosesan (Berjalan jika df_po sudah terisi dari Tab 1 atau Tab 2)
    if df_po is not None:
        st.write("---")
        st.write("Memproses pencocokan data... ⚙️")
        
        hasil_nama = []
        hasil_skor = []
        
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

        # 4. FITUR DOWNLOAD EXCEL
        st.write("---")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_po.to_excel(writer, index=False, sheet_name='Hasil_Mapping')
        hasil_excel = output.getvalue()
        
        st.download_button(
            label="📥 Download Hasil (Excel)",
            data=hasil_excel,
            file_name="Hasil_Pembersihan_PO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )