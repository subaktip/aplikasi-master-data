import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

st.set_page_config(layout="wide")
st.title("🛠️ Pembersih Master Data PO Otomatis")
st.write("Sistem otomatis penstandarisasi dan pengelompokan nama barang Purchasing.")
st.write("---")

# 1. BACA DATABASE LANGSUNG DARI GOOGLE SHEETS
@st.cache_data(ttl=60)
def load_master_data():
    # Ini Link Sakti Anda yang sudah saya perbaiki 100%
    url_sheet = "https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv"
    
    # Membaca data langsung dari internet
    df = pd.read_csv(url_sheet) 
    
    # Memperbaiki sel kosong (merge cells) di Kategori
    if 'KATEGORI' in df.columns:
        df['KATEGORI'] = df['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df.columns:
        df['DETAIL KATEGORI'] = df['DETAIL KATEGORI'].ffill()
        
    return df

try:
    df_master = load_master_data()
    list_baku = df_master['NAMA BAKU'].dropna().astype(str).tolist()
    dict_kategori = dict(zip(df_master['NAMA BAKU'], df_master['KATEGORI']))
    dict_detail = dict(zip(df_master['NAMA BAKU'], df_master['DETAIL KATEGORI']))
except Exception as e:
    st.error(f"⚠️ Gagal membaca Google Sheets. Pastikan link sudah benar dan aksesnya 'Anyone with the link'. Error: {e}")
    st.stop()

st.write("### Masukkan Data PO Kotor")

# 2. TAB UNTUK PILIHAN INPUT
tab1, tab2 = st.tabs(["📋 Copy-Paste Teks (Cepat)", "📁 Upload File Excel"])

df_po = None
kolom_kotor = 'Nama Item User'

with tab1:
    st.info("Cara cepat: Copy daftar nama barang dari mana saja dan Paste di bawah.")
    teks_po = st.text_area("Paste daftar nama barang kotor di sini (Satu barang per baris):", height=200)
    
    if st.button("🚀 Proses Teks Copy-Paste"):
        if teks_po.strip():
            daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
            df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])
        else:
            st.warning("Kotak teks masih kosong!")

with tab2:
    st.info("Gunakan opsi ini jika data input Anda panjang dan sudah dalam bentuk file Excel.")
    file_po = st.file_uploader("Upload Data PO Kotor (Excel)", type=["xlsx"])
    if file_po:
        df_po = pd.read_excel(file_po)
        if kolom_kotor not in df_po.columns:
            st.error(f"Error: Pastikan ada kolom bernama tepat '{kolom_kotor}' di Excel Anda.")
            df_po = None

# 3. LOGIKA PEMROSESAN & PENCARIAN KATEGORI
if df_po is not None:
    st.write("---")
    st.write("Memproses standarisasi dan pengelompokan... ⚙️")
    
    hasil_nama = []
    hasil_kategori = []
    hasil_detail = []
    hasil_skor = []
    
    for nama_kotor in df_po[kolom_kotor]:
        match = process.extractOne(str(nama_kotor), list_baku, scorer=fuzz.token_set_ratio)
        
        if match:
            skor = round(match[1], 2)
            if skor >= 80:
                nama_baku = match[0]
                hasil_nama.append(nama_baku)
                hasil_kategori.append(dict_kategori.get(nama_baku, "Tidak Ada Kategori"))
                hasil_detail.append(dict_detail.get(nama_baku, "Tidak Ada Detail"))
            else:
                hasil_nama.append(f"⚠️ Cek Manual (Maksudnya: {match[0]}?)")
                hasil_kategori.append("-")
                hasil_detail.append("-")
            hasil_skor.append(skor)
        else:
            hasil_nama.append("Tidak Ditemukan")
            hasil_kategori.append("-")
            hasil_detail.append("-")
            hasil_skor.append(0)
            
    df_po['Nama Baku (Hasil Mapping)'] = hasil_nama
    df_po['Kategori'] = hasil_kategori
    df_po['Detail Kategori'] = hasil_detail
    df_po['Akurasi (%)'] = hasil_skor
    
    st.write("### ✨ Hasil Akhir:")
    st.dataframe(df_po)

    # 4. FITUR DOWNLOAD EXCEL
    st.write("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_po.to_excel(writer, index=False, sheet_name='Hasil_Pembersihan')
    hasil_excel = output.getvalue()
    
    st.download_button(
        label="📥 Download Hasil (Excel)",
        data=hasil_excel,
        file_name="Data_PO_Bersih_Berkategori.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )