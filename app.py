import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time

st.set_page_config(layout="wide")
st.title("🛠️ Pembersih Master Data PO (Versi Pintar)")
st.write("Sistem otomatis yang memahami nama baku resmi dan istilah/singkatan lapangan.")
st.write("---")

# 1. BACA DATABASE DENGAN MODE ANTI-BADAI
@st.cache_data(ttl=10) # Ingatan dipendekkan jadi 10 detik saja
def load_master_data():
    url_sheet = f"https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv&gid=0&t={time.time()}"
    df = pd.read_csv(url_sheet) 
    
    # Menghapus spasi gaib di judul kolom otomatis!
    df.columns = df.columns.str.strip().str.upper()
    
    if 'KATEGORI' in df.columns:
        df['KATEGORI'] = df['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df.columns:
        df['DETAIL KATEGORI'] = df['DETAIL KATEGORI'].ffill()
    
    if 'KATA KUNCI' not in df.columns:
        df['KATA KUNCI'] = ""
    else:
        df['KATA KUNCI'] = df['KATA KUNCI'].fillna("")
        
    return df

try:
    df_master = load_master_data()
    
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    list_lookup = df_master['Lookup'].tolist()
    
    map_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
    map_kategori = dict(zip(df_master['Lookup'], df_master['KATEGORI']))
    map_detail = dict(zip(df_master['Lookup'], df_master['DETAIL KATEGORI']))

except Exception as e:
    st.error(f"⚠️ Gagal membaca Google Sheets. Error: {e}")
    st.stop()

# 2. FITUR BARU: INTIP OTAK MESIN 👀
with st.expander("👀 Klik di sini untuk mengintip data yang sedang dibaca oleh aplikasi:"):
    st.info("Cari tulisan 'KNEE' di tabel bawah ini. Jika tidak ada, berarti server Google masih menahan data Anda!")
    st.dataframe(df_master[['NAMA BAKU', 'KATA KUNCI']])

st.write("### Masukkan Data PO Kotor")

tab1, tab2 = st.tabs(["📋 Copy-Paste Teks", "📁 Upload File Excel"])

df_po = None
kolom_kotor = 'Nama Item User'

with tab1:
    teks_po = st.text_area("Paste daftar nama barang di sini:", height=200)
    if st.button("🚀 Proses Teks Copy-Paste"):
        if teks_po.strip():
            daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
            df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])

with tab2:
    file_po = st.file_uploader("Upload Excel Data Kotor", type=["xlsx"])
    if file_po:
        df_po = pd.read_excel(file_po)

if df_po is not None and kolom_kotor in df_po.columns:
    st.write("---")
    st.write("Memproses pencocokan dengan Kamus Pintar... ⚙️")
    
    hasil_nama, hasil_kategori, hasil_detail, hasil_skor = [], [], [], []
    
    for nama_kotor in df_po[kolom_kotor]:
        match = process.extractOne(str(nama_kotor), list_lookup, scorer=fuzz.token_set_ratio)
        
        if match:
            skor = round(match[1], 2)
            kunci_ditemukan = match[0]
            
            if skor >= 70:
                hasil_nama.append(map_baku[kunci_ditemukan])
                hasil_kategori.append(map_kategori[kunci_ditemukan])
                hasil_detail.append(map_detail[kunci_ditemukan])
            else:
                hasil_nama.append(f"⚠️ Cek Manual (Saran: {map_baku[kunci_ditemukan]}?)")
                hasil_kategori.append("-")
                hasil_detail.append("-")
            hasil_skor.append(skor)
        else:
            hasil_nama.append("Tidak Ditemukan")
            hasil_kategori.append("-"); hasil_detail.append("-"); hasil_skor.append(0)
            
    df_po['Nama Baku (Hasil Mapping)'] = hasil_nama
    df_po['Kategori'] = hasil_kategori
    df_po['Detail Kategori'] = hasil_detail
    df_po['Akurasi (%)'] = hasil_skor
    
    st.write("### ✨ Hasil Akhir:")
    st.dataframe(df_po)

    st.write("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_po.to_excel(writer, index=False, sheet_name='Hasil_Pembersihan')
    
    st.download_button(
        label="📥 Download Hasil (Excel)",
        data=output.getvalue(),
        file_name="Data_PO_Bersih_Pintar.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )