import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io

# 1. KONFIGURASI HALAMAN (JUDUL DI TAB BROWSER)
st.set_page_config(layout="wide", page_title="Master Data PO - Panca Budi", page_icon="📦")

# 2. HEADER DENGAN LOGO PANCA BUDI
col1, col2 = st.columns([1, 8]) # Membagi kolom agar logo ada di kiri
with col1:
    # Mengambil logo langsung dari link website Panca Budi
    st.image("https://pancabudi.com/themes/frontend/img/logo.png", width=120)
with col2:
    st.title("Pembersih Master Data PO")
    st.write("Sistem otomatis standarisasi dan pengelompokan nama barang Purchasing.")
st.write("---")

# 3. BACA DATABASE DARI GOOGLE SHEETS
@st.cache_data(ttl=60)
def load_master_data():
    url_sheet = "https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv"
    df = pd.read_csv(url_sheet) 
    
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

st.write("### Masukkan Data PO Kotor")

# 4. TAB INPUT
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

# 5. LOGIKA PENCARIAN PINTAR
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
        file_name="Data_PO_PancaBudi_Bersih.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )