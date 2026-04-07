import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time

st.set_page_config(layout="wide")
st.title("🛠️ Master Data")
st.write("Sistem otomatis yang memahami nama baku resmi dan istilah/singkatan lapangan.")
st.write("---")

# 1. BACA DATABASE ANTI-BADAI
@st.cache_data(ttl=10)
def load_master_data():
    url_sheet = f"https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv&gid=0&t={time.time()}"
    df = pd.read_csv(url_sheet) 
    
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
    map_katakunci = dict(zip(df_master['Lookup'], df_master['KATA KUNCI']))

except Exception as e:
    st.error(f"⚠️ Gagal membaca Google Sheets. Error: {e}")
    st.stop()

# 2. FITUR 3 TAB (DENGAN MESIN PENCARI BARU)
tab1, tab2, tab3 = st.tabs(["📋 Copy-Paste PO", "📁 Upload File Excel", "🔍 Cari Barang Manual"])

df_po = None
kolom_kotor = 'Nama Item User'

with tab1:
    st.info("Gunakan tab ini untuk membersihkan banyak daftar barang sekaligus.")
    teks_po = st.text_area("Paste daftar nama barang di sini:", height=150)
    if st.button("🚀 Start"):
        if teks_po.strip():
            daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
            df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])

with tab2:
    st.info("Gunakan tab ini untuk membersihkan data dari file Excel ratusan baris.")
    file_po = st.file_uploader("Upload Excel Data Kotor", type=["xlsx"])
    if file_po:
        df_po = pd.read_excel(file_po)

# --- INI FITUR PENCARIAN BARUNYA ---
with tab3:
    st.write("### 🔎 Mesin Pencari Master Data")
    st.write("Ketik nama barang untuk melihat semua kecocokan di database beserta skornya.")
    
    kata_cari = st.text_input("Ketik nama barang atau singkatan (contoh: knee):")
    
    if kata_cari:
        # Mesin mengambil 10 data paling mirip
        hasil_cari = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
        
        data_tabel = []
        for match in hasil_cari:
            skor = round(match[1], 2)
            kunci = match[0]
            
            # Hanya tampilkan yang skornya di atas 30% agar tidak muncul barang acak
            if skor >= 30: 
                data_tabel.append({
                    "Skor Kemiripan": f"{skor}%",
                    "Nama Baku di Sistem": map_baku[kunci],
                    "Kata Kunci Terdaftar": map_katakunci[kunci],
                    "Kategori": map_kategori[kunci]
                })
        
        if data_tabel:
            st.dataframe(pd.DataFrame(data_tabel), use_container_width=True)
        else:
            st.warning("⚠️ Tidak ada barang yang mirip di database.")

# 3. PROSES PEMBERSIHAN OTOMATIS (TAB 1 & 2)
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
                hasil_nama.append(f"⚠️ Cek Manual (Mungkin: {map_baku[kunci_ditemukan]})")
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
    st.dataframe(df_po, use_container_width=True)

    st.write("---")
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_po.to_excel(writer, index=False, sheet_name='Hasil_Pembersihan')
    
    st.download_button(
        label="📥 Download Hasil (Excel)",
        data=output.getvalue(),
        file_name="Data_PO_Bersih.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )