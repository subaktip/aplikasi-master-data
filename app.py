import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time
import json
import gspread
from google.oauth2.service_account import Credentials

# 1. KONFIGURASI HALAMAN & SIDEBAR
st.set_page_config(layout="wide", page_title="Master Data - Purchasing Regional")

st.sidebar.image("logo.png", width=150) 
st.sidebar.title("Sistem Master Data")
st.sidebar.write("**Purchasing Regional**")
st.sidebar.write("---")

menu = st.sidebar.radio(
    "Pilih Layanan:",
    ["🧹 Pembersihan Nama Baku", "📥 Update Master Data (Input)", "⚙️ Menu Tambahan (Coming Soon)"]
)

# --- FUNGSI PENGHUBUNG ROBOT GOOGLE SHEETS ---
def get_gspread_client():
    # Mengambil kunci rahasia dari brankas Streamlit
    key_dict = json.loads(st.secrets["google_json"])
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_info(key_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client

# ID Google Sheets Master Data Anda
SHEET_ID = "1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ"

# 2. FUNGSI LOAD DATA (DATABASE GOOGLE SHEETS)
@st.cache_data(ttl=10)
def load_master_data():
    url_sheet = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0&t={time.time()}"
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


# ==========================================
# LOGIKA MENU 1: PEMBERSIHAN NAMA BAKU
# ==========================================
if menu == "🧹 Pembersihan Nama Baku":
    st.header("Pembersihan Master Data PO")
    st.write("Gunakan menu ini untuk menstandarisasi nama barang kotor dari user/lapangan.")
    
    tab_copy, tab_excel, tab_cari = st.tabs(["📋 Copy-Paste", "📁 Upload Excel", "🔍 Cari Manual"])
    df_po = None
    kolom_kotor = 'Nama Item User'

    with tab_copy:
        teks_po = st.text_area("Paste daftar nama barang kotor di sini:", height=150)
        if st.button("🚀 Proses Teks"):
            if teks_po.strip():
                daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
                df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])

    with tab_excel:
        file_po = st.file_uploader("Upload Excel Data Kotor", type=["xlsx"])
        if file_po:
            df_po = pd.read_excel(file_po)

    with tab_cari:
        st.write("### 🔎 Mesin Pencari Master Data")
        kata_cari = st.text_input("Ketik nama barang atau singkatan (contoh: knee):")
        if kata_cari:
            hasil_cari = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
            data_tabel = []
            for match in hasil_cari:
                skor = round(match[1], 2)
                kunci = match[0]
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

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_po.to_excel(writer, index=False, sheet_name='Hasil_Pembersihan')
        
        st.download_button("📥 Download Hasil (Excel)", data=output.getvalue(), file_name="Data_PO_Bersih.xlsx")


# ==========================================
# LOGIKA MENU 2: UPDATE MASTER DATA (INPUT)
# ==========================================
elif menu == "📥 Update Master Data (Input)":
    st.header("Input Data ke Master Database")
    st.info("Fitur ini akan langsung mengirim data baru ke Google Sheets Panca Budi.")
    
    st.write("### Form Tambah Barang Baru")
    new_nama = st.text_input("NAMA BAKU (Nama Resmi Barang):")
    
    kategori_unik = sorted([str(k) for k in df_master['KATEGORI'].dropna().unique() if str(k).strip() != ""])
    new_kat = st.selectbox("KATEGORI:", kategori_unik)
    
    detail_terkait = df_master[df_master['KATEGORI'] == new_kat]['DETAIL KATEGORI'].dropna().unique()
    detail_unik = sorted([str(d) for d in detail_terkait if str(d).strip() != ""])
    detail_unik.append("✨ + Tambah Detail Kategori Baru...")
    
    new_detail_pilihan = st.selectbox("DETAIL KATEGORI:", detail_unik)
    if new_detail_pilihan == "✨ + Tambah Detail Kategori Baru...":
        new_detail = st.text_input("Ketik Detail Kategori Baru:")
    else:
        new_detail = new_detail_pilihan
        
    new_keyword = st.text_area("KATA KUNCI (Singkatan/Nama Lapangan):", help="Pisahkan dengan koma")
    
    submit_update = st.button("🚀 TEMBAK KE GOOGLE SHEETS!")
    
    if submit_update:
        if new_nama:
            with st.spinner("Mengirim data ke Google Sheets..."):
                try:
                    # Sambungkan robotnya
                    client = get_gspread_client()
                    sheet = client.open_by_key(SHEET_ID).sheet1
                    
                    # Tembakkan ke baris paling bawah (Urutan: Kategori, Detail, Nama Baku, Kata Kunci)
                    # Note: Sesuaikan urutan array ini jika kolom di Google Sheets Anda berbeda urutannya.
                    sheet.append_row([new_kat, new_detail, new_nama, new_keyword])
                    
                    st.success(f"🔥 BERHASIL! Data '{new_nama}' langsung meluncur dan tersimpan di Google Sheets!")
                    # Membersihkan cache agar pencarian langsung terupdate
                    st.cache_data.clear()
                    
                except Exception as e:
                    st.error(f"Gagal mengirim data. Error: {e}")
        else:
            st.error("Nama Baku tidak boleh kosong!")


# ==========================================
# LOGIKA MENU 3: MENU TAMBAHAN
# ==========================================
elif menu == "⚙️ Menu Tambahan (Coming Soon)":
    st.header("Fitur Mendatang")
    st.write("Ruang kosong ini disiapkan untuk ekspansi sistem Purchasing Anda selanjutnya.")