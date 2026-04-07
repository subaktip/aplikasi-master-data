import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time

# 1. KONFIGURASI HALAMAN & SIDEBAR
st.set_page_config(layout="wide", page_title="Master Data - Purchasing Regional")

# --- SIDEBAR MENU ---
# Menggunakan gambar logo.png HD dari dalam GitHub
st.sidebar.image("logo.png", width=150) 
st.sidebar.title("Sistem Master Data")
st.sidebar.write("**Purchasing Regional**")
st.sidebar.write("---")

menu = st.sidebar.radio(
    "Pilih Layanan:",
    ["🧹 Pembersihan Nama Baku", "📥 Update Master Data (Input)", "⚙️ Menu Tambahan (Coming Soon)"]
)

# 2. FUNGSI LOAD DATA (DATABASE GOOGLE SHEETS)
@st.cache_data(ttl=10)
def load_master_data():
    # Link Google Sheets Anda dengan sistem Anti-Badai
    url_sheet = f"https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv&gid=0&t={time.time()}"
    df = pd.read_csv(url_sheet)
    
    # Bersihkan spasi gaib di judul kolom
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

# Persiapan Data Kamus Pintar
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
        st.write("Ketik nama barang untuk melihat semua kecocokan di database beserta skornya.")
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

    # Jika ada data yang diinput (Copy-paste / Excel), lakukan pembersihan
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


# ==========================================
# LOGIKA MENU 2: UPDATE MASTER DATA (INPUT)
# ==========================================
elif menu == "📥 Update Master Data (Input)":
    st.header("Input Data ke Master Database")
    st.info("Fitur ini digunakan untuk menambah barang baru ke dalam Master Data di Google Sheets secara otomatis.")
    
    with st.form("form_input_master"):
        st.write("### Form Tambah Barang Baru")
        new_nama = st.text_input("NAMA BAKU (Nama Resmi Barang):")
        # Mengambil daftar kategori asli secara otomatis dari Google Sheets
        kategori_unik = sorted([str(k) for k in df_master['KATEGORI'].dropna().unique() if str(k).strip() != ""])
        new_kat = st.selectbox("KATEGORI:", kategori_unik)
        new_detail = st.text_input("DETAIL KATEGORI:")
        new_keyword = st.text_area("KATA KUNCI (Singkatan/Nama Lapangan):", help="Pisahkan dengan koma, misal: aki, battery, batere")
        
        submit_update = st.form_submit_button("💾 Simpan ke Google Sheets")
    
    if submit_update:
        if new_nama:
            st.warning("Sistem sedang menyiapkan API penghubung ke Google Sheets...")
            st.success(f"Simulasi Berhasil! Data '{new_nama}' siap dijadwalkan masuk ke database.")
            st.info("Catatan: Fungsi tulis ke Excel belum aktif sepenuhnya. Kita perlu men-setting API Key Google terlebih dahulu.")
        else:
            st.error("Nama Baku tidak boleh kosong!")


# ==========================================
# LOGIKA MENU 3: MENU TAMBAHAN
# ==========================================
elif menu == "⚙️ Menu Tambahan (Coming Soon)":
    st.header("Fitur Mendatang")
    st.write("Ruang kosong ini disiapkan untuk ekspansi sistem Purchasing Anda selanjutnya, seperti:")
    st.write("1. Dashboard Statistik (Grafik PO per bulan)")
    st.write("2. Laporan Perbandingan Harga Vendor")
    st.write("3. Kalkulator SKU Otomatis")
    st.info("Menu ini siap diisi kapan pun Anda membutuhkannya.")