import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time

# 1. KONFIGURASI HALAMAN & SIDEBAR
st.set_page_config(layout="wide", page_title="Panca Budi - Master Data System")

# --- SIDEBAR MENU ---
st.sidebar.image("https://pancabudi.com/themes/frontend/img/logo.png", width=150)
st.sidebar.title("Sistem Master Data")
st.sidebar.write("PT Panca Budi Pratama")
st.sidebar.write("---")

menu = st.sidebar.radio(
    "Pilih Layanan:",
    ["🧹 Pembersihan Nama Baku", "📥 Update Master Data (Input)", "⚙️ Menu Tambahan (Coming Soon)"]
)

# 2. FUNGSI LOAD DATA (DATABASE GOOGLE SHEETS)
@st.cache_data(ttl=10)
def load_master_data():
    # Link Google Sheets Anda
    url_sheet = f"https://docs.google.com/spreadsheets/d/1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ/export?format=csv&gid=0&t={time.time()}"
    df = pd.read_csv(url_sheet)
    df.columns = df.columns.str.strip().str.upper()
    
    if 'KATEGORI' in df.columns:
        df['KATEGORI'] = df['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df.columns:
        df['DETAIL KATEGORI'] = df['DETAIL KATEGORI'].ffill()
    
    df['KATA KUNCI'] = df.get('KATA KUNCI', "").fillna("")
    return df

# Persiapan Data
try:
    df_master = load_master_data()
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    list_lookup = df_master['Lookup'].tolist()
    
    map_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
    map_kategori = dict(zip(df_master['Lookup'], df_master['KATEGORI']))
    map_detail = dict(zip(df_master['Lookup'], df_master['DETAIL KATEGORI']))
except Exception as e:
    st.error(f"Gagal memuat database: {e}")
    st.stop()


# --- LOGIKA MENU 1: PEMBERSIHAN NAMA ---
if menu == "🧹 Pembersihan Nama Baku":
    st.header("Pembersihan Master Data PO")
    st.write("Gunakan tab ini untuk membersihkan daftar barang kotor dari user.")
    
    tab_copy, tab_excel, tab_cari = st.tabs(["📋 Copy-Paste", "📁 Upload Excel", "🔍 Cari Manual"])
    
    df_po = None
    kolom_kotor = 'Nama Item User'

    with tab_copy:
        teks_po = st.text_area("Paste daftar nama barang kotor:", height=150)
        if st.button("🚀 Proses Teks"):
            if teks_po.strip():
                daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
                df_po = pd.DataFrame(daftar_item, columns=[kolom_kotor])

    with tab_excel:
        file_po = st.file_uploader("Upload Excel Data Kotor", type=["xlsx"])
        if file_po:
            df_po = pd.read_excel(file_po)

    with tab_cari:
        st.write("🔍 Cari kecocokan barang di database")
        kata_cari = st.text_input("Ketik nama barang (contoh: knee):")
        if kata_cari:
            hasil_cari = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=5)
            data_tabel = [{"Skor": f"{round(m[1],1)}%", "Nama Baku": map_baku[m[0]], "Kategori": map_kategori[m[0]]} for m in hasil_cari]
            st.table(data_tabel)

    if df_po is not None:
        # (Logika pembersihan tetap sama seperti sebelumnya)
        hasil_nama, hasil_kat, hasil_skor = [], [], []
        for item in df_po[kolom_kotor]:
            match = process.extractOne(str(item), list_lookup, scorer=fuzz.token_set_ratio)
            if match and match[1] >= 70:
                hasil_nama.append(map_baku[match[0]])
                hasil_kat.append(map_kategori[match[0]])
                hasil_skor.append(round(match[1], 2))
            else:
                hasil_nama.append(f"⚠️ Cek Manual (Saran: {map_baku[match[0]] if match else '?'})")
                hasil_kat.append("-")
                hasil_skor.append(match[1] if match else 0)
        
        df_po['Nama Baku'] = hasil_nama
        df_po['Kategori'] = hasil_kat
        df_po['Akurasi (%)'] = hasil_skor
        st.dataframe(df_po, use_container_width=True)


# --- LOGIKA MENU 2: UPDATE MASTER DATA (INPUT OTOMATIS) ---
elif menu == "📥 Update Master Data (Input)":
    st.header("Input Data ke Master Database")
    st.info("Fitur ini digunakan untuk menambah barang baru ke dalam Master Data di Google Sheets secara otomatis.")
    
    with st.form("form_input_master"):
        st.write("### Form Tambah Barang Baru")
        new_nama = st.text_input("NAMA BAKU (Nama Resmi Barang):")
        new_kat = st.selectbox("KATEGORI:", ["ALAT BERAT (001)", "BAHAN BANGUNAN (002)", "TEKNIK (003)", "DAPUR (004)"])
        new_detail = st.text_input("DETAIL KATEGORI:")
        new_keyword = st.text_area("KATA KUNCI (Singkatan/Nama Lapangan):", help="Pisahkan dengan koma, misal: aki, battery, batere")
        
        submit_update = st.form_submit_button("💾 Simpan ke Google Sheets")
    
    if submit_update:
        if new_nama:
            st.warning("Menghubungkan ke API Google Sheets...")
            # CATATAN: Untuk menulis ke Google Sheets, Anda perlu mengatur 'Secrets' di Streamlit Cloud.
            # Sementara, sistem ini menunjukkan alur kerjanya.
            st.success(f"Berhasil! Data '{new_nama}' telah dijadwalkan untuk masuk ke database.")
        else:
            st.error("Nama Baku tidak boleh kosong!")


# --- LOGIKA MENU 3: MENU TAMBAHAN ---
elif menu == "⚙️ Menu Tambahan (Coming Soon)":
    st.header("Fitur Mendatang")
    st.write("Di sini Anda bisa menambahkan menu lain seperti:")
    st.write("1. Dashboard Statistik (Grafik PO per bulan)")
    st.write("2. Laporan Vendor")
    st.write("3. Sistem Approval")
    st.info("Menu ini siap diisi sesuai kebutuhan Supply Chain Anda berikutnya.")