import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time
import json
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. KONFIGURASI HALAMAN & SIDEBAR
# ==========================================
st.set_page_config(layout="wide", page_title="Master Data - Purchasing Regional")

st.sidebar.image("logo.png", width=150) 
st.sidebar.title("Sistem Master Data")
st.sidebar.write("**Purchasing Regional**")
st.sidebar.write("---")

menu = st.sidebar.radio(
    "Pilih Layanan:",
    ["🧹 Pembersihan Nama Baku", "📥 Update Master Data (Input)", "🔍 Cari Vendor", "⚙️ Menu Tambahan"]
)

# ==========================================
# 2. KONFIGURASI GOOGLE SHEETS
# ==========================================
SHEET_ID = "1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ"
GID_MASTER = "0"          # GID untuk Master Item (Tab 1)
GID_VENDOR = "168217676"  # GID untuk Data Vendor (Tab 2)

def get_gspread_client():
    key_dict = json.loads(st.secrets["google_json"])
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(key_dict, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=10)
def load_data(gid):
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={gid}&t={time.time()}"
    df = pd.read_csv(url)
    return df

# --- PERSIAPAN KAMUS PINTAR (MASTER DATA) ---
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    
    # Forward fill untuk kategori yang di-merge (gabung cell)
    if 'KATEGORI' in df_master.columns:
        df_master['KATEGORI'] = df_master['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df_master.columns:
        df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill()
        
    df_master['KATA KUNCI'] = df_master.get('KATA KUNCI', "").fillna("")
    
    # Gabungkan Nama Baku dan Kata Kunci untuk pencarian cerdas
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    
    # OBAT ANTI-ERROR: Hapus paksa data ganda (duplikat) di memori robot
    master_map = df_master.drop_duplicates(subset=['NAMA BAKU']).set_index('NAMA BAKU').to_dict('index')
    
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}")
    st.stop()


# ==========================================
# MENU 1: PEMBERSIHAN NAMA BAKU (AUTO-FILL)
# ==========================================
if menu == "🧹 Pembersihan Nama Baku":
    st.header("Pembersihan & Update Laporan PO")
    st.write("Sistem akan otomatis mendeteksi 'Nama Lapangan', mengubahnya jadi Nama Baku, dan mengisi kolom Kategori/SKU yang kosong.")
    
    file_po = st.file_uploader("Upload Excel PO User (Pastikan ada kolom NAMA ITEM, QTY, dll)", type=["xlsx"])
    
    if file_po:
        df_po = pd.read_excel(file_po)
        
        # Cari kolom nama barang kotor
        kolom_kotor = "NAMA ITEM"
        if kolom_kotor not in df_po.columns.str.upper():
            possible_cols = [c for c in df_po.columns if 'ITEM' in c.upper() or 'NAMA' in c.upper()]
            kolom_kotor = possible_cols[0] if possible_cols else df_po.columns[0]

        if st.button("✨ Bersihkan & Lengkapi Data"):
            hasil_rows = []
            
            for index, row in df_po.iterrows():
                nama_kotor = str(row[kolom_kotor])
                # Pencarian cerdas menggunakan Rapidfuzz
                match = process.extractOne(nama_kotor, list_lookup, scorer=fuzz.token_set_ratio)
                
                if match and match[1] >= 70: # Jika kemiripan di atas 70%
                    baku = lookup_to_baku[match[0]]
                    info = master_map.get(baku, {})
                    
                    # Menyusun data agar pas dengan kolom Google Sheets Anda
                    row_data = {
                        "NAMA ITEM": nama_kotor,
                        "NAMA BAKU": baku,
                        "KATEGORI": info.get('KATEGORI', '-'),
                        "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'),
                        "NOMOR SKU": info.get('NOMOR SKU', '-'),
                        "KET": row.get('KET', '-'),
                        "SATUAN": info.get('SATUAN', row.get('SATUAN', '-')), # Utamakan master data
                        "HARGA": row.get('HARGA', 0),
                        "QTY": row.get('QTY', 0),
                        "VENDOR": row.get('VENDOR', '-'),
                        "GRUP": row.get('GRUP', '-'),
                        "TANGGAL": str(row.get('TANGGAL', '-'))
                    }
                else:
                    # Jika nama terlalu ngawur / tidak ketemu di database
                    row_data = {
                        "NAMA ITEM": nama_kotor, "NAMA BAKU": "⚠️ CEK MANUAL",
                        "KATEGORI": "-", "DETAIL KATEGORI": "-", "NOMOR SKU": "-",
                        "KET": row.get('KET', '-'), "SATUAN": row.get('SATUAN', '-'),
                        "HARGA": row.get('HARGA', 0), "QTY": row.get('QTY', 0),
                        "VENDOR": row.get('VENDOR', '-'), "GRUP": row.get('GRUP', '-'),
                        "TANGGAL": str(row.get('TANGGAL', '-'))
                    }
                hasil_rows.append(row_data)
            
            st.session_state['hasil_bersih'] = pd.DataFrame(hasil_rows)
            st.success("Data berhasil dibersihkan dan dilengkapi!")

        if 'hasil_bersih' in st.session_state:
            st.write("### Preview Hasil Akhir (Siap Kirim):")
            st.dataframe(st.session_state['hasil_bersih'], use_container_width=True)

            if st.button("🚀 TEMBAK KE GOOGLE SHEETS Laporan PO"):
                try:
                    with st.spinner("Sedang mengirim..."):
                        client = get_gspread_client()
                        # Mengirim ke sheet Master Laporan (Tab Index 0 / Paling Kiri)
                        sheet = client.open_by_key(SHEET_ID).get_worksheet(0) 
                        sheet.append_rows(st.session_state['hasil_bersih'].values.tolist())
                        
                        st.success("🔥 BOOM! Semua data berhasil masuk ke Google Sheets!")
                        del st.session_state['hasil_bersih']
                except Exception as e:
                    st.error(f"Gagal kirim: {e}")

# ==========================================
# MENU 2: UPDATE MASTER DATA (INPUT)
# ==========================================
elif menu == "📥 Update Master Data (Input)":
    st.header("Input Master Item Baru")
    st.info("Formulir untuk menambah barang baru ke Master Data.")
    
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
    
    if st.button("💾 Simpan ke Master Data"):
        st.warning("⚠️ Untuk keamanan data, saat ini penambahan master data langsung dilakukan dari Google Sheets.")

# ==========================================
# MENU 3: CARI VENDOR
# ==========================================
elif menu == "🔍 Cari Vendor":
    st.header("Database Vendor - Purchasing Regional")
    st.info("Ketik kata kunci (Nama barang, Kategori, atau Nama Vendor) untuk mencari supplier.")
    
    try:
        df_vendor = load_data(GID_VENDOR)
        df_vendor.columns = df_vendor.columns.str.strip()
        
        keyword = st.text_input("Cari Vendor / Barang:")
        
        if keyword:
            # Cari keyword di 3 kolom sekaligus: Nama Vendor, Kategori, dan Alamat
            mask = (
                df_vendor['NAMA VENDOR'].astype(str).str.contains(keyword, case=False, na=False) |
                df_vendor['KATEGORI'].astype(str).str.contains(keyword, case=False, na=False) |
                df_vendor['ALAMAT'].astype(str).str.contains(keyword, case=False, na=False)
            )
            hasil = df_vendor[mask]
            
            if not hasil.empty:
                st.success(f"Ditemukan {len(hasil)} Vendor yang cocok!")
                
                # Menampilkan data dengan Expand (Bisa dilipat)
                for _, v in hasil.iterrows():
                    with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - ({v.get('KATEGORI', '-')})"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**📍 Alamat:** {v.get('ALAMAT', '-')}")
                            st.write(f"**👤 PIC:** {v.get('PIC', '-')}")
                            st.write(f"**📞 Kontak:** {v.get('KONTAK', '-')}")
                            st.write(f"**📧 Email:** {v.get('EMAIL', '-')}")
                        with col2:
                            st.write(f"**💳 Rekening:** {v.get('REKENING', '-')}")
                            st.write(f"**🏦 Atas Nama:** {v.get('ATAS NAMA REKENING', '-')}")
                            st.write(f"**⏳ TOP:** {v.get('TOP', '-')}")
                            st.write(f"**🪪 NPWP:** {v.get('NPWP', '-')}")
            else:
                st.warning(f"Vendor dengan kata kunci '{keyword}' tidak ditemukan.")
    except Exception as e:
        st.error(f"Gagal memuat data vendor. Pastikan GID '{GID_VENDOR}' sudah benar. Error: {e}")

# ==========================================
# MENU 4: MENU TAMBAHAN
# ==========================================
elif menu == "⚙️ Menu Tambahan":
    st.header("Fitur Mendatang")
    st.write("Ruang ini disiapkan untuk Dashboard Grafik & Rekap PO per bulan.")