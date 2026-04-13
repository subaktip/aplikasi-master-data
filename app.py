import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time
import json
import gspread
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu

# ==========================================
# 1. KONFIGURASI HALAMAN & SIDEBAR ELEGAN
# ==========================================
st.set_page_config(layout="wide", page_title="Master Data - Purchasing Regional")

with st.sidebar:
    st.image("logo.png", width=150) 
    st.title("Sistem Master Data")
    st.write("**Purchasing Regional**")
    st.write("---")
    
    menu = option_menu(
        menu_title="", 
        options=["Pembersihan Nama", "Update Master Data", "Cari Vendor", "Fitur Mendatang"],
        icons=["magic", "database-add", "search", "gear"], 
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "transparent"},
            "icon": {"color": "#2e7b32", "font-size": "16px"}, 
            "nav-link": {
                "font-size": "14px", 
                "text-align": "left", 
                "margin":"0px", 
                "--hover-color": "#e2e6ea", 
                "border-radius": "8px" 
            },
            "nav-link-selected": {"background-color": "#2e7b32", "color": "white", "icon-color": "white"},
        }
    )

# ==========================================
# 2. KONFIGURASI GOOGLE SHEETS
# ==========================================
SHEET_ID = "1MZRYFgzzrmBY2vY5qZRmw_-_jmRg-5eq34Nejin-SaQ"
GID_MASTER = "0"          
GID_VENDOR = "168217676"  

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

# Fungsi bantuan untuk format Rupiah
def format_rupiah(angka):
    try:
        return f"Rp {int(angka):,}".replace(',', '.')
    except:
        return "Rp 0"

# --- PERSIAPAN KAMUS PINTAR ---
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    
    df_master = df_master.dropna(subset=['NAMA BAKU'])
    df_master = df_master[df_master['NAMA BAKU'].astype(str).str.strip().str.lower() != "(blank)"]
    
    if 'KATEGORI' in df_master.columns:
        df_master['KATEGORI'] = df_master['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df_master.columns:
        df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill()
        
    df_master['KATA KUNCI'] = df_master.get('KATA KUNCI', "").fillna("")
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    
    master_map = df_master.drop_duplicates(subset=['NAMA BAKU']).set_index('NAMA BAKU').to_dict('index')
    
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}")
    st.stop()


# ==========================================
# MENU 1: PEMBERSIHAN NAMA BAKU
# ==========================================
if menu == "Pembersihan Nama":
    st.header("Pembersihan Master Data PO")
    st.write("Gunakan menu ini untuk menstandarisasi nama barang kotor dari user/lapangan.")
    
    tab_copy, tab_excel, tab_cari = st.tabs(["Copy-Paste", "Upload Excel", "Cari Manual"])
    
    # --- TAB 1: COPY PASTE ---
    with tab_copy:
        st.write("### Mode Cepat: Copy-Paste Teks")
        st.info("Ketik satu kata atau lebih. Sistem akan menampilkan daftar barang yang paling mendekati.")
        teks_po = st.text_area("Paste daftar nama barang kotor di sini (satu baris untuk satu barang):", height=150)
        
        if st.button("Proses Teks", type="primary"):
            if teks_po.strip():
                daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
                hasil_teks = []
                
                for nama_kotor in daftar_item:
                    matches = process.extract(nama_kotor, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
                    ditemukan = False
                    for match in matches:
                        skor = round(match[1], 2)
                        if skor >= 40: 
                            baku = lookup_to_baku[match[0]]
                            info = master_map.get(baku, {})
                            if not any(d.get('Nama Baku (Sistem)') == baku and d.get('Input User') == nama_kotor for d in hasil_teks):
                                hasil_teks.append({
                                    "Input User": nama_kotor, "Nama Baku (Sistem)": baku,
                                    "Kategori": info.get('KATEGORI', '-'), "Detail Kategori": info.get('DETAIL KATEGORI', '-'),
                                    "Akurasi (%)": skor
                                })
                                ditemukan = True
                    if not ditemukan:
                        hasil_teks.append({
                            "Input User": nama_kotor, "Nama Baku (Sistem)": "⚠️ Tidak Ditemukan",
                            "Kategori": "-", "Detail Kategori": "-", "Akurasi (%)": 0
                        })
                
                df_hasil = pd.DataFrame(hasil_teks)
                if not df_hasil.empty:
                    df_hasil['Akurasi (%)'] = pd.to_numeric(df_hasil['Akurasi (%)'], errors='coerce').fillna(0)
                    df_hasil = df_hasil.sort_values(by=['Input User', 'Akurasi (%)'], ascending=[True, False]).reset_index(drop=True)
                    df_hasil['Akurasi (%)'] = df_hasil['Akurasi (%)'].apply(lambda x: round(x, 1) if isinstance(x, (int, float)) else x)
                st.dataframe(df_hasil, use_container_width=True)

    # --- TAB 2: UPLOAD EXCEL ---
    with tab_excel:
        st.write("### Mode Lengkap: Upload & Tembak ke Laporan")
        file_po = st.file_uploader("Upload Excel PO User (Pastikan ada kolom NAMA ITEM, QTY, dll)", type=["xlsx"])
        
        if file_po:
            df_po = pd.read_excel(file_po)
            df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
            
            kolom_kotor = "NAMA ITEM"
            if kolom_kotor not in df_po.columns:
                possible_cols = [c for c in df_po.columns if 'ITEM' in c or 'NAMA' in c]
                kolom_kotor = possible_cols[0] if possible_cols else df_po.columns[0]

            if st.button("Bersihkan & Lengkapi Data Laporan", type="primary"):
                hasil_rows = []
                for index, row in df_po.iterrows():
                    nama_kotor = str(row[kolom_kotor])
                    match = process.extractOne(nama_kotor, list_lookup, scorer=fuzz.token_set_ratio)
                    
                    if match and match[1] >= 70: 
                        baku = lookup_to_baku[match[0]]
                        info = master_map.get(baku, {})
                        row_data = {
                            "NAMA ITEM": nama_kotor, "NAMA BAKU": baku,
                            "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'),
                            "NOMOR SKU": info.get('NOMOR SKU', '-'), "KET": row.get('KET', '-'),
                            "SATUAN": info.get('SATUAN', row.get('SATUAN', '-')), "HARGA": row.get('HARGA', 0),
                            "QTY": row.get('QTY', 0), "VENDOR": row.get('VENDOR', '-'),
                            "GRUP": row.get('GRUP', '-'), "TANGGAL": str(row.get('TANGGAL', '-'))
                        }
                    else:
                        row_data = {
                            "NAMA ITEM": nama_kotor, "NAMA BAKU": "⚠️ CEK MANUAL",
                            "KATEGORI": "-", "DETAIL KATEGORI": "-", "NOMOR SKU": "-",
                            "KET": row.get('KET', '-'), "SATUAN": row.get('SATUAN', '-'),
                            "HARGA": row.get('HARGA', 0), "QTY": row.get('QTY', 0),
                            "VENDOR": row.get('VENDOR', '-'), "GRUP": row.get('GRUP', '-'),
                            "TANGGAL": str(row.get('TANGGAL', '-'))
                        }
                    hasil_rows.append(row_data)
                
                st.session_state['hasil_bersih_excel'] = pd.DataFrame(hasil_rows)
                st.success("Data berhasil dibersihkan dan dilengkapi!")

            if 'hasil_bersih_excel' in st.session_state:
                df_hasil = st.session_state['hasil_bersih_excel']
                
                tab_detail, tab_rekap = st.tabs(["📄 Detail per PO", "📊 Rekap Bulanan (Konsolidasi)"])
                
                with tab_detail:
                    st.write("### Preview Hasil Akhir (Detail)")
                    # [TAMPILAN RUPIAH] Format khusus tampilan di web, data asli tetap angka
                    df_tampil_detail = df_hasil.copy()
                    df_tampil_detail['HARGA'] = df_tampil_detail['HARGA'].apply(format_rupiah)
                    st.dataframe(df_tampil_detail, use_container_width=True)
                
                with tab_rekap:
                    st.write("### Rekapitulasi Pengadaan per Item")
                    try:
                        df_rekap = df_hasil.copy()
                        df_rekap['QTY'] = pd.to_numeric(df_rekap['QTY'], errors='coerce').fillna(0)
                        df_rekap['HARGA'] = pd.to_numeric(df_rekap['HARGA'], errors='coerce').fillna(0)
                        df_rekap['TOTAL_NILAI'] = df_rekap['QTY'] * df_rekap['HARGA']
                        
                        df_group = df_rekap.groupby(['NAMA BAKU', 'KATEGORI', 'SATUAN']).agg(
                            TOTAL_QTY=('QTY', 'sum'),
                            TOTAL_BELANJA=('TOTAL_NILAI', 'sum'),
                            HARGA_TERTINGGI=('HARGA', 'max'),
                            FREKUENSI_ORDER=('NAMA BAKU', 'count')
                        ).reset_index()
                        
                        df_group['HARGA_RATA_RATA'] = (df_group['TOTAL_BELANJA'] / df_group['TOTAL_QTY']).fillna(0).round(0)
                        
                        kolom_tampil = ['NAMA BAKU', 'KATEGORI', 'TOTAL_QTY', 'SATUAN', 'HARGA_RATA_RATA', 'HARGA_TERTINGGI', 'FREKUENSI_ORDER']
                        df_tampil_rekap = df_group[kolom_tampil].copy()
                        
                        # [TAMPILAN RUPIAH] Format Harga untuk Rekap
                        df_tampil_rekap['HARGA_RATA_RATA'] = df_tampil_rekap['HARGA_RATA_RATA'].apply(format_rupiah)
                        df_tampil_rekap['HARGA_TERTINGGI'] = df_tampil_rekap['HARGA_TERTINGGI'].apply(format_rupiah)
                        
                        st.dataframe(df_tampil_rekap, use_container_width=True)
                        st.info("💡 **Tips:** Tabel ini menjumlahkan total QTY dari semua Plant yang memesan barang yang sama. Sangat berguna untuk negosiasi ke vendor!")
                    except Exception as e:
                        st.warning("Gagal membuat rekap. Pastikan file Excel memiliki kolom 'QTY' dan 'HARGA' yang berisi angka.")

                if st.button("🚀 TEMBAK KE GOOGLE SHEETS Laporan PO", type="primary"):
                    try:
                        with st.spinner("Sedang mengirim (Angka murni disuntik ke Sheets)..."):
                            client = get_gspread_client()
                            sheet = client.open_by_key(SHEET_ID).get_worksheet(0) 
                            # Mengirim data ASLI (Angka) bukan data TAMPILAN (Rupiah)
                            sheet.append_rows(st.session_state['hasil_bersih_excel'].values.tolist())
                            st.success("🔥 BOOM! Semua data berhasil masuk ke Google Sheets dengan aman!")
                            del st.session_state['hasil_bersih_excel']
                    except Exception as e:
                        st.error(f"Gagal kirim: {e}")

    # --- TAB 3: CARI MANUAL ---
    with tab_cari:
        st.write("### Mesin Pencari Master Data")
        kata_cari = st.text_input("Ketik nama barang atau singkatan (contoh: knee, aki, kabel):")
        if kata_cari:
            hasil_cari = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
            data_tabel = []
            for match in hasil_cari:
                skor = round(match[1], 2)
                kunci = match[0]
                if skor >= 30:
                    baku = lookup_to_baku[kunci]
                    info = master_map.get(baku, {})
                    data_tabel.append({
                        "Skor Kemiripan": f"{skor}%", "Nama Baku di Sistem": baku,
                        "Kategori": info.get('KATEGORI', '-'), "SKU": info.get('NOMOR SKU', '-')
                    })
            if data_tabel:
                st.dataframe(pd.DataFrame(data_tabel), use_container_width=True)
            else:
                st.warning("⚠️ Tidak ada barang yang mirip di database.")

# ==========================================
# MENU 2: UPDATE MASTER DATA (INPUT)
# ==========================================
elif menu == "Update Master Data":
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
    
    if st.button("Simpan ke Master Data", type="primary"):
        st.warning("⚠️ Untuk keamanan data, saat ini penambahan master data langsung dilakukan dari Google Sheets.")

# ==========================================
# MENU 3: CARI VENDOR
# ==========================================
elif menu == "Cari Vendor":
    st.header("Database Vendor - Purchasing Regional")
    st.info("Ketik kata kunci (Nama barang, Kategori, atau Nama Vendor) untuk mencari supplier.")
    
    try:
        df_vendor = load_data(GID_VENDOR)
        df_vendor.columns = df_vendor.columns.str.strip()
        keyword = st.text_input("Cari Vendor / Barang:")
        
        if keyword:
            mask = (
                df_vendor['NAMA VENDOR'].astype(str).str.contains(keyword, case=False, na=False) |
                df_vendor['KATEGORI'].astype(str).str.contains(keyword, case=False, na=False) |
                df_vendor['ALAMAT'].astype(str).str.contains(keyword, case=False, na=False)
            )
            hasil = df_vendor[mask]
            
            if not hasil.empty:
                st.success(f"Ditemukan {len(hasil)} Vendor yang cocok!")
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
# MENU 4: FITUR MENDATANG
# ==========================================
elif menu == "Fitur Mendatang":
    st.header("Fitur Mendatang")
    st.write("Ruang ini disiapkan untuk Dashboard Grafik & Rekap PO per bulan.")