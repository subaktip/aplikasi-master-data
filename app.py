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
# 1. KONFIGURASI HALAMAN & SIDEBAR
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
            "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#e2e6ea", "border-radius": "8px"},
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

def format_rupiah(angka):
    try: return f"Rp {int(angka):,}".replace(',', '.')
    except: return "Rp 0"

# --- PERSIAPAN KAMUS PINTAR ---
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    df_master = df_master.dropna(subset=['NAMA BAKU'])
    df_master = df_master[df_master['NAMA BAKU'].astype(str).str.strip().str.lower() != "(blank)"]
    if 'KATEGORI' in df_master.columns: df_master['KATEGORI'] = df_master['KATEGORI'].ffill()
    if 'DETAIL KATEGORI' in df_master.columns: df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill()
    df_master['KATA KUNCI'] = df_master.get('KATA KUNCI', "").fillna("")
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    master_map = df_master.drop_duplicates(subset=['NAMA BAKU']).set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

# ==========================================
# MENU 1: PEMBERSIHAN NAMA BAKU
# ==========================================
if menu == "Pembersihan Nama":
    st.header("Pembersihan Master Data PO")
    tab_copy, tab_excel, tab_cari = st.tabs(["Copy-Paste", "Upload Excel", "Cari Manual"])
    
    with tab_copy:
        st.write("### Mode Cepat: Copy-Paste Teks")
        teks_po = st.text_area("Paste daftar nama barang kotor di sini:", height=150)
        if st.button("Proses Teks", type="primary"):
            if teks_po.strip():
                daftar_item = [item.strip() for item in teks_po.split('\n') if item.strip()]
                hasil_teks = []
                for item_kotor in daftar_item:
                    matches = process.extract(item_kotor, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
                    ditemukan = False
                    for m in matches:
                        if m[1] >= 40:
                            baku = lookup_to_baku[m[0]]; info = master_map.get(baku, {})
                            if not any(d.get('Nama Baku (Sistem)') == baku and d.get('Input User') == item_kotor for d in hasil_teks):
                                hasil_teks.append({"Input User": item_kotor, "Nama Baku (Sistem)": baku, "Kategori": info.get('KATEGORI', '-'), "Detail Kategori": info.get('DETAIL KATEGORI', '-'), "Akurasi (%)": round(m[1], 1)})
                                ditemukan = True
                    if not ditemukan:
                        hasil_teks.append({"Input User": item_kotor, "Nama Baku (Sistem)": "⚠️ Tidak Ditemukan", "Kategori": "-", "Detail Kategori": "-", "Akurasi (%)": 0})
                
                df_hasil = pd.DataFrame(hasil_teks)
                if not df_hasil.empty:
                    df_hasil['Akurasi (%)'] = pd.to_numeric(df_hasil['Akurasi (%)'], errors='coerce').fillna(0)
                    df_hasil = df_hasil.sort_values(by=['Input User', 'Akurasi (%)'], ascending=[True, False]).reset_index(drop=True)
                st.dataframe(df_hasil, use_container_width=True)

    with tab_excel:
        st.write("### Mode Lengkap: Auto-Detect Format Laporan")
        file_po = st.file_uploader("Upload Excel (.xlsx/.xls)", type=["xlsx", "xls"])
        
        if file_po:
            try:
                # 1. Deteksi Header Otomatis
                raw_excel = pd.read_excel(file_po, header=None)
                header_idx = -1
                for i, row in raw_excel.iterrows():
                    # PERBAIKAN BUG DI SINI (Posisi kurung diperbaiki)
                    row_str = " ".join(row.astype(str)).upper()
                    if 'NAMA BARANG' in row_str or 'NAMA ITEM' in row_str:
                        header_idx = i
                        break
                
                if header_idx != -1:
                    df_po = pd.read_excel(file_po, skiprows=header_idx)
                    df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
                    
                    # 2. Algoritma 'Context Tracker' (Vendor & Tanggal)
                    vendor_saat_ini = "-"
                    tgl_saat_ini = "-"
                    final_data = []
                    
                    # Cari tau kolom krusial
                    col_barang = [c for c in df_po.columns if 'NAMA BARANG' in c or 'NAMA ITEM' in c][0]
                    col_qty = [c for c in df_po.columns if 'QTY' in c]
                    col_qty_name = col_qty[0] if col_qty else None
                    col_harga = [c for c in df_po.columns if 'HARGA' in c]
                    col_harga_name = col_harga[0] if col_harga else None
                    col_tgl = [c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c]
                    col_ref = df_po.columns[0] # Kolom pertama biasanya buat vendor/no bukti
                    
                    for i, row in df_po.iterrows():
                        val_ref = str(row[col_ref]) if not pd.isna(row[col_ref]) else ""
                        val_barang = str(row[col_barang]) if not pd.isna(row[col_barang]) else ""
                        
                        # Logika Tangkap Vendor
                        if (not val_barang or val_barang.lower() == 'nan') and val_ref and val_ref.lower() != 'nan':
                            if "JUMLAH" not in val_ref.upper() and "SUBTOTAL" not in val_ref.upper(): 
                                vendor_saat_ini = val_ref
                        
                        # Logika Tangkap Tanggal
                        if col_tgl and not pd.isna(row[col_tgl[0]]) and str(row[col_tgl[0]]).lower() != 'nan':
                            tgl_val = str(row[col_tgl[0]])
                            if len(tgl_val) > 4: 
                                tgl_saat_ini = tgl_val
                        
                        # Logika Simpan Barang
                        if val_barang and val_barang.lower() != 'nan' and "JUMLAH" not in val_barang.upper() and "SUBTOTAL" not in val_barang.upper():
                            qty_val = row[col_qty_name] if col_qty_name else 0
                            harga_val = row[col_harga_name] if col_harga_name else 0
                            
                            final_data.append({
                                "ITEM_KOTOR": val_barang, "QTY": qty_val, "HARGA": harga_val,
                                "VENDOR": vendor_saat_ini, "TANGGAL": tgl_saat_ini
                            })
                    
                    df_clean = pd.DataFrame(final_data)
                    st.success(f"🤖 Berhasil memproses {len(df_clean)} baris barang dengan metode 'Context Tracker'!")
                    
                    if st.button("Proses Pembersihan & Rekap", type="primary"):
                        hasil_rows = []
                        for _, r in df_clean.iterrows():
                            match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                            if match and match[1] >= 70:
                                baku = lookup_to_baku[match[0]]; info = master_map.get(baku, {})
                                hasil_rows.append({
                                    "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, "KATEGORI": info.get('KATEGORI', '-'), 
                                    "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), "NOMOR SKU": info.get('NOMOR SKU', '-'),
                                    "SATUAN": info.get('SATUAN', '-'), "HARGA": r['HARGA'], "QTY": r['QTY'], 
                                    "VENDOR": r['VENDOR'], "GRUP": "-", "TANGGAL": r['TANGGAL']
                                })
                            else:
                                hasil_rows.append({
                                    "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ CEK MANUAL", "KATEGORI": "-", 
                                    "DETAIL KATEGORI": "-", "NOMOR SKU": "-", "SATUAN": "-", 
                                    "HARGA": r['HARGA'], "QTY": r['QTY'], "VENDOR": r['VENDOR'], 
                                    "GRUP": "-", "TANGGAL": r['TANGGAL']
                                })
                        st.session_state['hasil_bersih'] = pd.DataFrame(hasil_rows)
                else:
                    st.error("Gagal menemukan kolom 'Nama Barang'. Pastikan file Excel tidak kosong.")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file: {e}")

        if 'hasil_bersih' in st.session_state:
            df_res = st.session_state['hasil_bersih']
            t1, t2 = st.tabs(["📄 Detail per PO", "📊 Rekap Bulanan"])
            
            with t1:
                df_v = df_res.copy(); df_v['HARGA'] = df_v['HARGA'].apply(format_rupiah)
                st.dataframe(df_v, use_container_width=True)
                if st.button("🚀 TEMBAK DETAIL KE SHEETS", key="btn_dtl", type="primary"):
                    try:
                        with st.spinner("Mengirim..."):
                            client = get_gspread_client(); sheet = client.open_by_key(SHEET_ID).worksheets()[-1]
                            sheet.append_rows(df_res.fillna("").values.tolist())
                            st.success(f"🔥 Selesai mendarat di tab '{sheet.title}'!"); del st.session_state['hasil_bersih']
                    except Exception as e: st.error(f"Gagal: {e}")
            
            with t2:
                try:
                    df_rk = df_res.copy()
                    df_rk['QTY'] = pd.to_numeric(df_rk['QTY'], errors='coerce').fillna(0)
                    df_rk['HARGA'] = pd.to_numeric(df_rk['HARGA'], errors='coerce').fillna(0)
                    df_rk['TOTAL_NILAI'] = df_rk['QTY'] * df_rk['HARGA']
                    
                    df_g = df_rk.groupby(['NAMA BAKU', 'KATEGORI', 'SATUAN']).agg(
                        TOTAL_QTY=('QTY', 'sum'), TOTAL_BELANJA=('TOTAL_NILAI', 'sum'),
                        HARGA_TERTINGGI=('HARGA', 'max'), FREKUENSI_ORDER=('NAMA BAKU', 'count')
                    ).reset_index()
                    df_g['HARGA_RATA_RATA'] = (df_g['TOTAL_BELANJA'] / df_g['TOTAL_QTY']).fillna(0).round(0)
                    
                    kolom_tampil = ['NAMA BAKU', 'KATEGORI', 'TOTAL_QTY', 'SATUAN', 'HARGA_RATA_RATA', 'HARGA_TERTINGGI', 'FREKUENSI_ORDER']
                    df_g_murni = df_g[kolom_tampil].copy()
                    
                    df_gv = df_g[kolom_tampil].copy()
                    df_gv['HARGA_RATA_RATA'] = df_gv['HARGA_RATA_RATA'].apply(format_rupiah)
                    df_gv['HARGA_TERTINGGI'] = df_gv['HARGA_TERTINGGI'].apply(format_rupiah)
                    
                    st.dataframe(df_gv, use_container_width=True)
                    
                    if st.button("🚀 TEMBAK REKAP KE SHEETS", key="btn_rkp", type="primary"):
                        try:
                            with st.spinner("Mengirim Rekap..."):
                                client = get_gspread_client(); sheet = client.open_by_key(SHEET_ID).worksheets()[-1]
                                sheet.append_rows([df_g_murni.columns.tolist()] + df_g_murni.fillna("").values.tolist())
                                st.success(f"🔥 Rekap berhasil mendarat di tab '{sheet.title}'!"); del st.session_state['hasil_bersih']
                        except Exception as e: st.error(f"Gagal: {e}")
                except Exception as e:
                    st.warning(f"Gagal memproses rekap. Pastikan ada angka di QTY dan Harga. Error: {e}")

elif menu == "Cari Vendor":
    st.header("Database Vendor")
    keyword = st.text_input("Cari Vendor / Barang:")
    if keyword:
        try:
            df_v = load_data(GID_VENDOR)
            df_v.columns = df_v.columns.str.strip().str.upper()
            res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
            if not res.empty:
                for _, v in res.iterrows():
                    with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - {v.get('KATEGORI', '-')}"):
                        st.write(f"**PIC:** {v.get('PIC', '-')} | **Kontak:** {v.get('KONTAK', '-')}")
            else: st.warning("Vendor tidak ditemukan.")
        except Exception as e: st.error("Gagal Load Vendor")

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

elif menu == "Fitur Mendatang":
    st.header("Fitur Mendatang")
    st.write("Ruang ini disiapkan untuk Dashboard Grafik & Rekap PO per bulan.")