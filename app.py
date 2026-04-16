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
st.set_page_config(layout="wide", page_title="Sistem PO - Purchasing Regional")

with st.sidebar:
    st.image("logo.png", width=150) 
    st.title("Sistem Master Data")
    st.write("**Purchasing Regional**")
    
    if st.button("🔄 Sinkronisasi Data (Refresh)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
        
    st.write("---")
    
    menu = option_menu(
        menu_title="", 
        options=["Pembersihan PO", "Pencarian Barang", "Database Vendor", "Dashboard Laporan"],
        icons=["magic", "search", "shop", "bar-chart-line"], 
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
GID_DASHBOARD = "1722600044" 

def get_gspread_client():
    key_dict = json.loads(st.secrets["google_json"])
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_info(key_dict, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
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
    
    if 'KATEGORI' in df_master.columns: 
        df_master['KATEGORI'] = df_master['KATEGORI'].ffill().astype(str).str.strip().str.upper().replace('NAN', '-')
    if 'DETAIL KATEGORI' in df_master.columns: 
        df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill().astype(str).str.strip().str.upper().replace('NAN', '-')
    if 'VENDOR' in df_master.columns:
        df_master['VENDOR'] = df_master['VENDOR'].astype(str).str.strip().str.upper().replace('NAN', '-')
    
    kata_kunci = df_master.get('KATA KUNCI', df_master.get('NAMA ITEM', ""))
    df_master['KATA KUNCI'] = kata_kunci.fillna("")
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI'].astype(str)
    df_master_unique = df_master.drop_duplicates(subset=['NAMA BAKU'], keep='last')
    master_map = df_master_unique.set_index('NAMA BAKU').to_dict('index')
    
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

# ==========================================
# MENU 1: PEMBERSIHAN PO
# ==========================================
if menu == "Pembersihan PO":
    st.header("Upload & Pembersihan Laporan PO")
    st.write("Silakan pilih **Unit Kerja** lalu upload file Excel laporan (.xlsx atau .xls).")
    
    pilihan_unit = ["- Pilih Unit Kerja -", "PBI CPR", "PBI PML", "PIH", "PIH BHN PENOLONG", "RA", "PGP"]
    unit_kerja = st.selectbox("🏢 Laporan ini untuk Unit Kerja / Grup apa?", pilihan_unit)
    
    file_po = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
    
    if file_po:
        if unit_kerja == "- Pilih Unit Kerja -":
            st.warning("⚠️ Silakan pilih Unit Kerja terlebih dahulu di atas sebelum memproses file.")
        else:
            try:
                raw_excel = pd.read_excel(file_po, header=None)
                header_idx = -1
                for i, row in raw_excel.iterrows():
                    row_str = " ".join([str(val).upper() for val in row.values])
                    if 'NAMA BARANG' in row_str or 'NAMA ITEM' in row_str:
                        header_idx = i
                        break
                
                if header_idx != -1:
                    df_po = pd.read_excel(file_po, skiprows=header_idx)
                    df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
                    
                    vendor_saat_ini = "-"
                    tgl_saat_ini = "-"
                    final_data = []
                    
                    col_po_name = next((c for c in df_po.columns if 'BUKTI' in c or 'PO' in c), df_po.columns[0])
                    col_barang = next((c for c in df_po.columns if 'BARANG' in c or 'ITEM' in c), df_po.columns[1])
                    col_qty_name = next((c for c in df_po.columns if 'QTY' in c), None)
                    col_harga_name = next((c for c in df_po.columns if 'HARGA' in c), None)
                    col_tgl_name = next((c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c or c.replace('.', '').strip() in ['T', 'DATE']), None)
                    
                    for i, row in df_po.iterrows():
                        val_barang = str(row[col_barang]).strip()
                        is_barang_empty = (val_barang == '' or val_barang.lower() == 'nan' or 'UNNAMED' in val_barang.upper())
                        
                        if is_barang_empty:
                            for val in row.values:
                                v_str = str(val).strip()
                                if v_str and v_str.lower() != 'nan':
                                    v_up = v_str.upper()
                                    if not any(x in v_up for x in ["JUMLAH", "SUBTOTAL", "RP", "TOTAL", "LAPORAN", "S/D"]):
                                        if len(v_str) > 2 and not v_str.replace('.', '').replace(',', '').isdigit():
                                            vendor_saat_ini = v_str
                                            break 
                        
                        if col_tgl_name:
                            t_val = str(row[col_tgl_name]).strip()
                            if t_val and t_val.lower() != 'nan':
                                if "00:00:00" in t_val: t_val = t_val.split()[0]
                                if len(t_val) >= 4 and "JUMLAH" not in t_val.upper():
                                    tgl_saat_ini = t_val
                                    
                        po_val = str(row[col_po_name]).strip() if col_po_name else "-"
                        if po_val.lower() == 'nan' or not po_val: po_val = "-"
                        
                        if not is_barang_empty and "JUMLAH" not in val_barang.upper() and "SUBTOTAL" not in val_barang.upper() and val_barang.upper() != "RP":
                            qty_val = row[col_qty_name] if col_qty_name else 0
                            harga_val = row[col_harga_name] if col_harga_name else 0
                            
                            final_data.append({
                                "UNIT KERJA": unit_kerja, "NO PO": po_val,
                                "TANGGAL": tgl_saat_ini, "VENDOR": vendor_saat_ini,
                                "ITEM_KOTOR": val_barang, "QTY": qty_val, "HARGA": harga_val
                            })
                    
                    df_clean = pd.DataFrame(final_data)
                    st.success(f"🤖 Berhasil memproses {len(df_clean)} baris barang untuk Plant {unit_kerja}.")
                    
                    if st.button("Proses Standardisasi Data", type="primary", use_container_width=True):
                        hasil_rows = []
                        for _, r in df_clean.iterrows():
                            match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                            if match and match[1] >= 70:
                                baku = lookup_to_baku[match[0]]; info = master_map.get(baku, {})
                                hasil_rows.append({
                                    "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'],
                                    "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                    "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, 
                                    "QTY": r['QTY'], "SATUAN": info.get('SATUAN', '-'), "HARGA": r['HARGA'], 
                                    "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), 
                                    "SKU": info.get('NOMOR SKU', '-')
                                })
                            else:
                                hasil_rows.append({
                                    "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'],
                                    "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                    "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ CEK MANUAL", 
                                    "QTY": r['QTY'], "SATUAN": "-", "HARGA": r['HARGA'], 
                                    "KATEGORI": "-", "DETAIL KATEGORI": "-", 
                                    "SKU": "-"
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
            if st.button("🚀 TEMBAK DETAIL KE GOOGLE SHEETS", key="btn_dtl", type="primary"):
                try:
                    with st.spinner("Mengirim..."):
                        client = get_gspread_client(); sheet = client.open_by_key(SHEET_ID).worksheets()[-1]
                        sheet.append_rows(df_res.fillna("").values.tolist())
                        st.success(f"🔥 Berhasil dikirim ke tab '{sheet.title}'!"); del st.session_state['hasil_bersih']
                except Exception as e: st.error(f"Gagal: {e}")
        
        with t2:
            st.write("Cek Detail per PO untuk mengirim data ke database.")

# ==========================================
# MENU 2: PENCARIAN BARANG
# ==========================================
elif menu == "Pencarian Barang":
    st.header("🔍 Kamus & Histori Barang")
    st.write("Ketik nama barang acak dari lapangan. Sistem akan menampilkan standar nama, riwayat harga, dan vendor terakhir.")
    kata_cari = st.text_input("Ketik Nama Barang / Singkatan (Misal: Accu, Timbangan, Besi):")
    if kata_cari:
        hasil_cari = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
        data_tabel = []
        for match in hasil_cari:
            skor = round(match[1], 1)
            if skor >= 40: 
                kunci = match[0]
                baku = lookup_to_baku[kunci]
                info = master_map.get(baku, {}) 
                data_tabel.append({
                    "Akurasi": f"{skor}%", "Nama Baku (Standar)": baku, "SKU": info.get('NOMOR SKU', '-'),
                    "Kategori": info.get('KATEGORI', '-'), "Detail": info.get('DETAIL KATEGORI', '-'),
                    "Satuan": info.get('SATUAN', '-'), "Harga Terakhir": str(info.get('HARGA', '-')),
                    "Vendor Terakhir": info.get('VENDOR', '-'), "Tgl Pembelian": str(info.get('TANGGAL', '-'))
                })
        if data_tabel:
            st.dataframe(pd.DataFrame(data_tabel), use_container_width=True)
        else:
            st.warning("⚠️ Tidak ada barang yang mirip di database.")

# ==========================================
# MENU 3: DATABASE VENDOR
# ==========================================
elif menu == "Database Vendor":
    st.header("Database Pencarian Vendor")
    keyword = st.text_input("Cari Supplier:", placeholder="Misal: Besi, Kimia, atau nama PT...")
    if keyword:
        try:
            df_v = load_data(GID_VENDOR)
            df_v.columns = df_v.columns.str.strip().str.upper()
            res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
            if not res.empty:
                for _, v in res.iterrows():
                    with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - {v.get('KATEGORI', '-')} (PIC: {v.get('PIC', '-')})"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**📍 Alamat:** {v.get('ALAMAT', '-')}")
                            st.write(f"**👤 PIC:** {v.get('PIC', '-')}")
                        with col2:
                            st.write(f"**📞 Kontak:** {v.get('KONTAK', '-')}")
                            st.write(f"**⏳ TOP:** {v.get('TOP', '-')}")
            else: st.warning("Maaf, vendor tidak ditemukan.")
        except Exception: st.error("Gagal Load Database Vendor.")

# ==========================================
# MENU 4: DASHBOARD LAPORAN (FORMAT PPT MANAJEMEN)
# ==========================================
elif menu == "Dashboard Laporan":
    st.header("📊 Executive Dashboard Purchasing")
    st.write("Laporan Rekapitulasi Pembelian & Frekuensi PO (Berdasarkan Sheet 4)")
    
    try:
        df_dash = load_data(GID_DASHBOARD)
        df_dash.columns = df_dash.columns.str.strip().str.upper()
        
        # Pastikan kolom utama tersedia
        if not df_dash.empty and 'NO PO' in df_dash.columns and 'UNIT KERJA' in df_dash.columns:
            
            # --- PEMBERSIHAN DATA ---
            harga_str = df_dash['HARGA'].astype(str).str.upper().str.replace('RP', '', regex=False)
            harga_str = harga_str.str.split(',').str[0].str.replace(r'[^0-9]', '', regex=True)
            df_dash['HARGA_NUM'] = pd.to_numeric(harga_str, errors='coerce').fillna(0)
            df_dash['QTY_NUM'] = pd.to_numeric(df_dash['QTY'], errors='coerce').fillna(0)
            df_dash['TOTAL_NILAI'] = df_dash['QTY_NUM'] * df_dash['HARGA_NUM']
            
            df_dash['TANGGAL_PARSED'] = pd.to_datetime(df_dash['TANGGAL'], errors='coerce')
            df_dash['BULAN'] = df_dash['TANGGAL_PARSED'].dt.strftime('%B %Y').fillna('Lainnya')
            
            # --- 3 KARTU KPI UTAMA ---
            total_pembelian = df_dash['TOTAL_NILAI'].sum()
            total_po = df_dash['NO PO'].nunique()
            total_item = len(df_dash)
            
            col1, col2, col3 = st.columns(3)
            col1.info(f"💰 **Total Pembelian:** {format_rupiah(total_pembelian)}")
            col2.success(f"📄 **Total Dokumen PO:** {total_po} PO")
            col3.warning(f"📦 **Total Baris Item:** {total_item} Item")
            
            st.write("---")
            
            # --- BARIS 1: TABEL BULAN & TOP PO ---
            c1, c2 = st.columns(2)
            with c1:
                st.write("#### 📅 Rekapitulasi per Bulan")
                rekap_bulan = df_dash.groupby('BULAN').agg(
                    Jumlah_PO=('NO PO', 'nunique'),
                    Total_Harga=('TOTAL_NILAI', 'sum')
                ).reset_index()
                rekap_bulan['Total_Harga'] = rekap_bulan['Total_Harga'].apply(format_rupiah)
                st.dataframe(rekap_bulan, use_container_width=True)
                
            with c2:
                st.write("#### 🏆 Top 10 Item (Berdasarkan Total PO)")
                top_po = df_dash[~df_dash['NAMA BAKU'].str.contains('CEK MANUAL', na=False)].groupby('NAMA BAKU').agg(
                    TOTAL_PO=('NO PO', 'nunique')
                ).reset_index().sort_values('TOTAL_PO', ascending=False).head(10)
                st.dataframe(top_po, use_container_width=True)
                
            st.write("---")
            
            # --- BARIS 2: TABEL UNIT KERJA & TOP QTY ---
            c3, c4 = st.columns(2)
            with c3:
                st.write("#### 🏢 Pembelian & Frekuensi per Unit Kerja")
                rekap_unit = df_dash.groupby('UNIT KERJA').agg(
                    Total_Pembelian=('TOTAL_NILAI', 'sum'),
                    Jumlah_PO=('NO PO', 'nunique')
                ).reset_index().sort_values('Total_Pembelian', ascending=False)
                
                jml_bulan = df_dash['BULAN'].nunique()
                jml_bulan = jml_bulan if jml_bulan > 0 else 1
                rekap_unit['Rata-Rata PO/Bln'] = (rekap_unit['Jumlah_PO'] / jml_bulan).astype(int)
                rekap_unit['Total_Pembelian'] = rekap_unit['Total_Pembelian'].apply(format_rupiah)
                st.dataframe(rekap_unit, use_container_width=True)
                    
            with c4:
                st.write("#### 📈 Top 10 Item (Berdasarkan Kuantitas)")
                top_qty = df_dash[~df_dash['NAMA BAKU'].str.contains('CEK MANUAL', na=False)].groupby('NAMA BAKU').agg(
                    TOTAL_QTY=('QTY_NUM', 'sum')
                ).reset_index().sort_values('TOTAL_QTY', ascending=False).head(10)
                
                # Tambahkan kolom Satuan agar sama persis seperti PPT
                satuan_dict = df_dash.drop_duplicates('NAMA BAKU').set_index('NAMA BAKU')['SATUAN'].to_dict()
                top_qty['SATUAN'] = top_qty['NAMA BAKU'].map(satuan_dict)
                
                st.dataframe(top_qty, use_container_width=True)
                
        else:
            st.warning("⚠️ Data Transaksi Sheet 4 masih kosong atau Judul Kolom salah. Pastikan sudah mengubah Header di Sheet 4 sesuai panduan!")
    except Exception as e:
        st.error(f"Gagal memuat Dashboard: {e}")