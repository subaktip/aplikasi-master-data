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
    
    # [UPDATE]: Menu disederhanakan hanya menjadi 2 opsi utama
    menu = option_menu(
        menu_title="", 
        options=["Pembersihan PO", "Database Vendor"],
        icons=["magic", "search"], 
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
    master_map = df_master.drop_duplicates(subset=['NAMA BAKU']).set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

# ==========================================
# MENU 1: PEMBERSIHAN PO (UPLOAD EXCEL ONLY)
# ==========================================
if menu == "Pembersihan PO":
    st.header("Upload & Pembersihan Laporan PO")
    st.write("Silakan upload file Excel laporan (.xlsx atau .xls). Sistem akan otomatis merapikan format dan menstandarisasi nama barang.")
    
    file_po = st.file_uploader("", type=["xlsx", "xls"])
    
    if file_po:
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
                
                col_barang = next((c for c in df_po.columns if 'BARANG' in c or 'ITEM' in c), df_po.columns[1])
                col_qty_name = next((c for c in df_po.columns if 'QTY' in c), None)
                col_harga_name = next((c for c in df_po.columns if 'HARGA' in c), None)
                col_tgl_name = next((c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c or c.replace('.', '').strip() in ['T', 'DATE']), None)
                
                for i, row in df_po.iterrows():
                    val_barang = str(row[col_barang]).strip()
                    is_barang_empty = (val_barang == '' or val_barang.lower() == 'nan' or 'UNNAMED' in val_barang.upper())
                    
                    # Sapu Baris untuk Vendor
                    if is_barang_empty:
                        for val in row.values:
                            v_str = str(val).strip()
                            if v_str and v_str.lower() != 'nan':
                                v_up = v_str.upper()
                                if not any(x in v_up for x in ["JUMLAH", "SUBTOTAL", "RP", "TOTAL", "LAPORAN", "S/D"]):
                                    if len(v_str) > 2 and not v_str.replace('.', '').replace(',', '').isdigit():
                                        vendor_saat_ini = v_str
                                        break 
                    
                    # Tangkap Tanggal
                    if col_tgl_name:
                        t_val = str(row[col_tgl_name]).strip()
                        if t_val and t_val.lower() != 'nan':
                            if "00:00:00" in t_val: t_val = t_val.split()[0]
                            if len(t_val) >= 4 and "JUMLAH" not in t_val.upper():
                                tgl_saat_ini = t_val
                    
                    # Simpan Item
                    if not is_barang_empty and "JUMLAH" not in val_barang.upper() and "SUBTOTAL" not in val_barang.upper() and val_barang.upper() != "RP":
                        qty_val = row[col_qty_name] if col_qty_name else 0
                        harga_val = row[col_harga_name] if col_harga_name else 0
                        
                        final_data.append({
                            "TANGGAL": tgl_saat_ini, "VENDOR": vendor_saat_ini,
                            "ITEM_KOTOR": val_barang, "QTY": qty_val, "HARGA": harga_val
                        })
                
                df_clean = pd.DataFrame(final_data)
                st.success(f"🤖 Berhasil memproses {len(df_clean)} baris barang.")
                
                if st.button("Proses Standardisasi Data", type="primary", use_container_width=True):
                    hasil_rows = []
                    for _, r in df_clean.iterrows():
                        match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                        if match and match[1] >= 70:
                            baku = lookup_to_baku[match[0]]; info = master_map.get(baku, {})
                            hasil_rows.append({
                                "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, 
                                "QTY": r['QTY'], "SATUAN": info.get('SATUAN', '-'), "HARGA": r['HARGA'], 
                                "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), 
                                "NOMOR SKU": info.get('NOMOR SKU', '-')
                            })
                        else:
                            hasil_rows.append({
                                "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ CEK MANUAL", 
                                "QTY": r['QTY'], "SATUAN": "-", "HARGA": r['HARGA'], 
                                "KATEGORI": "-", "DETAIL KATEGORI": "-", 
                                "NOMOR SKU": "-"
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
                
                if st.button("🚀 TEMBAK REKAP KE GOOGLE SHEETS", key="btn_rkp", type="primary"):
                    try:
                        with st.spinner("Mengirim Rekap..."):
                            client = get_gspread_client(); sheet = client.open_by_key(SHEET_ID).worksheets()[-1]
                            sheet.append_rows([df_g_murni.columns.tolist()] + df_g_murni.fillna("").values.tolist())
                            st.success(f"🔥 Rekap berhasil dikirim ke tab '{sheet.title}'!"); del st.session_state['hasil_bersih']
                    except Exception as e: st.error(f"Gagal: {e}")
            except Exception as e:
                st.warning("Gagal memproses rekap. Pastikan ada angka di QTY dan Harga.")

# ==========================================
# MENU 2: DATABASE VENDOR
# ==========================================
elif menu == "Database Vendor":
    st.header("Database Pencarian Vendor")
    st.write("Ketik nama vendor, nama PIC, alamat, atau jenis barang yang dicari.")
    
    keyword = st.text_input("Cari Supplier:", placeholder="Misal: Besi, Kimia, atau nama PT...")
    
    if keyword:
        try:
            df_v = load_data(GID_VENDOR)
            df_v.columns = df_v.columns.str.strip().str.upper()
            res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
            
            if not res.empty:
                st.success(f"Ditemukan {len(res)} vendor yang cocok.")
                for _, v in res.iterrows():
                    with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - {v.get('KATEGORI', '-')} (PIC: {v.get('PIC', '-')})"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.write(f"**📍 Alamat:** {v.get('ALAMAT', '-')}")
                            st.write(f"**👤 PIC:** {v.get('PIC', '-')}")
                            st.write(f"**📞 Kontak:** {v.get('KONTAK', '-')}")
                            st.write(f"**📧 Email:** {v.get('EMAIL', '-')}")
                        with col2:
                            st.write(f"**💳 Rekening:** {v.get('REKENING', '-')}")
                            st.write(f"**🏦 A/N Rekening:** {v.get('ATAS NAMA REKENING', '-')}")
                            st.write(f"**⏳ TOP:** {v.get('TOP', '-')}")
                            st.write(f"**🪪 NPWP:** {v.get('NPWP', '-')}")
            else: 
                st.warning("Maaf, vendor dengan kata kunci tersebut tidak ditemukan di database.")
        except Exception as e: 
            st.error("Gagal Load Database Vendor. Pastikan tab vendor ada dan GID benar.")