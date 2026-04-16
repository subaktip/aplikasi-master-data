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

def extract_code(text):
    try: return text.split('(')[1].split(')')[0].strip()
    except: return "000"

# --- PERSIAPAN KAMUS PINTAR ---
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    df_master = df_master.dropna(subset=['NAMA BAKU'])
    
    if 'KATEGORI' in df_master.columns: df_master['KATEGORI'] = df_master['KATEGORI'].ffill().astype(str).str.strip().str.upper()
    if 'DETAIL KATEGORI' in df_master.columns: df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill().astype(str).str.strip().str.upper()
    
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master.get('KATA KUNCI', "").fillna("").astype(str)
    master_map = df_master.drop_duplicates(subset=['NAMA BAKU'], keep='last').set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

# --- FUNGSI AUTO SKU GENERATOR ---
def generate_new_sku(kat_full, det_full):
    c_kat = extract_code(kat_full)
    c_det = extract_code(det_full)
    prefix = "001"
    
    pattern = f"{prefix}-{c_kat}-{c_det}-"
    df_match = df_master[df_master['NOMOR SKU'].astype(str).str.contains(pattern, na=False)]
    
    if not df_match.empty:
        # Ambil 3 angka terakhir
        last_nums = []
        for s in df_match['NOMOR SKU'].astype(str):
            try: last_nums.append(int(s.split('-')[-1]))
            except: pass
        next_val = max(last_nums) + 1 if last_nums else 1
    else:
        next_val = 1
        
    return f"{prefix}-{c_kat}-{c_det}-{next_val:03d}"

# ==========================================
# MENU 1: PEMBERSIHAN PO + AUTO SKU COMBO
# ==========================================
if menu == "Pembersihan PO":
    st.header("Upload & Pembersihan Laporan PO")
    
    pilihan_unit = ["- Pilih Unit Kerja -", "PBI CPR", "PBI PML", "PIH", "PIH BHN PENOLONG", "RA", "PGP"]
    unit_kerja = st.selectbox("🏢 Pilih Unit Kerja:", pilihan_unit)
    
    file_po = st.file_uploader("Upload Excel Laporan (.xlsx/.xls)", type=["xlsx", "xls"])
    
    if file_po and unit_kerja != "- Pilih Unit Kerja -":
        try:
            # PROSES BACA EXCEL CERDAS
            raw_excel = pd.read_excel(file_po, header=None)
            header_idx = -1
            for i, row in raw_excel.iterrows():
                row_str = " ".join([str(val).upper() for val in row.values])
                if 'NAMA BARANG' in row_str or 'NAMA ITEM' in row_str:
                    header_idx = i; break
            
            if header_idx != -1:
                df_po = pd.read_excel(file_po, skiprows=header_idx)
                df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
                
                vendor_saat_ini, tgl_saat_ini, final_data = "-", "-", []
                col_po = next((c for c in df_po.columns if 'BUKTI' in c or 'PO' in c), df_po.columns[0])
                col_barang = next((c for c in df_po.columns if 'BARANG' in c or 'ITEM' in c), df_po.columns[1])
                col_qty = next((c for c in df_po.columns if 'QTY' in c), None)
                col_harga = next((c for c in df_po.columns if 'HARGA' in c), None)
                col_tgl = next((c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c or c.replace('.', '').strip() in ['T', 'DATE']), None)
                
                for i, row in df_po.iterrows():
                    val_barang = str(row[col_barang]).strip()
                    is_empty = (val_barang == '' or val_barang.lower() == 'nan' or 'UNNAMED' in val_barang.upper())
                    
                    if is_empty:
                        for val in row.values:
                            v_str = str(val).strip()
                            if v_str and v_str.lower() != 'nan' and not any(x in v_str.upper() for x in ["JUMLAH", "RP", "TOTAL"]):
                                if len(v_str) > 2 and not v_str.replace('.', '').replace(',', '').isdigit():
                                    vendor_saat_ini = v_str; break 
                    
                    if col_tgl and not pd.isna(row[col_tgl]):
                        t_val = str(row[col_tgl]).strip()
                        if len(t_val) >= 4 and "JUMLAH" not in t_val.upper(): tgl_saat_ini = t_val.split()[0]
                    
                    po_val = str(row[col_po]).strip() if col_po else "-"
                    
                    if not is_empty and "JUMLAH" not in val_barang.upper() and val_barang.upper() != "RP":
                        final_data.append({
                            "UNIT KERJA": unit_kerja, "NO PO": po_val, "TANGGAL": tgl_saat_ini, 
                            "VENDOR": vendor_saat_ini, "ITEM_KOTOR": val_barang, 
                            "QTY": row[col_qty] if col_qty else 0, "HARGA": row[col_harga] if col_harga else 0
                        })
                
                df_clean = pd.DataFrame(final_data)
                
                if st.button("🚀 Proses Pembersihan & Sinkronisasi SKU", type="primary", use_container_width=True):
                    hasil_rows = []
                    for _, r in df_clean.iterrows():
                        match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                        if match and match[1] >= 75:
                            baku = lookup_to_baku[match[0]]; info = master_map.get(baku, {})
                            hasil_rows.append({
                                "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, "QTY": r['QTY'], "SATUAN": info.get('SATUAN', '-'),
                                "HARGA": r['HARGA'], "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), "SKU": info.get('NOMOR SKU', '-')
                            })
                        else:
                            hasil_rows.append({
                                "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'],
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ BARANG BARU", "QTY": r['QTY'], "SATUAN": "-",
                                "HARGA": r['HARGA'], "KATEGORI": "-", "DETAIL KATEGORI": "-", "SKU": "-"
                            })
                    st.session_state['hasil_po'] = pd.DataFrame(hasil_rows)
                    st.rerun()

        except Exception as e: st.error(f"Error: {e}")

    # TAMPILAN TAB HASIL
    if 'hasil_po' in st.session_state:
        t1, t2, t3 = st.tabs(["📄 Hasil Pembersihan", "🆕 Registrasi SKU Otomatis", "📊 Rekap"])
        
        with t1:
            st.dataframe(st.session_state['hasil_po'], use_container_width=True)
            if st.button("💾 Kirim Data Bersih ke Sheet 4"):
                try:
                    with st.spinner("Mengirim ke database..."):
                        client = get_gspread_client()
                        sheet_4 = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD))
                        sheet_4.append_rows(st.session_state['hasil_po'].fillna("").values.tolist())
                        st.success("Berhasil dikirim ke Sheet 4!")
                        del st.session_state['hasil_po']; st.rerun()
                except Exception as e: st.error(e)

        with t2:
            st.write("### 🤖 Asisten Registrasi Barang Baru")
            df_curr = st.session_state['hasil_po']
            new_items = df_curr[df_curr['NAMA BAKU'] == "⚠️ BARANG BARU"]['NAMA ITEM'].unique()
            
            if len(new_items) > 0:
                item_select = st.selectbox("Pilih barang yang ingin didaftarkan SKU-nya:", new_items)
                
                c_a, c_b = st.columns(2)
                with c_a:
                    kat_list = sorted(df_master['KATEGORI'].unique())
                    kat_sel = st.selectbox("Kategori:", kat_list)
                with c_b:
                    det_list = sorted(df_master[df_master['KATEGORI'] == kat_sel]['DETAIL KATEGORI'].unique())
                    det_sel = st.selectbox("Detail Kategori:", det_list)
                
                sku_baru = generate_new_sku(kat_sel, det_sel)
                st.info(f"**Saran SKU Baru:** `{sku_baru}`")
                
                if st.button("🔥 Daftarkan & Update PO", type="primary"):
                    try:
                        client = get_gspread_client()
                        sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                        
                        # Injek ke Master Data (Sheet 1)
                        sheet_master.append_row([item_select, item_select, "", kat_sel, det_sel, sku_baru, "PCS"])
                        
                        # Update tabel PO yang sedang aktif di session
                        st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'NAMA BAKU'] = item_select
                        st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'SKU'] = sku_baru
                        st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'KATEGORI'] = kat_sel
                        st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'DETAIL KATEGORI'] = det_sel
                        
                        st.success(f"Barang {item_select} resmi terdaftar dengan SKU {sku_baru}!")
                        time.sleep(1); st.rerun()
                    except Exception as e: st.error(e)
            else:
                st.success("Semua barang di laporan ini sudah terdaftar. Mantap!")

        with t3:
            st.write("Lakukan pengiriman data di Tab 1 untuk melihat rekapitulasi.")

# ==========================================
# MENU LAINNYA (DIPERTAHANKAN)
# ==========================================
elif menu == "Pencarian Barang":
    st.header("🔍 Kamus & Histori Barang")
    kata_cari = st.text_input("Ketik Nama Barang / SKU:")
    if kata_cari:
        hasil = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
        res_list = []
        for m in hasil:
            if m[1] >= 40:
                baku = lookup_to_baku[m[0]]; info = master_map.get(baku, {})
                res_list.append({
                    "Akurasi": f"{m[1]}%", "Nama Baku": baku, "SKU": info.get('NOMOR SKU', '-'),
                    "Kategori": info.get('KATEGORI', '-'), "Harga Terakhir": info.get('HARGA', '-'),
                    "Vendor Terakhir": info.get('VENDOR', '-')
                })
        st.dataframe(pd.DataFrame(res_list), use_container_width=True)

elif menu == "Database Vendor":
    st.header("Database Pencarian Vendor")
    keyword = st.text_input("Cari Supplier:")
    if keyword:
        df_v = load_data(GID_VENDOR)
        df_v.columns = df_v.columns.str.strip().str.upper()
        res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
        for _, v in res.iterrows():
            with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - {v.get('KATEGORI', '-')}"):
                st.write(f"**PIC:** {v.get('PIC', '-')} | **Kontak:** {v.get('KONTAK', '-')}")
                st.write(f"**Alamat:** {v.get('ALAMAT', '-')}")

elif menu == "Dashboard Laporan":
    st.header("📊 Executive Dashboard Purchasing")
    try:
        client = get_gspread_client()
        sheet_dash = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD))
        data_dash = sheet_dash.get_all_values()
        
        if len(data_dash) > 1:
            df_d = pd.DataFrame(data_dash[1:], columns=data_dash[0])
            df_d.columns = df_d.columns.str.strip().str.upper()
            
            # Deteksi Kolom Dinamis
            c_po = next((c for c in df_d.columns if 'PO' in c or 'BUKTI' in c), None)
            c_unit = next((c for c in df_d.columns if 'UNIT' in c or 'GRUP' in c), None)
            c_harga = next((c for c in df_d.columns if 'HARGA' in c), None)
            
            if c_po and c_unit and c_harga:
                h_str = df_d[c_harga].astype(str).str.upper().str.replace('RP', '', regex=False).str.split(',').str[0].str.replace(r'[^0-9]', '', regex=True)
                df_d['H_NUM'] = pd.to_numeric(h_str, errors='coerce').fillna(0)
                df_d['Q_NUM'] = pd.to_numeric(df_d['QTY'], errors='coerce').fillna(0)
                df_d['TOTAL'] = df_d['H_NUM'] * df_d['Q_NUM']
                
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Belanja", format_rupiah(df_d['TOTAL'].sum()))
                col2.metric("📄 Total PO", df_d[c_po].nunique())
                col3.metric("🏢 Unit Aktif", df_d[c_unit].nunique())
                
                st.write("---")
                c_a, c_b = st.columns(2)
                with c_a:
                    st.write("#### 📅 Pengeluaran per Unit Kerja")
                    rekap_u = df_d.groupby(c_unit).agg(Total=('TOTAL', 'sum'), Jml_PO=(c_po, 'nunique')).reset_index()
                    rekap_u['Total'] = rekap_u['Total'].apply(format_rupiah)
                    st.dataframe(rekap_u, use_container_width=True)
                with c_b:
                    st.write("#### 🏆 Top 10 Item (by PO)")
                    top_i = df_d[df_d['NAMA BAKU'] != ""].groupby('NAMA BAKU')[c_po].nunique().sort_values(ascending=False).head(10)
                    st.bar_chart(top_i)
        else: st.warning("Database transaksi masih kosong.")
    except Exception as e: st.error(f"Dashboard Error: {e}")