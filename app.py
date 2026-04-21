import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import time
import json
import re
import gspread
from google.oauth2.service_account import Credentials
from streamlit_option_menu import option_menu
import plotly.express as px

# ==========================================
# 1. KONFIGURASI HALAMAN & TAMPILAN
# ==========================================
st.set_page_config(layout="wide", page_title="Sistem Purchasing Panca Budi", page_icon="📦")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 2. SISTEM KONEKSI GOOGLE SHEETS
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

# ==========================================
# 3. HELPER FUNCTIONS (FUNGSI PEMBANTU)
# ==========================================
def format_rupiah(angka):
    try: return f"Rp {int(angka):,}".replace(',', '.')
    except: return "Rp 0"

def convert_gdrive_link(url):
    if not isinstance(url, str): return ""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    if match:
        return f"https://drive.google.com/thumbnail?id={match.group(1)}&sz=w800"
    return url

def extract_code(text):
    try: return text.split('(')[1].split(')')[0].strip().zfill(3) 
    except: return "000"

def generate_new_sku(prefix_val, kat_full, det_full, current_df):
    prefix = str(prefix_val).strip().zfill(3)
    c_kat = extract_code(str(kat_full))
    c_det = extract_code(str(det_full))
    pattern = f"{prefix}-{c_kat}-{c_det}-"
    df_match = current_df[current_df['NOMOR SKU'].astype(str).str.contains(pattern, na=False)]
    if not df_match.empty:
        last_nums = []
        for s in df_match['NOMOR SKU'].astype(str):
            try: last_nums.append(int(s.split('-')[-1]))
            except: pass
        next_val = max(last_nums) + 1 if last_nums else 1
    else: next_val = 1
    return f"{prefix}-{c_kat}-{c_det}-{next_val:03d}"

# ==========================================
# 4. LOAD & PERSIAPAN MASTER DATA
# ==========================================
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    df_master = df_master.dropna(subset=['NAMA BAKU'])
    
    if 'KATEGORI' in df_master.columns: df_master['KATEGORI'] = df_master['KATEGORI'].ffill().astype(str).str.strip().str.upper()
    if 'DETAIL KATEGORI' in df_master.columns: df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill().astype(str).str.strip().str.upper()
    
    # Pre-processing untuk Fuzzy Matching & Pencarian
    df_master['LOOKUP'] = df_master['NAMA BAKU'].astype(str).str.upper()
    if 'NAMA ITEM' in df_master.columns:  # Mendukung kolom Nama Lapangan/Alias
        df_master['LOOKUP'] += " " + df_master['NAMA ITEM'].fillna("").astype(str).str.upper()
        
    df_master_unique = df_master.drop_duplicates(subset=['NAMA BAKU'], keep='last')
    master_map = df_master_unique.set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['LOOKUP'].tolist()
    lookup_to_baku = dict(zip(df_master['LOOKUP'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data dari Sheet 1: {e}"); st.stop()

# ==========================================
# 5. SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3050/3050253.png", width=80) 
    st.title("Sistem Master Data")
    st.write("**Purchasing Regional**")
    
    if st.button("🔄 Sinkronisasi Data (Refresh)", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
        
    st.write("---")
    
    menu = option_menu(
        menu_title="", 
        options=["Pembersihan PO", "Pencarian Barang", "E-Catalog & Studio", "Database Vendor", "Dashboard Laporan", "Maintenance Data"],
        icons=["magic", "search", "images", "shop", "bar-chart-line", "tools"], 
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": "transparent"},
            "icon": {"color": "#2e7b32", "font-size": "16px"}, 
            "nav-link": {"font-size": "14px", "text-align": "left", "margin":"0px", "--hover-color": "#e2e6ea", "border-radius": "8px"},
            "nav-link-selected": {"background-color": "#2e7b32", "color": "white", "icon-color": "white"},
        }
    )

# ==========================================
# MENU 1: PEMBERSIHAN PO (DUAL-FORMAT)
# ==========================================
if menu == "Pembersihan PO":
    st.header("✨ Upload & Pembersihan Laporan PO")
    
    col_u, col_f = st.columns(2)
    with col_u:
        pilihan_unit = ["PBI CPR", "PBI PML", "PBI MAUK", "PP CUP", "PIH", "RA", "PGP", "- Auto-Detect -"]
        unit_kerja = st.selectbox("🏢 Pilih Unit Kerja (Default):", pilihan_unit)
    with col_f:
        tipe_format = st.radio("⚙️ Tipe File Excel:", ["Format ERP (Laporan per No Bukti)", "Format Standar (Tabel Biasa)"])

    file_po = st.file_uploader("Upload Excel Laporan (.xlsx/.xls)", type=["xlsx", "xls"])
    
    if file_po:
        try:
            df_raw = pd.read_excel(file_po, header=None)
            final_data = []
            
            # --- MESIN 1: PARSER FORMAT ERP (Miring-Miring) ---
            if tipe_format == "Format ERP (Laporan per No Bukti)":
                current_po, current_tgl, current_vendor, current_curr = "-", "-", "-", "RP"
                
                for i, row in df_raw.iterrows():
                    row_vals = [str(x).strip() for x in row.values if str(x).strip() not in ['nan', 'None', '']]
                    full_str = " | ".join(row_vals).upper()

                    if not row_vals or "SUBTOTAL" in full_str or "LAPORAN PO" in full_str or "GRAND TOTAL" in full_str or "NO TRANS" in full_str:
                        continue

                    # Deteksi Header Transaksi
                    if "EXCLUDE" in full_str or "INCLUDE" in full_str:
                        current_po = row_vals[0]
                        for val in row_vals:
                            if "AM" in val.upper() or "PM" in val.upper() or re.match(r'\d{2}/\d{2}/\d{4}', val):
                                current_tgl = val.split()[0]; break
                        current_vendor = "-"
                        for val in row_vals:
                            if " - " in val:
                                current_vendor = val.split(" - ")[-1].strip(); break
                        current_curr = "RP"
                        for curr in ["RP", "EUR", "CNY", "USD"]:
                            if curr in [v.upper() for v in row_vals]:
                                current_curr = curr; break
                        continue 

                    # Deteksi Item Barang
                    if current_po != "-":
                        teks_candidates = [v for v in row_vals if not re.match(r'^[0-9.,]+$', v)]
                        if len(teks_candidates) >= 2: nama_barang = teks_candidates[1]
                        elif len(teks_candidates) == 1: nama_barang = teks_candidates[0]
                        else: continue

                        angka_candidates = []
                        for v in row_vals:
                            if re.match(r'^[0-9.,]+$', v):
                                try: angka_candidates.append(float(v.replace('.', '').replace(',', '.')))
                                except: pass
                        
                        if len(angka_candidates) >= 2:
                            qty = angka_candidates[0]
                            harga = angka_candidates[2] if len(angka_candidates) >= 3 else angka_candidates[1]
                            unit_final = "PBI CPR" if "ceper" in file_po.name.lower() else "PBI PML" if "pemalang" in file_po.name.lower() else unit_kerja
                            
                            final_data.append({
                                "UNIT KERJA": unit_final, "NO PO": current_po, "TANGGAL": current_tgl, 
                                "VENDOR": current_vendor, "MATA UANG": current_curr, 
                                "ITEM_KOTOR": nama_barang, "QTY": qty, "HARGA": harga
                            })
                df_clean = pd.DataFrame(final_data)

            # --- MESIN 2: PARSER FORMAT STANDAR (Tabel Flat) ---
            else:
                header_idx = -1
                for i, row in df_raw.iterrows():
                    row_str = " ".join([str(val).upper() for val in row.values])
                    if 'BARANG' in row_str or 'ITEM' in row_str or 'BAHAN' in row_str:
                        header_idx = i; break
                
                if header_idx != -1:
                    df_po = pd.read_excel(file_po, skiprows=header_idx)
                    df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
                    
                    col_po = next((c for c in df_po.columns if 'BUKTI' in c or 'PO' in c), df_po.columns[0])
                    col_barang = next((c for c in df_po.columns if 'BARANG' in c or 'ITEM' in c), df_po.columns[1])
                    col_qty = next((c for c in df_po.columns if 'QTY' in c or 'JUMLAH' in c and 'RP' not in c), None)
                    col_harga = next((c for c in df_po.columns if 'HARGA' in c), None)
                    col_tgl = next((c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c or 'DATE' in c), None)
                    col_vendor_khusus = next((c for c in df_po.columns if any(x in c for x in ['VENDOR', 'PEMASOK', 'SUPPLIER'])), None)
                    
                    v_saat_ini, t_saat_ini, p_saat_ini = "-", "-", "-"
                    for i, row in df_po.iterrows():
                        val_barang = str(row[col_barang]).strip()
                        is_empty = (val_barang == '' or val_barang.lower() == 'nan' or 'UNNAMED' in val_barang.upper())
                        
                        if col_vendor_khusus and not pd.isna(row[col_vendor_khusus]): v_saat_ini = str(row[col_vendor_khusus]).strip()
                        elif is_empty:
                            for val in row.values:
                                v_str = str(val).strip()
                                if v_str and v_str.lower() != 'nan' and not any(x in v_str.upper() for x in ["JUMLAH", "RP", "TOTAL", "PPN"]):
                                    if len(v_str) > 2 and not v_str.replace('.', '').replace(',', '').isdigit():
                                        v_saat_ini = v_str; break 
                        
                        if col_tgl and not pd.isna(row[col_tgl]): t_saat_ini = str(row[col_tgl]).strip().split()[0]
                        if col_po and str(row[col_po]).strip().lower() != 'nan': p_saat_ini = str(row[col_po]).strip()
                        
                        if not is_empty and "JUMLAH" not in val_barang.upper() and val_barang.upper() != "RP":
                            final_data.append({
                                "UNIT KERJA": unit_kerja, "NO PO": p_saat_ini, "TANGGAL": t_saat_ini, 
                                "VENDOR": v_saat_ini, "MATA UANG": "RP", "ITEM_KOTOR": val_barang, 
                                "QTY": row[col_qty] if col_qty else 0, "HARGA": row[col_harga] if col_harga else 0
                            })
                    df_clean = pd.DataFrame(final_data)
                else:
                    df_clean = pd.DataFrame()

            # --- PROSES FUZZY MATCHING (Mencari di Master Data) ---
            if not df_clean.empty:
                st.success(f"Berhasil membaca {len(df_clean)} baris item barang.")
                
                if st.button("🚀 Proses Pembersihan & Sinkronkan ke Master Data", type="primary", use_container_width=True):
                    hasil_rows = []
                    for _, r in df_clean.iterrows():
                        match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                        if match and match[1] >= 75:
                            baku = lookup_to_baku[match[0]]
                            info = master_map.get(baku, {})
                            hasil_rows.append({
                                "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'], "MATA UANG": r.get('MATA UANG', 'RP'),
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, "QTY": r['QTY'], "SATUAN": info.get('SATUAN', '-'),
                                "HARGA": r['HARGA'], "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), "SKU": info.get('NOMOR SKU', '-')
                            })
                        else:
                            hasil_rows.append({
                                "UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'], "MATA UANG": r.get('MATA UANG', 'RP'),
                                "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ BARANG BARU", "QTY": r['QTY'], "SATUAN": "-",
                                "HARGA": r['HARGA'], "KATEGORI": "-", "DETAIL KATEGORI": "-", "SKU": "-"
                            })
                    st.session_state['hasil_po'] = pd.DataFrame(hasil_rows)
                    st.rerun()

        except Exception as e: st.error(f"Error Eksekusi: {e}")

    # --- TAMPILAN HASIL & UPLOAD KE DASHBOARD ---
    if 'hasil_po' in st.session_state:
        st.write("### 📄 Hasil Pembersihan")
        st.dataframe(st.session_state['hasil_po'], use_container_width=True)
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("💾 Kirim Data Bersih ke Dashboard (Sheet 4)", type="primary", use_container_width=True):
                try:
                    with st.spinner("Mengirim ke database..."):
                        client = get_gspread_client()
                        sheet_4 = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD))
                        sheet_4.append_rows(st.session_state['hasil_po'].fillna("").values.tolist())
                        st.success("Berhasil dikirim! Data masuk ke Dashboard.")
                        del st.session_state['hasil_po']; st.rerun()
                except Exception as e: st.error(e)
        with col_btn2:
            if st.button("❌ Batal / Reset", use_container_width=True):
                del st.session_state['hasil_po']; st.rerun()

# ==========================================
# MENU 2: PENCARIAN BARANG
# ==========================================
elif menu == "Pencarian Barang":
    st.header("🔍 Kamus & Pencarian Barang (Master Data)")
    kata_cari = st.text_input("Ketik Nama Barang / Alias / SKU:")
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

# ==========================================
# MENU 3: E-CATALOG & STUDIO GAMBAR
# ==========================================
elif menu == "E-Catalog & Studio":
    st.title("🖼️ Ultimate E-Catalog Panca Budi")
    t_cat, t_studio = st.tabs(["📖 Galeri E-Catalog", "🛠️ Studio Jodoh Gambar"])
    
    with t_cat:
        st.write("Silakan cari atau filter barang untuk melihat wujud fisiknya.")
        col_s, col_f = st.columns([2, 1])
        with col_s: search_cat = st.text_input("🔍 Cari Nama Barang atau SKU:")
        with col_f:
            list_kat = ["Semua Kategori"] + sorted([k for k in df_master_unique['KATEGORI'].unique() if str(k).strip() != ""])
            filter_cat = st.selectbox("📁 Filter Kategori:", list_kat)
        
        df_show = df_master_unique.copy()
        if filter_cat != "Semua Kategori": df_show = df_show[df_show['KATEGORI'] == filter_cat]
        if search_cat:
            df_show = df_show[df_show['NAMA BAKU'].astype(str).str.contains(search_cat, case=False) | df_show['NOMOR SKU'].astype(str).str.contains(search_cat, case=False)]
        
        st.markdown("---")
        if df_show.empty:
            st.warning("Barang tidak ditemukan.")
        else:
            cols = st.columns(4)
            for idx, (_, row) in enumerate(df_show.iterrows()):
                with cols[idx % 4]:
                    raw_link = str(row.get('LINK GAMBAR', '')).strip()
                    img_url = convert_gdrive_link(raw_link)
                    
                    st.markdown(f"<div style='border:1px solid #ddd; border-radius:10px; padding:15px; margin-bottom:15px; height:100%; box-shadow: 2px 2px 5px rgba(0,0,0,0.05);'>", unsafe_allow_html=True)
                    if img_url and "drive.google" in img_url:
                        st.image(img_url, use_column_width=True)
                    else:
                        st.markdown(f"<div style='background-color:#f0f2f6; height:150px; border-radius:8px; display:flex; align-items:center; justify-content:center;'><span style='color:#888;'>🚫 Belum Ada Foto</span></div>", unsafe_allow_html=True)
                    
                    st.markdown(f"<h5 style='margin-top:10px; font-size:15px; color:#2e7b32;'>{row['NAMA BAKU']}</h5>", unsafe_allow_html=True)
                    st.markdown(f"<p style='font-size:12px; color:#666; margin:0;'>SKU: {row['NOMOR SKU']}</p>", unsafe_allow_html=True)
                    st.markdown(f"<p style='font-size:14px; font-weight:bold; margin-top:5px;'>{format_rupiah(row.get('HARGA', 0))}</p>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

    with t_studio:
        st.write("### 📸 Mesin Pemasang Foto Barang")
        df_no_pic = df_master_unique[df_master_unique['LINK GAMBAR'].astype(str).str.strip() == ""]
        if df_no_pic.empty:
            st.success("🎉 Luar biasa! Semua barang di Master Data sudah memiliki foto!")
        else:
            barang_pilih = st.selectbox("1️⃣ Pilih Nama Barang yang akan diberi foto:", df_no_pic['NAMA BAKU'].tolist())
            link_input = st.text_input("2️⃣ Paste Link Share Google Drive (Anyone with link):")
            
            if link_input:
                st.write("**Preview Gambar:**")
                st.image(convert_gdrive_link(link_input), width=300)
                if st.button("💾 Simpan Link ke Master Data", type="primary"):
                    try:
                        with st.spinner("Menyimpan ke Google Sheets..."):
                            client = get_gspread_client()
                            sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                            cell = sheet_master.find(barang_pilih, in_column=2) # Asumsi NAMA BAKU di kolom B
                            if cell:
                                headers = sheet_master.row_values(1)
                                try:
                                    col_link_idx = headers.index('LINK GAMBAR') + 1
                                    sheet_master.update_cell(cell.row, col_link_idx, link_input)
                                    st.success(f"Berhasil! Foto terpasang."); time.sleep(1.5); st.cache_data.clear(); st.rerun()
                                except ValueError: st.error("Kolom 'LINK GAMBAR' belum ada di baris pertama Sheet 1 Anda.")
                    except Exception as e: st.error(f"Gagal menyimpan: {e}")

# ==========================================
# MENU 4: DATABASE VENDOR
# ==========================================
elif menu == "Database Vendor":
    st.header("🏢 Database Pencarian Vendor")
    keyword = st.text_input("Cari Nama Supplier atau Barang:")
    if keyword:
        try:
            df_v = load_data(GID_VENDOR)
            df_v.columns = df_v.columns.str.strip().str.upper()
            res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
            for _, v in res.iterrows():
                with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} - {v.get('KATEGORI', '-')} ({v.get('GRUP', '-')})"):
                    st.write(f"**PIC:** {v.get('PIC', '-')} | **Kontak:** {v.get('KONTAK', '-')}")
                    st.write(f"**Alamat:** {v.get('ALAMAT', '-')}")
        except: st.warning("Database Vendor belum siap atau GID tidak valid.")

# ==========================================
# MENU 5: DASHBOARD LAPORAN (INTELIJEN PER ITEM)
# ==========================================
elif menu == "Dashboard Laporan":
    st.title("📊 Executive Dashboard & Item Intelligence")
    
    try:
        client = get_gspread_client()
        sheet_dash = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD))
        data_dash = sheet_dash.get_all_values()
        
        if len(data_dash) > 1:
            df_d = pd.DataFrame(data_dash[1:], columns=data_dash[0])
            df_d.columns = df_d.columns.str.strip().str.upper()
            
            c_po = next((c for c in df_d.columns if 'PO' in c or 'BUKTI' in c), None)
            c_unit = next((c for c in df_d.columns if 'UNIT' in c or 'GRUP' in c), None)
            c_harga = next((c for c in df_d.columns if 'HARGA' in c), None)
            c_baku = next((c for c in df_d.columns if 'BAKU' in c), None)
            c_tgl = next((c for c in df_d.columns if 'TANGGAL' in c or 'TGL' in c or 'DATE' in c), None)
            
            if not all([c_po, c_unit, c_harga, c_baku, c_tgl]):
                st.error("Gagal mendeteksi kolom. Pastikan Sheet 4 ada kolom NO PO, TANGGAL, NAMA BAKU, HARGA, UNIT KERJA."); st.stop()
            
            # Cleaning Data Angka & Tanggal
            df_d['H_NUM'] = pd.to_numeric(df_d[c_harga].astype(str).str.upper().str.replace('RP', '').str.replace(r'[^0-9]', '', regex=True), errors='coerce').fillna(0)
            df_d['Q_NUM'] = pd.to_numeric(df_d['QTY'].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
            df_d['TOTAL'] = df_d['H_NUM'] * df_d['Q_NUM']
            df_d['DATE_CLEAN'] = pd.to_datetime(df_d[c_tgl], errors='coerce')
            df_d = df_d.dropna(subset=['DATE_CLEAN'])
            
            tab_summary, tab_item = st.tabs(["🌎 Ringkasan Global", "🔍 Analisis Per Item (Intelijen)"])
            
            with tab_summary:
                list_unit = ["Semua Unit Kerja"] + sorted([u for u in df_d[c_unit].unique() if str(u).strip() != ""])
                filter_unit = st.selectbox("🎯 Filter Unit Kerja:", list_unit, key="sb_unit_global")
                df_filtered = df_d[df_d[c_unit] == filter_unit] if filter_unit != "Semua Unit Kerja" else df_d
                
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Belanja", format_rupiah(df_filtered['TOTAL'].sum()))
                col2.metric("📄 Total Transaksi PO", f"{df_filtered[c_po].replace('', pd.NA).dropna().nunique()}")
                col3.metric("🏢 Unit Aktif", f"{df_filtered[c_unit].nunique()} Pabrik")
                
                c_a, c_b = st.columns([1, 1.5])
                with c_a:
                    st.write("#### 🍩 Porsi Anggaran per Pabrik")
                    if filter_unit == "Semua Unit Kerja":
                        rekap_u = df_filtered.groupby(c_unit)['TOTAL'].sum().reset_index()
                        rekap_u = rekap_u[rekap_u[c_unit].str.strip() != ""] 
                        fig_pie = px.pie(rekap_u, names=c_unit, values='TOTAL', hole=0.5, color_discrete_sequence=px.colors.sequential.Teal)
                        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                        fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
                        st.plotly_chart(fig_pie, use_container_width=True)
                    else: st.info(f"Menampilkan data khusus untuk **{filter_unit}**.")

                with c_b:
                    st.write("#### 📊 Top 10 Barang Dipesan")
                    df_valid = df_filtered[~df_filtered[c_baku].str.contains('CEK MANUAL|BARANG BARU', case=False, na=False)]
                    if not df_valid.empty:
                        df_valid = df_valid[df_valid[c_baku].str.strip() != ""]
                        top_i = df_valid.groupby(c_baku)[c_po].nunique().reset_index()
                        top_i.columns = ['Nama Barang', 'Jumlah PO']
                        top_i = top_i.sort_values(by='Jumlah PO', ascending=False).head(10)
                        fig_bar = px.bar(top_i, x='Jumlah PO', y='Nama Barang', orientation='h', text='Jumlah PO', color='Jumlah PO', color_continuous_scale='Greens')
                        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=10, b=10, l=10, r=10), coloraxis_showscale=False)
                        st.plotly_chart(fig_bar, use_container_width=True)
            
            with tab_item:
                list_barang_histori = df_d.drop_duplicates(subset=[c_baku]).sort_values(by=c_baku)[c_baku].tolist()
                barang_pilih = st.selectbox("🔍 Cari & Pilih Nama Barang untuk Analisis:", list_barang_histori, placeholder="Ketik nama barang...")
                
                if barang_pilih:
                    st.markdown(f"### Laporan Intelijen: {barang_pilih}")
                    info_master = df_master_unique[df_master_unique['NAMA BAKU'] == barang_pilih]
                    
                    c_img, c_meta = st.columns([1, 2])
                    with c_img:
                        if not info_master.empty:
                            img_url = convert_gdrive_link(str(info_master.iloc[0].get('LINK GAMBAR', '')).strip())
                            if img_url and "drive.google" in img_url: st.image(img_url, use_column_width=True)
                            else: st.info("🚫 Belum Ada Foto")
                        else: st.warning("Barang tidak ada di Master Data (Sheet 1).")

                    with c_meta:
                        if not info_master.empty:
                            row_m = info_master.iloc[0]
                            st.markdown(f"""
                            * **NOMOR SKU:** `{row_m.get('NOMOR SKU', '-')}`
                            * **KATEGORI:** {row_m.get('KATEGORI', '-')}
                            * **SATUAN BAKU:** {row_m.get('SATUAN', '-')}
                            * **VENDOR UTAMA (Master):** {row_m.get('VENDOR', '-')}
                            """)
                    
                    st.markdown("---")
                    df_item_histori = df_d[df_d[c_baku] == barang_pilih].sort_values(by='DATE_CLEAN')
                    
                    if not df_item_histori.empty:
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("💰 Total Belanja", format_rupiah(df_item_histori['TOTAL'].sum()))
                        m2.metric("📄 Total Transaksi (PO)", f"{df_item_histori[c_po].nunique()} Kali")
                        m3.metric("📊 Rata-rata Harga", format_rupiah(df_item_histori['H_NUM'].mean()))
                        m4.metric("🏢 Vendor Terakhir", df_item_histori.iloc[-1].get('VENDOR', '-'))

                        g_harga, g_qty = st.columns(2)
                        with g_harga:
                            fig_harga = px.line(df_item_histori, x='DATE_CLEAN', y='H_NUM', text='H_NUM', title="Tren Perubahan Harga", markers=True)
                            fig_harga.update_traces(textposition="top right"); st.plotly_chart(fig_harga, use_container_width=True)
                        with g_qty:
                            df_monthly = df_item_histori.resample('M', on='DATE_CLEAN').sum().reset_index()
                            fig_qty = px.bar(df_monthly, x='DATE_CLEAN', y='Q_NUM', text='Q_NUM', title="Volume Dipesan per Bulan")
                            st.plotly_chart(fig_qty, use_container_width=True)
                        
                        st.write("#### 📑 Histori Transaksi Detail")
                        df_table = df_item_histori[[c_tgl, c_po, c_unit, 'VENDOR', 'QTY', 'H_NUM', 'TOTAL']].copy()
                        df_table['H_NUM'] = df_table['H_NUM'].map(format_rupiah)
                        df_table['TOTAL'] = df_table['TOTAL'].map(format_rupiah)
                        st.dataframe(df_table, use_container_width=True, hide_index=True)

        else: st.warning("Database transaksi (Sheet 4) masih kosong.")
    except Exception as e: st.error(f"Dashboard Error: {e}")

# ==========================================
# MENU 6: MAINTENANCE DATA
# ==========================================
elif menu == "Maintenance Data":
    st.header("🛠️ Maintenance & Auto-Fill SKU")
    st.write("Fitur ini bertugas memindai Master Data (Sheet 1) dan secara otomatis membuatkan Nomor SKU 12 digit untuk barang yang belum punya.")
    
    invalid_mask = df_master_unique['NOMOR SKU'].isna() | (df_master_unique['NOMOR SKU'].astype(str).str.strip().str.len() < 10)
    df_missing = df_master_unique[invalid_mask]
    
    if not df_missing.empty:
        st.warning(f"⚠️ Ditemukan **{len(df_missing)}** barang tanpa Nomor SKU yang valid di Master Data!")
        st.dataframe(df_missing[['NAMA BAKU', 'KATEGORI', 'DETAIL KATEGORI', 'NOMOR SKU']], use_container_width=True)
        
        if st.button("🚀 Eksekusi Auto-Fill SKU Sekarang!", type="primary", use_container_width=True):
            with st.spinner("Menyuntikkan SKU baru secara massal..."):
                try:
                    client = get_gspread_client()
                    sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                    all_data = sheet_master.get_all_values()
                    headers = [str(h).strip().upper() for h in all_data[0]]
                    df_m = pd.DataFrame(all_data[1:], columns=headers)
                    
                    col_sku = next((c for c in headers if 'SKU' in c), None)
                    col_kat = next((c for c in headers if 'KATEGORI' in c and 'DETAIL' not in c), None)
                    col_det = next((c for c in headers if 'DETAIL' in c), None)
                    
                    if col_sku and col_kat and col_det:
                        for idx, row in df_m.iterrows():
                            val_sku = str(row[col_sku]).strip()
                            if len(val_sku) < 10 or val_sku.upper() in ['NAN', 'NONE', 'NULL', '#N/A', '']:
                                df_m.at[idx, col_sku] = generate_new_sku("001", row[col_kat], row[col_det], current_df=df_m)
                        sheet_master.clear()
                        sheet_master.update(values=[df_m.columns.tolist()] + df_m.values.tolist())
                        st.success("✅ BERHASIL! Semua baris di Master Data kini memiliki SKU valid."); time.sleep(2); st.rerun()
                    else: st.error("Gagal mendeteksi kolom SKU/Kategori di Sheet 1.")
                except Exception as e: st.error(f"Terjadi kesalahan: {e}")
    else: st.success("🎉 Database Anda sehat! Semua barang di Sheet 1 memiliki SKU 12 Digit yang valid.")