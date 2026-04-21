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
    try: return text.split('(')[1].split(')')[0].strip().zfill(3) 
    except: return "000"

# --- FUNGSI CONVERT LINK GDRIVE ---
def convert_gdrive_link(url):
    if not isinstance(url, str): return ""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    if match:
        file_id = match.group(1)
        return f"https://drive.google.com/uc?export=view&id={file_id}"
    return url

# --- PERSIAPAN KAMUS PINTAR ---
try:
    df_master = load_data(GID_MASTER)
    df_master.columns = df_master.columns.str.strip().str.upper()
    df_master = df_master.dropna(subset=['NAMA BAKU'])
    
    if 'KATEGORI' in df_master.columns: df_master['KATEGORI'] = df_master['KATEGORI'].ffill().astype(str).str.strip().str.upper()
    if 'DETAIL KATEGORI' in df_master.columns: df_master['DETAIL KATEGORI'] = df_master['DETAIL KATEGORI'].ffill().astype(str).str.strip().str.upper()
    if 'KATA KUNCI' not in df_master.columns: df_master['KATA KUNCI'] = ""
    df_master['KATA KUNCI'] = df_master['KATA KUNCI'].fillna("").astype(str)
    
    # Deteksi Kolom LINK GAMBAR
    if 'LINK GAMBAR' not in df_master.columns: df_master['LINK GAMBAR'] = ""
    df_master['LINK GAMBAR'] = df_master['LINK GAMBAR'].fillna("").astype(str)
    
    df_master['Lookup'] = df_master['NAMA BAKU'].astype(str) + " " + df_master['KATA KUNCI']
    df_master_unique = df_master.drop_duplicates(subset=['NAMA BAKU'], keep='last')
    master_map = df_master_unique.set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['Lookup'].tolist()
    lookup_to_baku = dict(zip(df_master['Lookup'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

def generate_new_sku(prefix_val, kat_full, det_full, current_df=df_master):
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
# MENU 1: PEMBERSIHAN PO
# ==========================================
if menu == "Pembersihan PO":
    st.header("Upload & Pembersihan Laporan PO")
    pilihan_unit = ["- Auto-Detect dari Keterangan -", "PBI CPR", "PBI PML", "PBI MAUK", "PP CUP", "PIH", "PIH BHN PENOLONG", "RA", "PGP"]
    unit_kerja = st.selectbox("🏢 Pilih Unit Kerja:", pilihan_unit)
    file_po = st.file_uploader("Upload Excel Laporan (.xlsx/.xls)", type=["xlsx", "xls"])
    
    if file_po:
        try:
            raw_excel = pd.read_excel(file_po, header=None)
            header_idx = -1
            for i, row in raw_excel.iterrows():
                row_str = " ".join([str(val).upper() for val in row.values])
                if 'BARANG' in row_str or 'ITEM' in row_str or 'BAHAN' in row_str:
                    header_idx = i; break
            
            if header_idx != -1:
                df_po = pd.read_excel(file_po, skiprows=header_idx)
                df_po.columns = df_po.columns.astype(str).str.strip().str.upper()
                
                vendor_saat_ini, tgl_saat_ini, po_saat_ini, unit_saat_ini = "-", "-", "-", "BELUM DITENTUKAN"
                final_data = []
                
                col_po = next((c for c in df_po.columns if 'BUKTI' in c or 'PO' in c), df_po.columns[0])
                col_barang = next((c for c in df_po.columns if 'BARANG' in c or 'ITEM' in c), df_po.columns[1])
                col_qty = next((c for c in df_po.columns if 'QTY' in c or 'JUMLAH' in c and 'RP' not in c), None)
                col_harga = next((c for c in df_po.columns if 'HARGA' in c), None)
                col_tgl = next((c for c in df_po.columns if 'TGL' in c or 'TANGGAL' in c or 'DATE' in c), None)
                col_ket = next((c for c in df_po.columns if 'KETERANGAN' in c or 'KET' in c or 'ALAMAT' in c), None)
                col_vendor_khusus = next((c for c in df_po.columns if any(x in c for x in ['VENDOR', 'PEMASOK', 'SUPPLIER'])), None)
                
                for i, row in df_po.iterrows():
                    val_barang = str(row[col_barang]).strip()
                    is_empty = (val_barang == '' or val_barang.lower() == 'nan' or 'UNNAMED' in val_barang.upper())
                    
                    if col_vendor_khusus and not pd.isna(row[col_vendor_khusus]) and str(row[col_vendor_khusus]).strip().lower() != 'nan':
                        vendor_saat_ini = str(row[col_vendor_khusus]).strip()
                    elif is_empty:
                        for val in row.values:
                            v_str = str(val).strip()
                            if v_str and v_str.lower() != 'nan' and not any(x in v_str.upper() for x in ["JUMLAH", "RP", "TOTAL", "PPN"]):
                                if len(v_str) > 2 and not v_str.replace('.', '').replace(',', '').isdigit():
                                    vendor_saat_ini = v_str; break 
                    
                    if col_tgl and not pd.isna(row[col_tgl]):
                        t_val = str(row[col_tgl]).strip()
                        if len(t_val) >= 4 and "JUMLAH" not in t_val.upper(): tgl_saat_ini = t_val.split()[0]
                    
                    curr_po = str(row[col_po]).strip() if col_po else ""
                    if curr_po and curr_po.lower() != 'nan' and any(char.isdigit() for char in curr_po):
                        po_saat_ini = curr_po
                            
                    ket_val = str(row[col_ket]).strip().upper() if col_ket else ""
                    if ket_val and ket_val.lower() != 'nan':
                        if "KEAMANAN" in ket_val: unit_saat_ini = "PBI CPR"
                        elif "ARYA KEMUNING" in ket_val: unit_saat_ini = "PBI MAUK"
                        elif "AGUS HALIM" in ket_val: unit_saat_ini = "PP CUP"
                        elif "PEMALANG" in ket_val: unit_saat_ini = "PBI PML"
                            
                    row_unit = unit_saat_ini if unit_kerja == "- Auto-Detect dari Keterangan -" else unit_kerja
                    
                    if not is_empty and "JUMLAH" not in val_barang.upper() and val_barang.upper() != "RP":
                        final_data.append({
                            "UNIT KERJA": row_unit, "NO PO": po_saat_ini, "TANGGAL": tgl_saat_ini, 
                            "VENDOR": vendor_saat_ini, "ITEM_KOTOR": val_barang, 
                            "QTY": row[col_qty] if col_qty else 0, "HARGA": row[col_harga] if col_harga else 0
                        })
                
                df_clean = pd.DataFrame(final_data)
                
                if st.button("🚀 Proses Pembersihan & Auto-Detect", type="primary", use_container_width=True):
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

    if 'hasil_po' in st.session_state:
        t1, t2, t3 = st.tabs(["📄 Hasil Pembersihan", "🆕 Registrasi SKU", "📊 Rekap"])
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
                item_select = st.selectbox("Pilih barang:", new_items)
                c_p, c_a, c_b = st.columns(3)
                with c_p:
                    pref_list = ["001 - Sparepart Mesin", "002 - Supporting Material", "003 - Bahan Baku", "004 - ATK & Umum", "005 - IT", "✨ + Tambah Kode..."]
                    pd = st.selectbox("Tipe (Blok 1):", pref_list)
                    prefix_sel = st.text_input("Ketik Kode:", max_chars=3) if pd == "✨ + Tambah Kode..." else pd[:3]
                with c_a:
                    kl = sorted([k for k in df_master['KATEGORI'].unique() if k and k != '-']) + ["✨ + Tambah Kategori..."]
                    kd = st.selectbox("Kategori (Blok 2):", kl)
                    kat_sel = st.text_input("Kategori Baru:") if kd == "✨ + Tambah Kategori..." else kd
                with c_b:
                    dl = sorted([d for d in df_master[df_master['KATEGORI'] == kd]['DETAIL KATEGORI'].unique() if d and d != '-']) + ["✨ + Tambah Detail..."] if kd != "✨ + Tambah Kategori..." else ["✨ + Tambah Detail..."]
                    dd = st.selectbox("Detail (Blok 3):", dl)
                    det_sel = st.text_input("Detail Baru:") if dd == "✨ + Tambah Detail..." else dd
                
                if prefix_sel and kat_sel and det_sel:
                    sku_baru = generate_new_sku(prefix_sel, kat_sel, det_sel)
                    st.info(f"**Saran SKU Baru:** `{sku_baru}`")
                    if st.button("🔥 Daftarkan & Update PO", type="primary"):
                        try:
                            rd = st.session_state['hasil_po'][st.session_state['hasil_po']['NAMA ITEM'] == item_select].iloc[0]
                            client = get_gspread_client(); sm = client.open_by_key(SHEET_ID).get_worksheet(0)
                            new_row = [item_select, item_select, kat_sel, det_sel, sku_baru, "", "PCS", rd.get('HARGA', 0), rd.get('QTY', 0), rd.get('VENDOR', '-'), rd.get('UNIT KERJA', '-'), rd.get('TANGGAL', '-'), ""]
                            sm.append_row(new_row)
                            st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'NAMA BAKU'] = item_select
                            st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'SKU'] = sku_baru
                            st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'KATEGORI'] = kat_sel
                            st.session_state['hasil_po'].loc[st.session_state['hasil_po']['NAMA ITEM'] == item_select, 'DETAIL KATEGORI'] = det_sel
                            st.success(f"Mantap! Terdaftar dengan SKU {sku_baru}"); time.sleep(1); st.rerun()
                        except Exception as e: st.error(f"Gagal Registrasi: {e}")
            else: st.success("Semua barang sudah terdaftar.")

# ==========================================
# MENU 2: PENCARIAN BARANG
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
                    st.markdown(f"<p style='font-size:12px; color:#666; margin:0;'>Kat: {row['KATEGORI']}</p>", unsafe_allow_html=True)
                    st.markdown(f"<p style='font-size:14px; font-weight:bold; margin-top:5px;'>{format_rupiah(row.get('HARGA', 0))}</p>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)

    with t_studio:
        st.write("### 📸 Mesin Pemasang Foto Barang")
        st.info("Pilih barang yang belum punya foto, lalu Paste link Google Drive-nya di sini.")
        
        df_no_pic = df_master_unique[df_master_unique['LINK GAMBAR'].astype(str).str.strip() == ""]
        if df_no_pic.empty:
            st.success("🎉 Luar biasa! Semua barang di Master Data sudah memiliki foto!")
        else:
            barang_pilih = st.selectbox("1️⃣ Pilih Nama Barang yang akan diberi foto:", df_no_pic['NAMA BAKU'].tolist())
            link_input = st.text_input("2️⃣ Paste Link Share Google Drive (Anyone with link):", placeholder="https://drive.google.com/file/d/..../view?usp=sharing")
            
            if link_input:
                st.write("**Preview Gambar:**")
                st.image(convert_gdrive_link(link_input), width=300)
                
                if st.button("💾 Simpan Link ke Master Data", type="primary"):
                    try:
                        with st.spinner("Menyimpan ke Google Sheets..."):
                            client = get_gspread_client()
                            sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                            
                            cell = sheet_master.find(barang_pilih, in_column=2) # NAMA BAKU ada di kolom B
                            if cell:
                                row_idx = cell.row
                                headers = sheet_master.row_values(1)
                                try:
                                    col_link_idx = headers.index('LINK GAMBAR') + 1
                                    sheet_master.update_cell(row_idx, col_link_idx, link_input)
                                    st.success(f"Berhasil! Foto {barang_pilih} sudah terpasang. Silakan cek di Tab E-Catalog!")
                                    time.sleep(1.5)
                                    st.cache_data.clear()
                                    st.rerun()
                                except ValueError:
                                    st.error("Kolom 'LINK GAMBAR' belum dibuat di baris pertama Sheet 1 Anda! Buat dulu ya Bosku.")
                    except Exception as e:
                        st.error(f"Gagal menyimpan: {e}")

# ==========================================
# MENU 4: DATABASE VENDOR
# ==========================================
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

# ==========================================
# MENU 5: DASHBOARD LAPORAN
# ==========================================
elif menu == "Dashboard Laporan":
    st.title("📊 Executive Dashboard")
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
            
            if c_po and c_unit and c_harga and c_baku:
                h_str = df_d[c_harga].astype(str).str.upper().str.replace('RP', '', regex=False).str.split(',').str[0].str.replace(r'[^0-9]', '', regex=True)
                df_d['H_NUM'] = pd.to_numeric(h_str, errors='coerce').fillna(0)
                df_d['Q_NUM'] = pd.to_numeric(df_d['QTY'], errors='coerce').fillna(0)
                df_d['TOTAL'] = df_d['H_NUM'] * df_d['Q_NUM']
                
                list_unit = ["Semua Unit Kerja"] + sorted([u for u in df_d[c_unit].unique() if str(u).strip() != ""])
                col_f1, col_f2 = st.columns([1, 3])
                with col_f1: filter_unit = st.selectbox("🎯 Filter Berdasarkan:", list_unit)
                df_filtered = df_d[df_d[c_unit] == filter_unit] if filter_unit != "Semua Unit Kerja" else df_d
                
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                col1.metric("💰 Total Belanja", format_rupiah(df_filtered['TOTAL'].sum()))
                col2.metric("📄 Total Lembar PO", f"{df_filtered[c_po].replace('', pd.NA).dropna().nunique()} Transaksi")
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
                    st.write("#### 📊 Top 10 Barang Paling Sering Dipesan")
                    df_valid = df_filtered[~df_filtered[c_baku].str.contains('CEK MANUAL|BARANG BARU', case=False, na=False)]
                    if not df_valid.empty:
                        df_valid = df_valid[df_valid[c_baku].str.strip() != ""]
                        top_i = df_valid.groupby(c_baku)[c_po].nunique().reset_index()
                        top_i.columns = ['Nama Barang', 'Jumlah PO']
                        top_i = top_i.sort_values(by='Jumlah PO', ascending=False).head(10)
                        fig_bar = px.bar(top_i, x='Jumlah PO', y='Nama Barang', orientation='h', text='Jumlah PO', color='Jumlah PO', color_continuous_scale='Greens')
                        fig_bar.update_traces(textposition='outside')
                        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=10, b=10, l=10, r=10), coloraxis_showscale=False, xaxis_title=None, yaxis_title=None)
                        st.plotly_chart(fig_bar, use_container_width=True)
        else: st.warning("Database transaksi masih kosong.")
    except Exception as e: st.error(f"Dashboard Error: {e}")

# ==========================================
# MENU 6: MAINTENANCE DATA
# ==========================================
elif menu == "Maintenance Data":
    st.header("🛠️ Maintenance & Auto-Fill Master Data")
    invalid_mask = df_master['NOMOR SKU'].isna() | (df_master['NOMOR SKU'].astype(str).str.strip().str.len() < 10)
    df_missing = df_master[invalid_mask]
    if not df_missing.empty:
        st.warning(f"⚠️ Ditemukan **{len(df_missing)}** barang tanpa Nomor SKU (atau SKU tidak valid) di Master Data!")
        st.dataframe(df_missing[['NAMA BAKU', 'KATEGORI', 'DETAIL KATEGORI', 'NOMOR SKU']])
        if st.button("🚀 Eksekusi Auto-Fill SKU Sekarang!", type="primary", use_container_width=True):
            with st.spinner("Memindai Sheet 1 dan menyuntikkan SKU baru secara massal..."):
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
                            if len(val_sku) < 10 or val_sku.upper() in ['NAN', 'NONE', 'NULL', '#N/A']:
                                new_sku = generate_new_sku("001", row[col_kat], row[col_det], current_df=df_m)
                                df_m.at[idx, col_sku] = new_sku
                        sheet_master.clear()
                        sheet_master.update(values=[df_m.columns.tolist()] + df_m.values.tolist())
                        st.success("✅ BERHASIL! Semua baris kosong di Master Data kini telah memiliki SKU 12 Digit yang valid."); time.sleep(2); st.rerun()
                    else: st.error("Gagal mendeteksi kolom.")
                except Exception as e: st.error(f"Terjadi kesalahan: {e}")
    else: st.success("🎉 Database Anda sehat! Semua barang memiliki SKU 12 Digit yang valid.")