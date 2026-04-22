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
# 1. KONFIGURASI HALAMAN & TAMPILAN ERP INTERNASIONAL
# ==========================================
st.set_page_config(layout="wide", page_title="ERP Purchasing | Panca Budi", page_icon="🏢")

# --- CUSTOM CSS MODERN SAAS / ERP ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* Background utama dibuat off-white elegan */
    .main { 
        background-color: #F8FAFC; 
    }
    
    /* Desain Metric Cards / Kotak Angka */
    .stMetric { 
        background-color: #FFFFFF; 
        padding: 24px; 
        border-radius: 16px; 
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
        border: 1px solid #E2E8F0;
        transition: transform 0.2s ease-in-out;
    }
    .stMetric:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
    }
    
    /* Merapikan Tab */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    
    /* Tombol Utama */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        letter-spacing: 0.5px;
        transition: all 0.3s;
    }
    
    /* Judul Header */
    h1, h2, h3 {
        color: #0F172A;
        font-weight: 700;
    }
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
# 3. HELPER FUNCTIONS
# ==========================================
def format_rupiah(angka):
    try: return f"Rp {int(angka):,}".replace(',', '.')
    except: return "Rp 0"

def convert_gdrive_link(url):
    if not isinstance(url, str): return ""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    if match: return f"https://drive.google.com/thumbnail?id={match.group(1)}&sz=w800"
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
    
    df_master['LOOKUP'] = df_master['NAMA BAKU'].astype(str).str.upper()
    if 'NAMA ITEM' in df_master.columns: df_master['LOOKUP'] += " " + df_master['NAMA ITEM'].fillna("").astype(str).str.upper()
        
    df_master_unique = df_master.drop_duplicates(subset=['NAMA BAKU'], keep='last')
    master_map = df_master_unique.set_index('NAMA BAKU').to_dict('index')
    list_lookup = df_master['LOOKUP'].tolist()
    lookup_to_baku = dict(zip(df_master['LOOKUP'], df_master['NAMA BAKU']))
except Exception as e:
    st.error(f"⚠️ Gagal Load Master Data: {e}"); st.stop()

# ==========================================
# 5. SIDEBAR NAVIGATION (ELEGANT DESIGN)
# ==========================================
with st.sidebar:
    # Logo Typografi Corporate
    st.markdown("""
        <div style='text-align: center; padding: 10px 0 20px 0; border-bottom: 1px solid #E2E8F0; margin-bottom: 20px;'>
            <h1 style='color: #047857; font-weight: 800; margin: 0; font-size: 32px; letter-spacing: -1px;'>PANCA BUDI</h1>
            <p style='color: #64748B; font-size: 11px; font-weight: 700; letter-spacing: 2px; margin: 0;'>PURCHASING SYSTEM</p>
        </div>
        """, unsafe_allow_html=True)
    
    if st.button("🔄 Sync Database", use_container_width=True):
        st.cache_data.clear(); st.rerun()
        
    st.write("") # Spacer
    
    menu = option_menu(
        menu_title="", 
        options=["Pembersihan PO", "Pencarian Barang", "E-Catalog & Studio", "Database Vendor", "Dashboard Laporan", "Maintenance Data"],
        icons=["magic", "search", "images", "shop", "bar-chart-line", "tools"], 
        default_index=2, # Default ke E-Catalog biar Bosku langsung bisa cek hasilnya
        styles={
            "container": {"padding": "0!important", "background-color": "transparent"},
            "icon": {"color": "#64748B", "font-size": "18px"}, 
            "nav-link": {"color": "#334155", "font-weight": "500", "font-size": "15px", "text-align": "left", "margin":"4px 0", "--hover-color": "#F1F5F9", "border-radius": "8px"},
            "nav-link-selected": {"background-color": "#047857", "color": "white", "icon-color": "white", "font-weight": "600"},
        }
    )

# ==========================================
# MENU 1: PEMBERSIHAN PO
# ==========================================
if menu == "Pembersihan PO":
    st.markdown("<h2>✨ Pembersihan Data PO</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B;'>Upload laporan ERP mentah untuk diproses secara otomatis.</p>", unsafe_allow_html=True)
    
    col_u, col_f = st.columns(2)
    with col_u:
        unit_kerja = st.selectbox("🏢 Unit Kerja (Default):", ["PBI CPR", "PBI PML", "PBI MAUK", "PP CUP", "PIH", "RA", "PGP", "- Auto-Detect -"])
    with col_f:
        tipe_format = st.radio("⚙️ Tipe Dokumen:", ["Format ERP (Laporan per No Bukti)", "Format Standar (Tabel Biasa)"])

    file_po = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx", "xls"])
    
    if file_po:
        try:
            df_raw = pd.read_excel(file_po, header=None)
            final_rows = []
            
            if tipe_format == "Format ERP (Laporan per No Bukti)":
                curr_po, curr_tgl, curr_vendor, curr_money = "-", "-", "-", "RP"
                for i, row in df_raw.iterrows():
                    row_vals = [str(x).strip() for x in row.values if str(x).strip() not in ['nan', 'None', '']]
                    full_str = " | ".join(row_vals).upper()

                    if not row_vals or any(x in full_str for x in ["SUBTOTAL", "LAPORAN PO", "GRAND TOTAL", "NO TRANS"]): continue

                    if "EXCLUDE" in full_str or "INCLUDE" in full_str:
                        curr_po = row_vals[0]
                        for val in row_vals:
                            m1 = re.search(r'\d{4}-\d{2}-\d{2}', val)
                            m2 = re.search(r'\d{2}/\d{2}/\d{4}', val)
                            if m1: curr_tgl = m1.group(0); break
                            elif m2: curr_tgl = m2.group(0); break
                            
                        curr_vendor = "-"
                        for val in row_vals:
                            if " - " in val: curr_vendor = val.split(" - ")[-1].strip(); break
                        
                        curr_money = "RP"
                        for m in ["RP", "EUR", "CNY", "USD"]:
                            if m in row_vals: curr_money = m; break
                        continue 

                    if curr_po != "-":
                        num_data = []
                        for v in reversed(row_vals):
                            if re.match(r'^-?[0-9.,]+$', v):
                                try:
                                    v_str = str(v).strip()
                                    if ',' in v_str and '.' in v_str:
                                        if v_str.rfind(',') > v_str.rfind('.'): num_data.insert(0, float(v_str.replace('.', '').replace(',', '.')))
                                        else: num_data.insert(0, float(v_str.replace(',', '')))
                                    elif ',' in v_str: num_data.insert(0, float(v_str.replace(',', '.')))
                                    else: num_data.insert(0, float(v_str))
                                except: break
                            else: break 
                        
                        text_data = row_vals[:-len(num_data)] if len(num_data) > 0 else row_vals
                        
                        if len(text_data) >= 1 and len(num_data) >= 2:
                            item_name = text_data[1] if len(text_data) > 1 else text_data[0]
                            qty = num_data[0]
                            harga = num_data[2] if len(num_data) >= 3 else num_data[1]
                            unit_final = "PBI CPR" if "ceper" in file_po.name.lower() else "PBI PML" if "pemalang" in file_po.name.lower() else unit_kerja
                            
                            final_rows.append({"UNIT KERJA": unit_final, "NO PO": curr_po, "TANGGAL": curr_tgl, "VENDOR": curr_vendor, "MATA UANG": curr_money, "ITEM_KOTOR": item_name, "QTY": qty, "HARGA": harga})
                df_clean = pd.DataFrame(final_rows)
            else:
                st.warning("Gunakan Format ERP untuk data Panca Budi."); df_clean = pd.DataFrame()

            if not df_clean.empty:
                st.success(f"✔️ Berhasil mengekstrak {len(df_clean)} baris item.")
                st.write("**Preview Data Mentah:**")
                st.dataframe(df_clean.head(5), use_container_width=True)
                
                if st.button("🚀 Proses AI Matching", type="primary", use_container_width=True):
                    hasil_rows = []
                    for _, r in df_clean.iterrows():
                        match = process.extractOne(str(r['ITEM_KOTOR']), list_lookup, scorer=fuzz.token_set_ratio)
                        if match and match[1] >= 75:
                            baku = lookup_to_baku[match[0]]; info = master_map.get(baku, {})
                            hasil_rows.append({"UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'], "MATA UANG": r.get('MATA UANG', 'RP'), "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": baku, "QTY": r['QTY'], "SATUAN": info.get('SATUAN', '-'), "HARGA": r['HARGA'], "KATEGORI": info.get('KATEGORI', '-'), "DETAIL KATEGORI": info.get('DETAIL KATEGORI', '-'), "SKU": info.get('NOMOR SKU', '-')})
                        else:
                            hasil_rows.append({"UNIT KERJA": r['UNIT KERJA'], "NO PO": r['NO PO'], "TANGGAL": r['TANGGAL'], "VENDOR": r['VENDOR'], "MATA UANG": r.get('MATA UANG', 'RP'), "NAMA ITEM": r['ITEM_KOTOR'], "NAMA BAKU": "⚠️ BARANG BARU", "QTY": r['QTY'], "SATUAN": "-", "HARGA": r['HARGA'], "KATEGORI": "-", "DETAIL KATEGORI": "-", "SKU": "-"})
                    st.session_state['hasil_po'] = pd.DataFrame(hasil_rows); st.rerun()

        except Exception as e: st.error(f"System Error: {e}")

    if 'hasil_po' in st.session_state:
        st.markdown("### 📑 Review Hasil Pembersihan")
        st.dataframe(st.session_state['hasil_po'], use_container_width=True)
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("💾 Simpan ke Database Induk", type="primary", use_container_width=True):
                try:
                    with st.spinner("Sinkronisasi ke Google Sheets..."):
                        client = get_gspread_client()
                        sheet_4 = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD))
                        sheet_4.append_rows(st.session_state['hasil_po'].fillna("").values.tolist())
                        st.success("Tersimpan!"); del st.session_state['hasil_po']; st.rerun()
                except Exception as e: st.error(e)
        with col_btn2:
            if st.button("❌ Batal", use_container_width=True): del st.session_state['hasil_po']; st.rerun()

# ==========================================
# MENU 2: PENCARIAN BARANG
# ==========================================
elif menu == "Pencarian Barang":
    st.markdown("<h2>🔍 Global Search Engine</h2>", unsafe_allow_html=True)
    kata_cari = st.text_input("Ketik Kata Kunci (Nama Barang / SKU):")
    
    if kata_cari:
        hasil = process.extract(kata_cari, list_lookup, scorer=fuzz.token_set_ratio, limit=10)
        res_list = []
        for m in hasil:
            if m[1] >= 40:
                baku = lookup_to_baku[m[0]]; info = master_map.get(baku, {})
                res_list.append({"Match": f"{m[1]}%", "Nama Baku": baku, "SKU": info.get('NOMOR SKU', '-'), "Kategori": info.get('KATEGORI', '-'), "Est. Harga": info.get('HARGA', '-'), "Last Vendor": info.get('VENDOR', '-')})
        st.dataframe(pd.DataFrame(res_list), use_container_width=True)

# ==========================================
# MENU 3: E-CATALOG & STUDIO GAMBAR
# ==========================================
elif menu == "E-Catalog & Studio":
    st.markdown("<h2>🖼️ Enterprise Digital Catalog</h2>", unsafe_allow_html=True)
    t_cat, t_studio = st.tabs(["📖 Product Gallery", "🛠️ Asset Studio"])
    
    with t_cat:
        col_s, col_f = st.columns([2, 1])
        with col_s: 
            search_cat = st.text_input("🔍 Cari Produk:")
        with col_f:
            list_kat = ["All Categories"] + sorted([k for k in df_master_unique['KATEGORI'].unique() if str(k).strip() != ""])
            filter_cat = st.selectbox("📁 Kategori:", list_kat)
        
        df_show = df_master_unique.copy()
        if filter_cat != "All Categories": 
            df_show = df_show[df_show['KATEGORI'] == filter_cat]
        if search_cat: 
            df_show = df_show[df_show['NAMA BAKU'].astype(str).str.contains(search_cat, case=False) | df_show['NOMOR SKU'].astype(str).str.contains(search_cat, case=False)]
        
        st.markdown("---")
        
        if df_show.empty: 
            st.warning("Data tidak ditemukan.")
        else:
            cols = st.columns(4)
            for idx, (_, row) in enumerate(df_show.iterrows()):
                with cols[idx % 4]:
                    raw_link = str(row.get('LINK GAMBAR', '')).strip()
                    img_url = convert_gdrive_link(raw_link)
                    
                    # RAKIT FULL HTML CARD (Solusi kotak melayang)
                    if img_url and "drive.google" in img_url:
                        img_element = f"<img src='{img_url}' style='width:100%; height:160px; object-fit:contain; border-radius:8px; margin-bottom:12px;'>"
                    else:
                        img_element = f"<div style='background-color:#F1F5F9; height:160px; border-radius:8px; display:flex; align-items:center; justify-content:center; margin-bottom:12px;'><span style='color:#94A3B8; font-weight:600;'>No Image Asset</span></div>"
                    
                    card_html = f"""
                    <div style='background:white; border:1px solid #E2E8F0; border-radius:12px; padding:16px; margin-bottom:16px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); transition: 0.3s;'>
                        {img_element}
                        <h5 style='margin-top:0px; font-size:14px; font-weight:700; color:#0F172A; line-height:1.4; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;'>{row['NAMA BAKU']}</h5>
                        <p style='font-size:11px; color:#64748B; margin:4px 0;'>SKU: {row['NOMOR SKU']}</p>
                        <p style='font-size:15px; font-weight:800; color:#047857; margin-top:8px; margin-bottom:0px;'>{format_rupiah(row.get('HARGA', 0))}</p>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)

    with t_studio:
        st.write("### 📸 Inject Image Asset")
        df_no_pic = df_master_unique[df_master_unique['LINK GAMBAR'].astype(str).str.strip() == ""]
        if df_no_pic.empty: 
            st.success("Semua aset visual sudah lengkap.")
        else:
            barang_pilih = st.selectbox("Pilih Produk:", df_no_pic['NAMA BAKU'].tolist())
            link_input = st.text_input("G-Drive Link:")
            if link_input:
                st.image(convert_gdrive_link(link_input), width=300)
                if st.button("💾 Upload & Bind", type="primary"):
                    try:
                        with st.spinner("Binding asset..."):
                            client = get_gspread_client(); sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                            cell = sheet_master.find(barang_pilih, in_column=2)
                            if cell:
                                col_link_idx = sheet_master.row_values(1).index('LINK GAMBAR') + 1
                                sheet_master.update_cell(cell.row, col_link_idx, link_input)
                                st.success(f"Success!"); time.sleep(1); st.cache_data.clear(); st.rerun()
                    except Exception as e: st.error(f"Error: {e}")

# ==========================================
# MENU 4: DATABASE VENDOR
# ==========================================
elif menu == "Database Vendor":
    st.markdown("<h2>🏢 Supplier Directory</h2>", unsafe_allow_html=True)
    keyword = st.text_input("Cari Vendor / PIC / Item:")
    if keyword:
        try:
            df_v = load_data(GID_VENDOR); df_v.columns = df_v.columns.str.strip().str.upper()
            res = df_v[df_v.astype(str).apply(lambda x: x.str.contains(keyword, case=False)).any(axis=1)]
            for _, v in res.iterrows():
                with st.expander(f"🏢 {v.get('NAMA VENDOR', '-')} | Cat: {v.get('KATEGORI', '-')} | Level: {v.get('GRUP', '-')} "):
                    st.write(f"**Contact Person:** {v.get('PIC', '-')} 📞 {v.get('KONTAK', '-')}")
                    st.write(f"**Location Address:** {v.get('ALAMAT', '-')}")
        except: st.warning("Database Connection Error.")

# ==========================================
# MENU 5: DASHBOARD LAPORAN (ELEGANT SAAS UI)
# ==========================================
elif menu == "Dashboard Laporan":
    st.markdown("<h2>📊 Procurement Intelligence</h2>", unsafe_allow_html=True)
    
    try:
        client = get_gspread_client()
        data_dash = client.open_by_key(SHEET_ID).get_worksheet_by_id(int(GID_DASHBOARD)).get_all_values()
        df_v = load_data(GID_VENDOR); df_v.columns = df_v.columns.str.strip().str.upper()
        
        if len(data_dash) > 1:
            df_d = pd.DataFrame(data_dash[1:], columns=data_dash[0])
            df_d.columns = df_d.columns.str.strip().str.upper()
            
            c_po = next((c for c in df_d.columns if 'PO' in c or 'BUKTI' in c), None)
            c_unit = next((c for c in df_d.columns if 'UNIT' in c or 'GRUP' in c), None)
            c_harga = next((c for c in df_d.columns if 'HARGA' in c), None)
            c_baku = next((c for c in df_d.columns if 'BAKU' in c), None)
            c_tgl = next((c for c in df_d.columns if 'TANGGAL' in c or 'TGL' in c or 'DATE' in c), None)
            
            df_d['H_NUM'] = pd.to_numeric(df_d[c_harga].astype(str).str.upper().str.replace('RP', '').str.replace(r'[^0-9]', '', regex=True), errors='coerce').fillna(0)
            df_d['Q_NUM'] = pd.to_numeric(df_d['QTY'].astype(str).str.replace(r'[^0-9.]', '', regex=True), errors='coerce').fillna(0)
            df_d['TOTAL'] = df_d['H_NUM'] * df_d['Q_NUM']
            df_d['DATE_CLEAN'] = pd.to_datetime(df_d[c_tgl], errors='coerce')
            df_d = df_d.dropna(subset=['DATE_CLEAN'])
            
            tab_summary, tab_item = st.tabs(["🌐 Corporate Overview", "🔎 Item Analytics"])
            
            with tab_summary:
                list_unit = ["All Facilities"] + sorted([u for u in df_d[c_unit].unique() if str(u).strip() != ""])
                filter_unit = st.selectbox("📍 Select Facility:", list_unit)
                df_filtered = df_d[df_d[c_unit] == filter_unit] if filter_unit != "All Facilities" else df_d
                
                st.write("") # Spacer
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Procurement Value", format_rupiah(df_filtered['TOTAL'].sum()))
                col2.metric("PO Transactions", f"{df_filtered[c_po].replace('', pd.NA).dropna().nunique()}")
                col3.metric("Active Supply Facilities", f"{df_filtered[c_unit].nunique()}")
                st.write("")
                
                c_a, c_b = st.columns([1, 1.5])
                with c_a:
                    st.markdown("<h4 style='font-size:16px; color:#334155; margin-bottom:15px;'>Budget Distribution</h4>", unsafe_allow_html=True)
                    if filter_unit == "All Facilities":
                        rekap_u = df_filtered.groupby(c_unit)['TOTAL'].sum().reset_index()
                        rekap_u = rekap_u[rekap_u[c_unit].str.strip() != ""] 
                        fig_pie = px.pie(rekap_u, names=c_unit, values='TOTAL', hole=0.6, color_discrete_sequence=['#047857', '#10B981', '#34D399', '#6EE7B7'])
                        fig_pie.update_traces(textposition='inside', textinfo='percent')
                        fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                        st.plotly_chart(fig_pie, use_container_width=True)
                    else: st.info(f"Viewing specialized data for **{filter_unit}**.")

                with c_b:
                    st.markdown("<h4 style='font-size:16px; color:#334155; margin-bottom:15px;'>Top Procurement Items</h4>", unsafe_allow_html=True)
                    df_valid = df_filtered[~df_filtered[c_baku].str.contains('CEK MANUAL|BARANG BARU', case=False, na=False)]
                    if not df_valid.empty:
                        df_valid = df_valid[df_valid[c_baku].str.strip() != ""]
                        top_i = df_valid.groupby(c_baku)[c_po].nunique().reset_index()
                        top_i.columns = ['Nama Barang', 'Jumlah PO']
                        top_i = top_i.sort_values(by='Jumlah PO', ascending=False).head(8)
                        fig_bar = px.bar(top_i, x='Jumlah PO', y='Nama Barang', orientation='h', text='Jumlah PO', color_discrete_sequence=['#047857'])
                        fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, margin=dict(t=0, b=0, l=0, r=0), xaxis_title=None, yaxis_title=None, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig_bar, use_container_width=True)
            
            with tab_item:
                list_barang_histori = df_d.drop_duplicates(subset=[c_baku]).sort_values(by=c_baku)[c_baku].tolist()
                barang_pilih = st.selectbox("Search Product Intelligence:", list_barang_histori, placeholder="Type to search...")
                
                if barang_pilih:
                    info_master = df_master_unique[df_master_unique['NAMA BAKU'] == barang_pilih]
                    df_item_histori = df_d[df_d[c_baku] == barang_pilih].sort_values(by='DATE_CLEAN')
                    
                    kat_item = "NAN"
                    if not info_master.empty: kat_item = str(info_master.iloc[0].get('KATEGORI', '')).upper()

                    v_histori = sorted([str(v).strip() for v in df_item_histori['VENDOR'].unique() if str(v).strip() not in ['', '-', 'nan']])
                    v_database = []
                    if kat_item != "NAN" and not df_v.empty:
                        v_match = df_v[df_v['KATEGORI'].astype(str).str.contains(kat_item, case=False, na=False)]
                        v_database = sorted(v_match['NAMA VENDOR'].unique().tolist())

                    st.markdown("<br>", unsafe_allow_html=True)
                    c_img, c_meta = st.columns([1, 2.5])
                    with c_img:
                        if not info_master.empty:
                            img_url = convert_gdrive_link(str(info_master.iloc[0].get('LINK GAMBAR', '')).strip())
                            if img_url: st.markdown(f"<div style='border:1px solid #E2E8F0; border-radius:12px; padding:10px; background:white;'><img src='{img_url}' width='100%' style='border-radius:8px;'></div>", unsafe_allow_html=True)
                            else: st.info("🚫 No Asset")
                    with c_meta:
                        st.markdown(f"<h3 style='margin-top:0; color:#0F172A;'>{barang_pilih}</h3>", unsafe_allow_html=True)
                        if not info_master.empty:
                            row_m = info_master.iloc[0]
                            st.markdown(f"<p style='color:#64748B; font-weight:600; margin-bottom:15px;'>SKU: <span style='color:#047857;'>{row_m.get('NOMOR SKU', '-')}</span> &nbsp;|&nbsp; CAT: {kat_item} &nbsp;|&nbsp; UOM: {row_m.get('SATUAN', '-')}</p>", unsafe_allow_html=True)
                            
                            st.markdown("**🏭 Histori Supplier (Telah Digunakan):**")
                            st.markdown(f"<div style='background-color:#F8FAFC; border:1px solid #E2E8F0; padding:10px; border-radius:8px; font-size:14px;'>{', '.join(v_histori) if v_histori else 'Belum ada transaksi'}</div>", unsafe_allow_html=True)
                            
                            st.markdown("<br>**💡 Rekomendasi Supplier (Database Match):**", unsafe_allow_html=True)
                            st.markdown(f"<div style='background-color:#ECFDF5; border:1px solid #A7F3D0; padding:10px; border-radius:8px; font-size:14px; color:#065F46;'>{', '.join(v_database) if v_database else 'Tidak ada referensi di database'}</div>", unsafe_allow_html=True)
                    
                    st.markdown("<hr style='border:1px solid #E2E8F0; margin: 30px 0;'>", unsafe_allow_html=True)
                    
                    if not df_item_histori.empty:
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Item TCO (Total Cost)", format_rupiah(df_item_histori['TOTAL'].sum()))
                        m2.metric("Purchase Frequency", f"{df_item_histori[c_po].nunique()} Orders")
                        m3.metric("Average Unit Price", format_rupiah(df_item_histori['H_NUM'].mean()))
                        m4.metric("Supplier Count", f"{len(v_histori)} Vendors")

                        st.write("")
                        g_harga, g_qty = st.columns(2)
                        with g_harga:
                            fig_harga = px.line(df_item_histori, x='DATE_CLEAN', y='H_NUM', title="Price Volatility Trend", markers=True, color_discrete_sequence=['#F59E0B'])
                            fig_harga.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Price (IDR)")
                            st.plotly_chart(fig_harga, use_container_width=True)
                        with g_qty:
                            df_monthly = df_item_histori.resample('M', on='DATE_CLEAN').sum().reset_index()
                            fig_qty = px.bar(df_monthly, x='DATE_CLEAN', y='Q_NUM', title="Procurement Volume by Month", color_discrete_sequence=['#3B82F6'])
                            fig_qty.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="", yaxis_title="Quantity")
                            st.plotly_chart(fig_qty, use_container_width=True)
                        
                        st.markdown("<h4 style='font-size:16px; color:#334155; margin-bottom:10px;'>Transaction Ledger</h4>", unsafe_allow_html=True)
                        df_table = df_item_histori[[c_tgl, c_po, c_unit, 'VENDOR', 'QTY', 'H_NUM', 'TOTAL']].copy()
                        df_table['H_NUM'] = df_table['H_NUM'].map(format_rupiah)
                        df_table['TOTAL'] = df_table['TOTAL'].map(format_rupiah)
                        st.dataframe(df_table, use_container_width=True, hide_index=True)

        else: st.warning("Data Repository Empty.")
            
    except Exception as e: st.error(f"Engine Fault: {e}")

# ==========================================
# MENU 6: MAINTENANCE DATA
# ==========================================
elif menu == "Maintenance Data":
    st.markdown("<h2>🛠️ System Config & SKU Generator</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#64748B;'>Modul untuk injeksi SKU secara masif dan perawatan data.</p>", unsafe_allow_html=True)
    
    invalid_mask = df_master_unique['NOMOR SKU'].isna() | (df_master_unique['NOMOR SKU'].astype(str).str.strip().str.len() < 10)
    df_missing = df_master_unique[invalid_mask]
    
    if not df_missing.empty:
        st.warning(f"⚠️ Terdeteksi {len(df_missing)} item yang membutuhkan Nomor SKU.")
        if st.button("🚀 Execute SKU Injection", type="primary"):
            with st.spinner("Processing..."):
                try:
                    client = get_gspread_client(); sheet_master = client.open_by_key(SHEET_ID).get_worksheet(0)
                    all_data = sheet_master.get_all_values(); headers = [str(h).strip().upper() for h in all_data[0]]
                    df_m = pd.DataFrame(all_data[1:], columns=headers)
                    
                    c_s = next((c for c in headers if 'SKU' in c), None); c_k = next((c for c in headers if 'KATEGORI' in c and 'DETAIL' not in c), None); c_d = next((c for c in headers if 'DETAIL' in c), None)
                    if c_s and c_k and c_d:
                        for idx, row in df_m.iterrows():
                            val = str(row[c_s]).strip()
                            if len(val) < 10 or val.upper() in ['NAN', 'NONE', 'NULL', '#N/A', '']: df_m.at[idx, c_s] = generate_new_sku("001", row[c_k], row[c_d], df_m)
                        sheet_master.clear(); sheet_master.update(values=[df_m.columns.tolist()] + df_m.values.tolist())
                        st.success("Injection Complete!"); time.sleep(1.5); st.rerun()
                except Exception as e: st.error(f"Error: {e}")
    else: st.success("✔️ Database Sehat. Semua SKU terverifikasi.")