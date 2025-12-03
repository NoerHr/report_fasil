import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
import difflib
import numpy as np
from PIL import Image

# ==========================================
# 1. UI CONFIGURATION
# ==========================================
st.set_page_config(page_title="Admin Kelas Pro Max", layout="wide", page_icon="💎")

def local_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&display=swap');
    html, body, [class*="css"] { font-family: 'Outfit', sans-serif; }
    .stApp { background-color: #09090b; color: #f8fafc; }
    .glass-box {
        background: rgba(255, 255, 255, 0.03); border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 16px; padding: 24px; margin-bottom: 20px;
    }
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        background-color: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        color: white !important; border-radius: 8px !important;
    }
    div.stButton > button {
        background: linear-gradient(135deg, #0ea5e9 0%, #3b82f6 100%);
        color: white; font-weight: 700; border: none; border-radius: 8px;
        padding: 0.6rem 1.2rem; width: 100%; transition: 0.3s;
    }
    div.stButton > button:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(14, 165, 233, 0.4); }
    </style>
    """, unsafe_allow_html=True)
local_css()

# ==========================================
# 2. LOGIC & PARSING
# ==========================================

@st.cache_resource
def get_ocr_reader():
    import easyocr
    return easyocr.Reader(['en'], gpu=False)

def extract_text_from_image(image_file):
    try:
        reader = get_ocr_reader()
        img = Image.open(image_file)
        result = reader.readtext(np.array(img), detail=0)
        cleaned = []
        ignore = ["participants", "chat", "share", "record", "host", "me", "mute", "unmute"]
        for text in result:
            t = text.strip()
            if len(t) > 3 and not any(w in t.lower() for w in ignore):
                t = re.sub(r'^\d+[\.\)]?\s*', '', t)
                cleaned.append(t)
        return "\n".join(cleaned)
    except: return ""

def load_data_smart(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            try: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, sep=';')
            except: uploaded_file.seek(0); df = pd.read_csv(uploaded_file, sep=',')
        else: df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip().title() for c in df.columns]
        return df
    except: return None

def clean_nama_zoom(nama_raw):
    if not isinstance(nama_raw, str): return ""
    if any(x in nama_raw.lower() for x in ["fasil", "host", "admin"]): return "IGNORE"
    nama = re.sub(r'^[\d\-\_\.]+\s*', '', nama_raw) 
    nama = re.sub(r'[\-\_]\s*[A-Z]{2,3}$', '', nama) 
    return nama.strip().title()

def get_best_match_info(nama_zoom, list_db_names):
    nama_zoom = nama_zoom.lower()
    db_lower_map = {name.lower(): name for name in list_db_names}
    for db_low, db_real in db_lower_map.items():
        if nama_zoom in db_low or db_low in nama_zoom:
            conflicts = [db_lower_map[k] for k in db_lower_map if nama_zoom in k]
            return db_real, (conflicts if len(conflicts)>1 else [])
    matches = difflib.get_close_matches(nama_zoom, list_db_names, n=3, cutoff=0.6)
    if matches: return matches[0], (matches if len(matches)>1 else [])
    return None, []

def get_session_list(pertemuan_str):
    clean_str = str(pertemuan_str).replace('&', ',').replace('-', ',').replace('dan', ',').replace('Pertemuan', '')
    return [int(x) for x in clean_str.split(',') if x.strip().isdigit()]

def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, width)
    return output.getvalue()

def to_excel_multi_sheet(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            safe_name = re.sub(r'[\\/*?:\[\]]', '', str(sheet_name))[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            worksheet = writer.sheets[safe_name]
            for i, col in enumerate(df.columns):
                width = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, width)
    return output.getvalue()

def clean_matkul_smart(text):
    clean = re.sub(r'[\t\s]+', ' ', text).strip()
    match = re.match(r'^(.+?)\s*\1$', clean, re.IGNORECASE)
    if match: return match.group(1) 
    return clean

# --- ENRICHMENT LOGIC (AUTO FILL KODE KELAS) ---
def normalize_jam(text):
    match = re.search(r'(\d{1,2})[:.](\d{2})', str(text))
    if match: return f"{int(match.group(1)):02d}:{match.group(2)}"
    return ""

def enrich_with_db(df_batch, df_db):
    if df_db is None or df_batch is None: return df_batch
    
    # Keyword flexible
    key_mtk = ['mata', 'matkul', 'course', 'subject', 'mk']
    key_dos = ['dosen', 'pengajar', 'lecturer', 'instructor', 'nama dosen']
    key_jam = ['jam', 'waktu', 'time', 'pukul', 'schedule']
    key_kod = ['kode', 'code', 'class id', 'kd']

    col_mtk_db = next((c for c in df_db.columns if any(k in c.lower() for k in key_mtk)), None)
    col_dos_db = next((c for c in df_db.columns if any(k in c.lower() for k in key_dos)), None)
    col_jam_db = next((c for c in df_db.columns if any(k in c.lower() for k in key_jam)), None)
    col_kod_db = next((c for c in df_db.columns if any(k in c.lower() for k in key_kod)), None)
    
    if not col_mtk_db or not col_kod_db: return df_batch

    db_records = df_db.to_dict('records')
    
    def get_kode_smart(row):
        if row['Kode Kelas'] not in ["N/A", ""]: return row['Kode Kelas']
        tgt_mtk = str(row['Mata Kuliah']).lower().strip()
        tgt_dos = str(row['Nama Dosen']).lower().strip()
        tgt_jam = normalize_jam(row['Jam'])
        
        best_score = 0; best_kode = "N/A"
        for record in db_records:
            db_mtk = str(record.get(col_mtk_db, '')).lower().strip()
            db_dos = str(record.get(col_dos_db, '')).lower().strip() if col_dos_db else ""
            db_jam = normalize_jam(record.get(col_jam_db, '')) if col_jam_db else ""
            
            score_mtk = difflib.SequenceMatcher(None, tgt_mtk, db_mtk).ratio()
            score_dos = 0
            if col_dos_db:
                score_dos = difflib.SequenceMatcher(None, tgt_dos, db_dos).ratio()
                if tgt_dos in db_dos or db_dos in tgt_dos: score_dos = 1.0 
            
            score_jam = 1.0 if col_jam_db and tgt_jam and db_jam and tgt_jam == db_jam else 0.0
            
            # Bobot Skor
            if col_dos_db and col_jam_db: final_score = (score_mtk * 0.4) + (score_dos * 0.3) + (score_jam * 0.3)
            elif col_dos_db: final_score = (score_mtk * 0.6) + (score_dos * 0.4)
            else: final_score = score_mtk

            if final_score > best_score and final_score > 0.65: # Threshold 65%
                best_score = final_score
                best_kode = record[col_kod_db]
        return best_kode

    df_batch['Kode Kelas'] = df_batch.apply(get_kode_smart, axis=1)
    return df_batch

# --- PARSERS ---
def parse_data_template(text):
    data = {}
    jam_match = re.search(r'(\d{1,2}[\.:]\d{2})\s?-\s?(\d{1,2}[\.:]\d{2})', text)
    if jam_match:
        data['jam_mulai'] = jam_match.group(1).replace(':', '.')
        data['jam_full'] = jam_match.group(0).replace('.', ':')
        parts = text.split(jam_match.group(0))
        raw_prefix = parts[0].strip()
        days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        fasil_clean = raw_prefix
        for day in days:
            if fasil_clean.lower().startswith(day.lower()):
                fasil_clean = fasil_clean[len(day):].strip(); break
        data['fasil'] = fasil_clean if fasil_clean else "Fasil"
        sisa = parts[1] if len(parts) > 1 else text
    else:
        data['jam_mulai'] = "00.00"; data['jam_full'] = "00:00 - 00:00"; data['fasil'] = "Fasil"; sisa = text

    kode_match = re.search(r'([A-Za-z]{2,}\d{1,3})', sisa)
    if kode_match:
        data['kode'] = kode_match.group(1)
        parts_kode = sisa.split(data['kode'])
        data['matkul'] = clean_matkul_smart(parts_kode[0].strip())
        raw_dosen = parts_kode[1] if len(parts_kode) > 1 else ""
        clean_dosen = re.split(r'(\d{2}[A-Za-z]|\d{2}\s|Reg|Pro|Sulawesi|Bali|Java|Sumatera|Papua|Pertemuan|\d+\s?STI)', raw_dosen, flags=re.IGNORECASE)[0]
        data['dosen'] = clean_dosen.strip().strip(",").strip()
    else:
        data['kode'] = "KODE"; data['matkul'] = "Matkul"; data['dosen'] = "Dosen"

    pertemuan_match = re.search(r'(Pertemuan\s?[\d\s&,-]+)', text, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "Pertemuan 1"
    
    types = []
    if re.search(r'Reguler|Reg', text, re.IGNORECASE): types.append("Reguler")
    if re.search(r'Profesional|Pro', text, re.IGNORECASE): types.append("Profesional")
    data['tipe_str'] = types[0] if types else "Reguler"
    return data

def parse_random_batch_text(raw_text):
    import re
    import pandas as pd
    clean_text = re.sub(r'\s*_\s*', '_', raw_text)
    clean_text = re.sub(r'_(\d{1,2}[\.:]\d{2})\s*\([Pp][Aa][Rr][Tt]\s*(\d+)\)', r'_Pertemuan \2_\1', clean_text, flags=re.IGNORECASE)
    raw_parts = re.split(r'(_\d{1,2}[\.:]\d{2})', clean_text.strip())
    entries = []
    for i in range(0, len(raw_parts) - 1, 2):
        full_row = raw_parts[i] + raw_parts[i+1]
        if len(full_row) > 10: entries.append(full_row)
    
    parsed_data = []
    for entry in entries:
        try:
            parts = entry.split('_')
            jam = parts[-1].strip().replace('.', ':').zfill(5)
            dosen = parts[-2].strip()
            val_min_3 = parts[-3].strip()
            if re.search(r'(Reguler|Profesional|Professional|International|Reg|Pro)', val_min_3, re.IGNORECASE):
                tipe = val_min_3; tgl_raw = parts[-4].strip(); sesi_raw = parts[-5].strip(); front_blob = parts[:-5] 
            else:
                tipe = "Reguler"; tgl_raw = val_min_3; sesi_raw = parts[-4].strip(); front_blob = parts[:-4]

            front_text = "_".join(front_blob)
            days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
            fasil = "Fasil"; temp_text = front_text
            for day in days:
                match = re.search(f"(.*?)({day})", front_text, re.IGNORECASE)
                if match: fasil = match.group(1).strip(); temp_text = front_text[match.end():]; break
            
            match_tahun = re.search(r'202\d', temp_text)
            matkul_messy = temp_text[match_tahun.end():].strip() if match_tahun else temp_text
            matkul = clean_matkul_smart(matkul_messy)

            sessions_found = re.findall(r'\d+', sesi_raw)
            if not sessions_found: sessions_found = ['1']
            req_zoom_combined = f"{matkul}_{sesi_raw}_{tgl_raw}_{tipe}_{dosen}_{jam}"

            for sess_num in sessions_found:
                clean_tgl = tgl_raw.replace(",", "")
                clean_dos = re.sub(r'[\\/*?:"<>|]', "", dosen)
                clean_mat = re.sub(r'[\\/*?:"<>|]', "", matkul)
                file_foto = f"{clean_tgl}_{clean_dos}_{clean_mat}_{fasil}_Pertemuan {sess_num}.jpg"
                parsed_data.append({
                    "Tanggal": tgl_raw, "Fasilitator": fasil, "Jam": jam,
                    "Kode Kelas": "N/A", "Mata Kuliah": matkul, "Nama Dosen": dosen,
                    "Tipe": tipe, "Sesi": sess_num, "Req Zoom": req_zoom_combined, "Nama File Foto": file_foto 
                })
        except Exception as e: pass
    return pd.DataFrame(parsed_data)

# --- STATS & EXCEL ---
def run_analysis(info, txt_zoom, txt_onsite, db_names, df_fb):
    # 1. Matching
    list_zoom = [clean_nama_zoom(x) for x in str(txt_zoom).split('\n') if len(x)>3]
    hadir_zoom = set()
    for z in list_zoom:
        best, _ = get_best_match_info(z, db_names)
        if best: hadir_zoom.add(best)

    list_onsite = [clean_nama_zoom(x) for x in str(txt_onsite).split('\n') if len(x)>3]
    hadir_onsite = set()
    for z in list_onsite:
        best, _ = get_best_match_info(z, db_names)
        if best: hadir_onsite.add(best)
    
    total_hadir = hadir_zoom.union(hadir_onsite)

    # 2. Feedback
    final_fb = set()
    ghosts = []
    if df_fb is not None:
        col_fb = next((c for c in df_fb.columns if 'nama' in str(c).lower()), None)
        col_sesi = next((c for c in df_fb.columns if 'pertemuan' in str(c).lower() or 'sesi' in str(c).lower()), None)
        targets = [str(s) for s in get_session_list(info['pertemuan'])]
        if col_fb:
            for _, r in df_fb.iterrows():
                valid = True
                if col_sesi:
                    csv_val = re.findall(r'\d+', str(r[col_sesi]))
                    if not set(targets).intersection(set(csv_val)): valid = False
                if valid:
                    m, _ = get_best_match_info(str(r[col_fb]), db_names)
                    if m: 
                        final_fb.add(m)
                        if m not in total_hadir: ghosts.append(m)
    
    fb_ok = 0; fb_no = []
    for h in total_hadir:
        if h in final_fb: fb_ok += 1
        else: fb_no.append(h)
    
    stats = {'total': len(db_names), 'hadir_valid': len(total_hadir), 'online_count': len(hadir_zoom), 
             'onsite_count': len(hadir_onsite), 'fb_ok': fb_ok, 'fb_no': len(fb_no), 
             'pct': round(fb_ok/len(total_hadir)*100,1) if total_hadir else 0, 'ghosts': ghosts, 'fb_no_list': fb_no}
    
    return stats, hadir_zoom, hadir_onsite, final_fb

def generate_presensi_real(db, hadir_zoom, hadir_onsite, list_feedback, info):
    df = db.copy()
    col_nm = next((c for c in df.columns if 'nama' in str(c).lower()), df.columns[1])
    col_nim = next((c for c in df.columns if 'nim' in str(c).lower()), df.columns[0])
    df_out = df[[col_nim, col_nm]].copy()
    df_out.columns = ['NIM', 'Nama Mahasiswa']
    target_sessions = get_session_list(info['pertemuan'])
    
    for idx, r in df_out.iterrows():
        n = str(r['Nama Mahasiswa'])
        code = "A" # Default Alpha
        if n in hadir_onsite: code = "S" if n in list_feedback else "SF"
        elif n in hadir_zoom: code = "O" if n in list_feedback else "OF"
        
        for s in target_sessions:
            if 1 <= int(s) <= 16: df_out.loc[idx, f"Sesi {s}"] = code
    return df_out

def generate_output_excel(info, stats, filename, sks, role):
    summary = f"Kelas : {info['kode']}\nJam : {info['jam_full']}\nTotal: {stats['total']}\nHadir: {stats['hadir_valid']}\nFeedback: {stats['fb_ok']}"
    return pd.DataFrame({
        "Tanggal": [info['tgl']], "Dosen": [info['dosen']], "Matkul": [info['matkul']], 
        "Jam": [info['jam_full']], "Lokasi": [info['tipe_belajar']], "SKS": [sks], 
        "Tipe": [info['tipe']], "Sesi": [info['pertemuan']], 
        "Req Zoom": [info.get('req_zoom', '-')], "Bukti": [filename],
        "Hadir": [stats['hadir_valid']], "Feedback %": [f"{stats['pct']}%"], "Summary": [summary]
    })

def generate_gaji(info, fee, filename):
    jml = len(get_session_list(info['pertemuan'])) or 1
    return pd.DataFrame([{"Tanggal": info['tgl'], "Dosen": info['dosen'], "Matkul": info['matkul'], "Sesi": info['pertemuan'], "Fee": fee, "Total": fee*jml}])

# ==========================================
# 3. MAIN UI
# ==========================================
st.markdown('<h1 class="main-header">⚡ Admin Kelas Pro Max</h1>', unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🎛️ Control Panel")
    app_mode = st.radio("Mode:", ["👤 Single", "🚀 Batch Process", "🛠️ Buat Template"])
    inp_tipe_belajar = st.selectbox("📍 Lokasi Belajar:", ["Online", "Onsite", "Hybrid"])
    
    st.markdown("---")
    st.markdown("### 📂 DB Jadwal (Opsional)")
    up_db_jadwal = st.file_uploader("Upload DB Jadwal (Excel)", type=['xlsx', 'csv'])
    
    if up_db_jadwal:
        st.session_state.db_jadwal = load_data_smart(up_db_jadwal)
        st.success(f"✅ DB Jadwal Loaded: {len(st.session_state.db_jadwal)} data")
    elif 'db_jadwal' not in st.session_state:
        st.session_state.db_jadwal = None

    st.markdown("---")
    inp_rem_h1 = st.text_input("⏰ H-1", "13.00"); inp_rem_h30 = st.text_input("⏰ H-30m", "07.30")
    inp_sks = st.number_input("SKS", 3); inp_fee = st.number_input("Fee", 150000); inp_role = st.text_input("Peran", "Fasilitator Kelas")

# --- MODE 1: SINGLE ---
if app_mode == "👤 Single":
    with st.expander("📝 Input Data Kelas", expanded=True):
        has_db = st.session_state.db_jadwal is not None
        options = ["✍️ Paste Text"]
        if has_db: options.append("▼ Pilih dari Database")
        
        inp_source = st.radio("Sumber Input:", options, horizontal=True)
        defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"], "fasil": "Fasil"}
        
        if inp_source == "✍️ Paste Text":
            raw = st.text_area("Paste Jadwal:", height=100)
            if raw: defaults.update(parse_data_template(raw))
            if st.session_state.db_jadwal is not None and defaults.get('matkul'):
                # Auto Enrich if DB is present
                df_temp = pd.DataFrame([{"Mata Kuliah": defaults['matkul'], "Nama Dosen": defaults['dosen'], "Jam": defaults['jam_full'], "Kode Kelas": "N/A"}])
                df_temp = enrich_with_db(df_temp, st.session_state.db_jadwal)
                defaults['kode'] = df_temp.iloc[0]['Kode Kelas']
                
        elif inp_source == "▼ Pilih dari Database":
            df_sch = st.session_state.db_jadwal
            c_mtk = next((c for c in df_sch.columns if 'mata' in c.lower() or 'matkul' in c.lower()), df_sch.columns[0])
            c_dos = next((c for c in df_sch.columns if 'dosen' in c.lower() or 'pengajar' in c.lower()), df_sch.columns[1])
            c_jam = next((c for c in df_sch.columns if 'jam' in c.lower() or 'waktu' in c.lower()), None)
            
            df_sch['Label_UI'] = df_sch[c_mtk].astype(str) + " - " + df_sch[c_dos].astype(str)
            if c_jam: df_sch['Label_UI'] += " [" + df_sch[c_jam].astype(str) + "]"
            
            pilihan = st.selectbox("Pilih Kelas:", df_sch['Label_UI'])
            row = df_sch[df_sch['Label_UI'] == pilihan].iloc[0]
            
            defaults['matkul'] = str(row[c_mtk])
            defaults['dosen'] = str(row[c_dos])
            defaults['kode'] = str(row.get(next((c for c in df_sch.columns if 'kode' in c.lower()), ''), ''))
            if c_jam: 
                raw_jam = str(row[c_jam])
                defaults['jam_full'] = raw_jam
                match_j = re.search(r'(\d{1,2})[:.](\d{2})', raw_jam)
                if match_j: defaults['jam_mulai'] = f"{match_j.group(1)}:{match_j.group(2)}"

        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            i_tgl = st.text_input("Tanggal", defaults.get('jam_mulai', datetime.now().strftime("%d %B %Y")))
            i_matkul = st.text_input("Matkul", value=defaults.get('matkul',''))
            i_dosen = st.text_input("Dosen", value=defaults.get('dosen',''))
            i_fasil = st.text_input("Fasil", value=defaults.get('fasil',''))
            i_kode = st.text_input("Kode", value=defaults.get('kode',''))
        with c2:
            i_jam = st.text_input("Jam", value=defaults.get('jam_full',''))
            i_tipe = st.text_input("Tipe", value=defaults.get('tipe_str','Reguler'))
            i_sesi = st.text_input("Sesi (String)", value=defaults.get('pertemuan_str','Pertemuan 1'))
            
            req_zoom_out = f"{i_matkul}_{i_sesi}_{i_tgl}_{i_tipe}_{i_dosen}_{i_jam}"
            clean_tgl = str(i_tgl).replace(",", "")
            clean_dos = re.sub(r'[\\/*?:"<>|]', "", str(i_dosen))
            clean_mat = re.sub(r'[\\/*?:"<>|]', "", str(i_matkul))
            sesi_num = re.search(r'\d+', str(i_sesi)).group(0) if re.search(r'\d+', str(i_sesi)) else "1"
            file_foto_out = f"{clean_tgl}_{clean_dos}_{clean_mat}_{i_fasil}_Pertemuan {sesi_num}.jpg"

    st.info("📋 **Output Helper**")
    c_out1, c_out2 = st.columns(2)
    c_out1.text_input("🅰️ Req Zoom", value=req_zoom_out)
    c_out2.text_input("🅱️ Nama File Foto", value=file_foto_out)

    c_ocr, c_file = st.columns([1, 1.2])
    with c_ocr:
        st.write("📸 **Bukti Kehadiran**")
        up_img = st.file_uploader("Upload Foto", type=['jpg','png'])
        txt_zoom = st.text_area("List Zoom (OCR):", value=extract_text_from_image(up_img) if up_img else "", height=100)
        txt_onsite = st.text_area("List Onsite (Manual):", height=80)
    with c_file:
        st.write("📂 **Database**")
        up_master = st.file_uploader("Master Mhs", type=['xlsx','csv']); up_fb = st.file_uploader("Feedback", type=['xlsx','csv'])

    if st.button("🚀 PROSES DATA"):
        if not up_master: st.error("Master Wajib!"); st.stop()
        info = {"tgl":i_tgl, "matkul":i_matkul, "dosen":i_dosen, "kode":i_kode, "jam_full":i_jam, "pertemuan":i_sesi, "tipe":i_tipe, "tipe_belajar":inp_tipe_belajar, "rem_h1":inp_rem_h1, "rem_h30":inp_rem_h30, "req_zoom": req_zoom_out}
        db = load_data_smart(up_master); db_names = db[next((c for c in db.columns if 'nama' in c.lower()), db.columns[1])].astype(str).tolist()
        df_fb_data = load_data_smart(up_fb) if up_fb else None
        
        stats, hadir_zoom, hadir_onsite, final_fb = run_analysis(info, txt_zoom, txt_onsite, db_names, df_fb_data)
        
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("Hadir", stats['hadir_valid']); c2.metric("No Feedback", stats['fb_no']); c3.metric("Fee", f"Rp {inp_fee * len(get_session_list(i_sesi)):,.0f}")
        
        if stats['fb_no_list']: 
            with st.expander("📢 Belum Feedback"): st.code("\n".join(stats['fb_no_list']))

        df_out = generate_output_excel(info, stats, file_foto_out, inp_sks, inp_role)
        st.download_button("Download Laporan", to_excel_download(df_out), f"Laporan_{i_kode}.xlsx")
        
        with st.expander("🎓 Download Detail Absensi", expanded=True):
            df_presensi = generate_presensi_real(db, hadir_zoom, hadir_onsite, final_fb, info)
            st.dataframe(df_presensi)
            st.download_button("Download Presensi.xlsx", to_excel_download(df_presensi), "Presensi.xlsx")

# --- MODE 2: BATCH ---
elif app_mode == "🚀 Batch Process":
    batch_src = st.radio("Sumber:", ["Paste Text", "Upload Excel"], horizontal=True)
    if batch_src == "Paste Text":
        raw_batch = st.text_area("Paste Data:", height=100)
        if st.button("Parse"): 
            df_parsed = parse_random_batch_text(raw_batch)
            if st.session_state.db_jadwal is not None:
                df_parsed = enrich_with_db(df_parsed, st.session_state.db_jadwal)
                st.toast("✅ Auto-Fill Kode Kelas Selesai!")
            st.session_state.batch_df = df_parsed
    else:
        up_batch = st.file_uploader("Upload Excel", type=['xlsx'])
        if up_batch: st.session_state.batch_df = load_data_smart(up_batch)

    if 'batch_df' in st.session_state:
        df_proc = st.data_editor(st.session_state.batch_df, num_rows="dynamic")
        st.write("### 📤 Pendukung")
        c1, c2 = st.columns(2)
        up_imgs = c1.file_uploader("Foto Bukti", type=['jpg','png'], accept_multiple_files=True)
        up_master = c2.file_uploader("Master Mhs", type=['xlsx']); up_fb = c2.file_uploader("Feedback", type=['xlsx'])
        
        if st.button("⚡ JALANKAN BATCH") and up_master and up_imgs:
            db = load_data_smart(up_master); db_names = db[next((c for c in db.columns if 'nama' in c.lower()), db.columns[1])].astype(str).tolist()
            df_fb_data = load_data_smart(up_fb) if up_fb else None
            img_map = {f.name: f for f in up_imgs}
            
            res = []; res_gaji = []; batch_presensi_dfs = {}
            prog = st.progress(0)
            
            for idx, row in df_proc.iterrows():
                def g(k, d=''): return str(row.get(next((c for c in df_proc.columns if k.lower() in c.lower()), None), d))
                
                info = {"tgl":g('tanggal'), "matkul":g('mata'), "dosen":g('dosen'), "kode":g('kode','N/A'), "jam_full":g('jam'), 
                        "pertemuan":g('sesi'), "tipe":g('tipe'), "tipe_belajar":inp_tipe_belajar, "rem_h1":inp_rem_h1, "rem_h30":inp_rem_h30, "req_zoom": g('req zoom')}
                t_img = g('foto', g('file'))
                
                txt_zoom = extract_text_from_image(img_map[t_img]) if t_img in img_map else ""
                
                stats, h_zoom, h_onsite, f_fb = run_analysis(info, txt_zoom, "", db_names, df_fb_data)
                
                res.append(generate_output_excel(info, stats, t_img, inp_sks, inp_role))
                res_gaji.append(generate_gaji(info, inp_fee, t_img))
                
                df_pres = generate_presensi_real(db, h_zoom, h_onsite, f_fb, info)
                sheet_id = f"{idx+1}_{info['matkul'][:15]}_{info['pertemuan']}"
                batch_presensi_dfs[sheet_id] = df_pres
                
                prog.progress((idx+1)/len(df_proc))
            
            st.success("Selesai!")
            c1, c2, c3 = st.columns(3)
            c1.download_button("📥 Laporan Gabungan", to_excel_download(pd.concat(res)), "Batch_Laporan.xlsx")
            c2.download_button("📥 Rekap Gaji", to_excel_download(pd.concat(res_gaji)), "Batch_Gaji.xlsx")
            c3.download_button("📥 Detail Presensi (Multi-Sheet)", to_excel_multi_sheet(batch_presensi_dfs), "Batch_Presensi.xlsx")

# --- MODE 3: TEMPLATE ---
elif app_mode == "🛠️ Buat Template":
    st.write("Paste teks gancet untuk bikin Excel Batch siap pakai.")
    raw = st.text_area("Text:"); 
    if st.button("Convert"):
        df = parse_random_batch_text(raw)
        st.dataframe(df)
        st.download_button("Download Excel", to_excel_download(df), "Template_Batch.xlsx")