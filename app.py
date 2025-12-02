import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
import difflib
import easyocr
import numpy as np
from PIL import Image

# ==========================================
# 1. UI CONFIGURATION (CYBER GLASS THEME)
# ==========================================
st.set_page_config(page_title="Admin Kelas Pro Max", layout="wide", page_icon="⚡")

def local_css():
    st.markdown("""
    <style>
    /* IMPORT GOOGLE FONTS */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }

    /* BACKGROUND ANIMASI */
    .stApp {
        background-color: #000000;
        background-image: 
            radial-gradient(at 0% 0%, hsla(253,16%,7%,1) 0, transparent 50%), 
            radial-gradient(at 50% 0%, hsla(225,39%,30%,1) 0, transparent 50%), 
            radial-gradient(at 100% 0%, hsla(339,49%,30%,1) 0, transparent 50%);
        color: white;
    }

    /* GLASS CARD CONTAINER */
    .glass-box {
        background: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 25px;
        margin-bottom: 20px;
        transition: all 0.3s ease-in-out;
    }
    
    .glass-box:hover {
        border: 1px solid rgba(0, 255, 200, 0.5);
        box-shadow: 0 0 20px rgba(0, 255, 200, 0.2);
        transform: translateY(-2px);
    }

    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: rgba(10, 10, 20, 0.8);
        border-right: 1px solid rgba(255,255,255,0.05);
    }

    /* INPUT FIELDS */
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        background-color: rgba(0, 0, 0, 0.4) !important;
        border: 1px solid rgba(255, 255, 255, 0.15) !important;
        color: white !important;
        border-radius: 10px !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #00ffc8 !important;
    }

    /* BUTTON STYLING */
    div.stButton > button {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
        color: #004d40;
        font-weight: 800;
        border: none;
        border-radius: 12px;
        padding: 0.8rem 2rem;
        width: 100%;
        text-transform: uppercase;
        letter-spacing: 1px;
        transition: 0.3s;
    }
    div.stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 0 30px rgba(0, 201, 255, 0.6);
        color: white;
    }

    /* METRICS */
    div[data-testid="stMetricValue"] {
        font-size: 2.5rem !important;
        background: -webkit-linear-gradient(#00C9FF, #92FE9D);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
    }
    
    /* TABS */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: rgba(255,255,255,0.05);
        border-radius: 10px;
        color: white;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: rgba(0, 255, 200, 0.2) !important;
        color: #00ffc8 !important;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

local_css()

# ==========================================
# 2. FUNGSI LOGIKA (SMART PARSING & LOADING)
# ==========================================

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

def extract_text_from_image(image_file):
    reader = load_ocr()
    img = Image.open(image_file)
    result = reader.readtext(np.array(img), detail=0)
    cleaned = []
    ignore = ["participants", "chat", "share", "record", "host", "me", "mute", "unmute"]
    for text in result:
        t = text.strip()
        if len(t) > 3 and not any(w in t.lower() for w in ignore):
            t = re.sub(r'^\d+\.?\s*', '', t)
            cleaned.append(t)
    return "\n".join(cleaned)

def load_data_smart(uploaded_file):
    """Smart Loader: Membaca Excel/CSV dengan berbagai delimiter"""
    try:
        # Cek Ekstensi
        if uploaded_file.name.endswith('.csv'):
            try:
                # Coba titik koma dulu (format file Anda)
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=';')
                if len(df.columns) < 2: raise Exception
            except:
                # Fallback ke koma
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=',')
        else:
            # Excel
            df = pd.read_excel(uploaded_file)
            
        # Bersihkan nama kolom
        df.columns = [str(c).strip().title() for c in df.columns]
        return df
    except Exception as e:
        return None

def parse_data_template(text):
    data = {}
    
    # 1. Parsing Jam
    jam_match = re.search(r'(\d{1,2}[\.:]\d{2})\s?-\s?(\d{1,2}[\.:]\d{2})', text)
    if jam_match:
        data['jam_mulai'] = jam_match.group(1).replace(':', '.') 
        data['jam_full'] = jam_match.group(0).replace('.', ':')
        parts = text.split(jam_match.group(0))
        
        # Pisahkan Hari dari Fasilitator
        raw_prefix = parts[0].strip()
        days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        fasil_clean = raw_prefix
        for day in days:
            if fasil_clean.lower().startswith(day.lower()):
                fasil_clean = fasil_clean[len(day):].strip()
                break
        data['fasilitator'] = fasil_clean if fasil_clean else "Fasilitator"
        
        sisa = parts[1] if len(parts) > 1 else text
    else:
        data['jam_mulai'] = "00.00"; data['jam_full'] = "00:00 - 00:00"; data['fasilitator'] = "Fasilitator"; sisa = text

    # 2. Parsing Kode Kelas
    kode_match = re.search(r'([A-Za-z]{2,}\d{1,3})', sisa)
    if kode_match:
        data['kode_kelas'] = kode_match.group(1)
        parts_kode = sisa.split(data['kode_kelas'])
        data['matkul'] = parts_kode[0].strip()
        
        # Parsing Dosen
        raw_dosen = parts_kode[1]
        clean_dosen = re.split(r'(\d{2}[A-Za-z]|\d{2}\s|Reg|Pro|Sulawesi|Bali|Java|Sumatera|Papua|Pertemuan)', raw_dosen, flags=re.IGNORECASE)[0]
        data['dosen'] = clean_dosen.strip().strip(",").strip()
    else:
        data['kode_kelas'] = "KODE"; data['matkul'] = "Matkul"; data['dosen'] = "Dosen"

    # 3. Parsing Pertemuan & Tipe
    pertemuan_match = re.search(r'Pertemuan\s?([\d\s&,-]+)', text, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "1"

    types = []
    if re.search(r'Reguler|Reg', text, re.IGNORECASE): types.append("Reguler")
    if re.search(r'Profesional|Pro', text, re.IGNORECASE): types.append("Profesional")
    data['tipe_str'] = " & ".join(types) if types else "Reguler"

    return data

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
    clean_str = pertemuan_str.replace('&', ',').replace('-', ',').replace('dan', ',')
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

# GENERATOR LAPORAN
def generate_laporan_sprint(info, stats):
    tasks = ["Reminder H-1", "Dosen Hadir", "Host Claim", "Absensi", "Recording", "Upload GDrive", "Laporan Sistem", "Update Feedback"]
    return pd.DataFrame({"No": range(1,len(tasks)+1), "Task": tasks, "Status": ["v"]*len(tasks), "Ket": ["Done"]*len(tasks)})

def generate_laporan_fasilitator(info, stats, filename):
    summary = f"Kelas: {info['kode']}\nJam: {info['jam_full']}\nSesi: {info['pertemuan']}\nTerdaftar: {stats['total']}\nHadir: {stats['hadir']}\nFeedback: {stats['fb_ok']} | Belum: {stats['fb_no']}"
    return pd.DataFrame({"Tanggal": [info['tgl']], "Nama Dosen": [info['dosen']], "Mata Kuliah": [info['matkul']], "Jam": [info['jam_full']], "Tipe": ["Online"], "Sesi": [info['pertemuan']], "Bukti": [filename], "Validasi": ["Valid"], "Feedback Summary": [summary], "Persentase": [f"{stats['pct']}%"]})

def generate_presensi(db, hadir, fb, info):
    df = db.copy()
    col = [c for c in df.columns if "Nama" in c][0]
    for i in range(1, 17): df[f'Sesi {i}'] = ""
    target = get_session_list(info['pertemuan'])
    for idx, r in df.iterrows():
        n = str(r[col])
        code = "A"
        if n in hadir: code = "O" if n in fb else "OF"
        for s in target: 
            if 1 <= s <= 16: df.at[idx, f'Sesi {s}'] = code
    return df

def generate_gaji(info, fee, filename):
    jml = len(get_session_list(info['pertemuan'])) or 1
    return pd.DataFrame([{"Tanggal": info['tgl'], "Dosen": info['dosen'], "Matkul": info['matkul'], "Kode": info['kode'], "Sesi": info['pertemuan'], "Jml": jml, "Fee": fee, "Total": fee*jml, "Bukti": filename}])

# ==========================================
# 4. USER INTERFACE (DASHBOARD)
# ==========================================

st.markdown("# ⚡ Admin Kelas Pro")
st.markdown("### *Automation Dashboard v3.0*")

# --- SIDEBAR: CONTROL PANEL ---
with st.sidebar:
    st.markdown("## 🎛️ Control Panel")
    
    # MODE SELECTION
    mode_input = st.radio("Sumber Data:", ["✍️ Manual Paste", "📂 Database Excel"])
    defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"], "fasil": "Fasil"}
    
    if mode_input == "✍️ Manual Paste":
        raw_template = st.text_area("Jadwal Raw (Gancet):", height=80, 
                                    help="Contoh: SeninMika08.30...",
                                    placeholder="Paste disini...")
        if raw_template:
            parsed = parse_data_template(raw_template)
            defaults.update({k:v for k,v in parsed.items() if k in defaults})
            if 'fasilitator' in parsed: defaults['fasil'] = parsed['fasilitator']
            
    elif mode_input == "📂 Database Excel":
        up_db = st.file_uploader("Upload Jadwal (.xlsx)", type=['xlsx'])
        if up_db:
            df_db = load_data_smart(up_db)
            if df_db is not None:
                col_fasil = next((c for c in df_db.columns if 'fasil' in c.lower()), '')
                col_kode = next((c for c in df_db.columns if 'kode' in c.lower()), '')
                col_matkul = next((c for c in df_db.columns if 'mata' in c.lower()), '')
                
                if col_kode:
                    df_db['Label'] = df_db.apply(lambda x: f"{x.get(col_fasil, '')} - {x.get(col_matkul,'?')}", axis=1)
                    pilihan = st.selectbox("Pilih Kelas:", df_db['Label'])
                    row = df_db[df_db['Label'] == pilihan].iloc[0]
                    defaults.update({
                        'matkul': str(row.get(col_matkul, '')), 'dosen': str(row.get('Nama Dosen', '')),
                        'kode': str(row.get(col_kode, '')), 'fasil': str(row.get(col_fasil, '')) if col_fasil else '',
                        'jam_mulai': str(row.get('Jam Mulai', '00.00')).replace(':', '.'),
                        'jam_full': f"{str(row.get('Jam Mulai', '')).replace('.', ':')} - Selesai"
                    })
                    defaults['tipe'] = [t.strip() for t in str(row.get('Tipe', 'Reguler')).split(',')]

    st.markdown("---")
    st.markdown("### 📝 Edit Detail")
    
    inp_tgl = st.text_input("📅 Tanggal", datetime.now().strftime("%d %B %Y"))
    inp_fasil = st.text_input("👤 Fasilitator", value=defaults['fasil'])
    inp_matkul = st.text_input("📚 Mata Kuliah", value=defaults['matkul'])
    inp_dosen = st.text_input("👨‍🏫 Dosen", value=defaults['dosen'])
    
    c1, c2 = st.columns(2)
    inp_jam_dot = c1.text_input("⏰ Jam (08.00)", value=defaults['jam_mulai'])
    inp_kode = c2.text_input("🏷️ Kode", value=defaults['kode'])
    
    inp_tipe = st.multiselect("🎓 Tipe", ["Reguler", "Profesional", "Akselerasi"], default=defaults['tipe'])
    inp_tipe_str = " & ".join(inp_tipe) if inp_tipe else "Reguler"

    c3, c4 = st.columns(2)
    inp_pertemuan = c3.text_input("🔢 Sesi (1,2)", value="1")
    inp_fee = c4.number_input("💰 Fee", value=150000)

    info = {
        "tgl": inp_tgl, "matkul": inp_matkul, "dosen": inp_dosen, "jam_dot": inp_jam_dot, 
        "jam_full": defaults['jam_full'], "kode": inp_kode, 
        "pertemuan": inp_pertemuan, "tipe": inp_tipe_str, "fasil": inp_fasil, "lokasi": "Online"
    }

# --- SECTION 1: SMART GENERATOR ---
pert_fmt = f"Pertemuan {inp_pertemuan}".replace("&", "dan")
file_req = f"{inp_matkul}_{pert_fmt}_{inp_tgl}_{inp_tipe_str}_{inp_dosen}_{inp_jam_dot}"
file_bukti = f"{inp_tgl}_{inp_dosen}_{inp_matkul}_{inp_fasil}_{pert_fmt}"

st.markdown('<div class="glass-box">', unsafe_allow_html=True)
st.markdown("### 1️⃣ Generator Nama File")
c1, c2 = st.columns(2)
with c1:
    st.info("🅰️ Format Request Zoom")
    st.code(file_req, language="text")
with c2:
    st.success("🅱️ Format Rename Foto")
    st.code(f"{file_bukti}.jpg", language="text")
st.markdown('</div>', unsafe_allow_html=True)

# --- SECTION 2: UPLOAD ---
st.markdown('<div class="glass-box">', unsafe_allow_html=True)
st.markdown("### 2️⃣ Upload Data Presensi")

col_ocr, col_file = st.columns([1, 1.2])

with col_ocr:
    st.markdown("**📸 Screenshot Zoom (AI OCR)**")
    up_foto = st.file_uploader("Upload Foto Zoom (JPG/PNG)", type=['jpg','png'])
    ocr_res = ""
    if up_foto:
        with st.spinner("🔍 AI Sedang Membaca Teks..."):
            try: ocr_res = extract_text_from_image(up_foto)
            except: pass
    txt_zoom = st.text_area("List Nama (Hasil Scan):", value=ocr_res, height=120)
    
    st.markdown("**🙋‍♂️ Peserta Onsite / Manual**")
    txt_onsite = st.text_area("Input Manual (1 nama per baris):", height=80)

with col_file:
    st.markdown("**📂 Database Mahasiswa & Feedback**")
    st.warning("⚠️ Support Excel (.xlsx) dan CSV (.csv)")
    file_master = st.file_uploader("Master Mahasiswa", type=['xlsx', 'csv'])
    file_feedback = st.file_uploader("Data Feedback", type=['xlsx', 'xls', 'csv'])

st.markdown('</div>', unsafe_allow_html=True)

# --- ACTION BUTTON ---
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    process_btn = st.button("🚀 JALANKAN ANALISIS", type="primary")

# --- PROCESS ---
if process_btn:
    if (not txt_zoom and not txt_onsite) or not file_master:
        st.error("❌ Data tidak lengkap! Mohon isi data peserta dan upload Master Mahasiswa.")
        st.stop()
    
    # 1. MASTER
    db = load_data_smart(file_master)
    if db is None: st.error("Gagal baca Master."); st.stop()
    col_nm = next((c for c in db.columns if 'nama' in c.lower()), None)
    if not col_nm: st.error("Kolom 'Nama' tidak ditemukan di Master."); st.stop()
    db_names = db[col_nm].astype(str).tolist()
    
    # 2. MATCHING
    raw_names = [clean_nama_zoom(x) for x in (txt_zoom.split('\n') + txt_onsite.split('\n')) if len(x)>3]
    final_hadir, ambig = [], []
    for z in raw_names:
        best, conf = get_best_match_info(z, db_names)
        if best:
            final_hadir.append(best)
            if len(conf)>1: ambig.append({"Input":z, "Pilih":best})

    # 3. FEEDBACK
    final_fb = []
    if file_feedback:
        df_fb = load_data_smart(file_feedback)
        if df_fb is not None:
            col_fb = next((c for c in df_fb.columns if 'nama' in c.lower()), None)
            col_sesi = next((c for c in df_fb.columns if 'pertemuan' in c.lower() or 'sesi' in c.lower()), None)
            targets = [str(s) for s in get_session_list(inp_pertemuan)]
            if col_fb:
                for _, r in df_fb.iterrows():
                    valid = True
                    if col_sesi and not any(t in str(r[col_sesi]) for t in targets): valid = False
                    if valid:
                        m, _ = get_best_match_info(str(r[col_fb]), db_names)
                        if m: final_fb.append(m)

    uh, uf = list(set(final_hadir)), list(set(final_fb))
    fb_ok, fb_no = 0, []
    for h in uh:
        if h in uf: fb_ok += 1
        else: fb_no.append(h)
    
    stats = {'total':len(db), 'hadir':len(uh), 'fb_ok':fb_ok, 'fb_no':len(fb_no), 'pct':round(fb_ok/len(uh)*100,1) if uh else 0}

    # --- DASHBOARD HASIL ---
    st.markdown("---")
    st.write("### 📊 Dashboard Hasil")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Hadir", f"{len(uh)}/{len(db)}")
    m2.metric("Feedback OK", f"{fb_ok}")
    m3.metric("Tanpa Feedback", len(fb_no), delta_color="inverse")
    m4.metric("Estimasi Gaji", f"Rp {inp_fee * (len(get_session_list(inp_pertemuan)) or 1):,.0f}")

    if ambig: st.warning("⚠️ Ada nama ambigu yang dipilih sistem"); st.dataframe(ambig)
    if fb_no: 
        with st.expander("📢 Klik untuk lihat Mahasiswa Belum Feedback (Copy untuk Tagih)"):
            st.code("\n".join(fb_no))

    # DOWNLOAD & COPY PASTE
    st.write("### 📥 Download & Copy Laporan")
    t1, t2, t3, t4 = st.tabs(["✅ Sprint Checklist", "📝 Laporan Fasil", "🎓 Presensi Fakultas", "💰 Tabungan Gaji"])
    
    with t1:
        df = generate_laporan_sprint(info, stats)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.code(df.to_csv(sep='\t', index=False), language='text')
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Sprint.xlsx")
    with t2:
        df = generate_laporan_fasilitator(info, stats, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.code(df.to_csv(sep='\t', index=False), language='text')
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name=f"Laporan_{info['tgl']}.xlsx")
    with t3:
        df = generate_presensi(db, uh, uf, info)
        def color(v): return 'background-color:#004d40' if v=='O' else 'background-color:#4a1414' if 'F' in str(v) else ''
        cols = [c for c in df.columns if "Nama" in c or "nama" in c] + [f'Sesi {i}' for i in get_session_list(inp_pertemuan)]
        st.dataframe(df[cols].style.applymap(color), use_container_width=True)
        st.code(df[cols].to_csv(sep='\t', index=False), language='text')
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Presensi.xlsx")
    with t4:
        df = generate_gaji(info, inp_fee, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.code(df.to_csv(sep='\t', index=False), language='text')
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Gaji.xlsx")

st.markdown('</div>', unsafe_allow_html=True)