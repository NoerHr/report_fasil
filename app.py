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
# 1. ULTRA MODERN UI CONFIGURATION
# ==========================================
st.set_page_config(page_title="Admin Kelas AI", layout="wide", page_icon="✨")

def local_css():
    st.markdown("""
    <style>
    /* IMPORT FONT (Google Fonts: Poppins) */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }

    /* 1. BACKGROUND UTAMA (Deep Space Gradient) */
    .stApp {
        background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
        background-attachment: fixed;
        color: #ffffff;
    }

    /* 2. SIDEBAR GLASS (Frosted) */
    section[data-testid="stSidebar"] {
        background-color: rgba(255, 255, 255, 0.03);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.05);
    }

    /* 3. GLASS CARD CONTAINER (The "Glosarium" Effect) */
    .glass-container {
        background: rgba(255, 255, 255, 0.05);
        border-radius: 20px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border: 1px solid rgba(255, 255, 255, 0.1);
        padding: 25px;
        margin-bottom: 25px;
        transition: transform 0.3s ease;
    }
    
    .glass-container:hover {
        border: 1px solid rgba(0, 210, 255, 0.3); /* Neon glow on hover */
        transform: translateY(-2px);
    }

    /* HEADER STYLES */
    h1 {
        font-weight: 700;
        background: -webkit-linear-gradient(eee, #333);
        -webkit-background-clip: text;
        text-shadow: 0 0 20px rgba(0, 210, 255, 0.5);
        color: white;
        margin-bottom: 20px;
    }
    h2, h3 {
        color: #e0e0e0;
        font-weight: 600;
    }

    /* 4. INPUT FIELDS (Dark Glass) */
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        background-color: rgba(0, 0, 0, 0.3) !important;
        color: #fff !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        border-radius: 12px !important;
        padding: 10px;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #00d2ff !important;
        box-shadow: 0 0 10px rgba(0, 210, 255, 0.2);
    }

    /* 5. BUTTONS (Gradient Neon) */
    div.stButton > button {
        background: linear-gradient(90deg, #00d2ff 0%, #3a7bd5 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 1.5rem;
        letter-spacing: 1px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 210, 255, 0.3);
        width: 100%;
    }
    div.stButton > button:hover {
        transform: scale(1.02);
        box-shadow: 0 6px 20px rgba(0, 210, 255, 0.6);
        background: linear-gradient(90deg, #3a7bd5 0%, #00d2ff 100%);
    }

    /* 6. DATAFRAME & METRICS */
    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
        color: #00d2ff !important;
        text-shadow: 0 0 10px rgba(0, 210, 255, 0.3);
    }
    .stDataFrame {
        background-color: rgba(0, 0, 0, 0.2);
        border-radius: 10px;
        padding: 5px;
    }
    
    /* CUSTOM SCROLLBAR */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    ::-webkit-scrollbar-track {
        background: rgba(255, 255, 255, 0.05); 
    }
    ::-webkit-scrollbar-thumb {
        background: rgba(255, 255, 255, 0.2); 
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: rgba(255, 255, 255, 0.4); 
    }
    </style>
    """, unsafe_allow_html=True)

local_css()

# ==========================================
# 2. LOGIC (BACKEND - TETAP SAMA & KUAT)
# ==========================================

@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'], gpu=False)

def extract_text_from_image(image_file):
    reader = load_ocr()
    img = Image.open(image_file)
    result = reader.readtext(np.array(img), detail=0)
    cleaned = []
    ignore = ["participants", "chat", "share", "record", "host", "me", "mute", "unmute", "video"]
    for text in result:
        t = text.strip()
        if len(t) > 3 and not any(w in t.lower() for w in ignore):
            t = re.sub(r'^\d+\.?\s*', '', t)
            cleaned.append(t)
    return "\n".join(cleaned)

def load_master_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            try: df = pd.read_csv(uploaded_file)
            except: df = pd.read_csv(uploaded_file, sep=';')
        else:
            df = pd.read_excel(uploaded_file)
        df.columns = [c.strip().title() for c in df.columns]
        return df
    except: return None

def parse_data_template(text):
    data = {}
    jam_match = re.search(r'(\d{1,2}[\.:]\d{2})\s?-\s?(\d{1,2}[\.:]\d{2})', text)
    if jam_match:
        data['jam_mulai'] = jam_match.group(1).replace(':', '.') 
        data['jam_full'] = jam_match.group(0).replace('.', ':')
        parts = text.split(jam_match.group(0))
        
        # Fasil logic (Remove day name)
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
        data['jam_mulai'] = "00.00"
        data['jam_full'] = "00:00 - 00:00"
        data['fasilitator'] = "Fasilitator"
        sisa = text

    pertemuan_match = re.search(r'Pertemuan\s?([\d\s&,-]+)', sisa, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "1"

    kode_match = re.search(r'([A-Za-z]{2,}\d{1,3})', sisa)
    data['kode_kelas'] = kode_match.group(1) if kode_match else "KODE"

    types_found = []
    if re.search(r'Reguler|Reg', text, re.IGNORECASE): types_found.append("Reguler")
    if re.search(r'Profesional|Pro', text, re.IGNORECASE): types_found.append("Profesional")
    if re.search(r'Akselerasi|Aksel', text, re.IGNORECASE): types_found.append("Akselerasi")
    data['tipe_list'] = types_found if types_found else ["Reguler"]

    if data['kode_kelas'] in sisa:
        parts_matkul = sisa.split(data['kode_kelas'])
        data['matkul'] = parts_matkul[0].strip()
        raw_dosen = parts_matkul[1]
        clean_dosen = re.split(r'(\d{2}\sBisDig|Pro|Reg|Sulawesi|Pertemuan)', raw_dosen, flags=re.IGNORECASE)[0]
        data['dosen'] = clean_dosen.strip().strip(",").strip()
    else:
        data['matkul'] = "Matkul"
        data['dosen'] = "Dosen"
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
# 4. USER INTERFACE (PREMIUM DASHBOARD)
# ==========================================

st.title("✨ Admin Kelas AI")
st.markdown("### *Automation Dashboard for Facilitators*")

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Data Input")
    
    mode_input = st.radio("Sumber Data:", ["✍️ Manual / Paste", "📂 Database Excel"])
    defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"], "fasil": "Fasil"}
    
    if mode_input == "✍️ Manual / Paste":
        raw_template = st.text_area("Paste Jadwal Gancet:", height=70, help="Contoh: SeninMika08.30...")
        if raw_template:
            parsed = parse_data_template(raw_template)
            defaults.update({k:v for k,v in parsed.items() if k in defaults})
            if 'fasilitator' in parsed: defaults['fasil'] = parsed['fasilitator']
            
    elif mode_input == "📂 Database Excel":
        uploaded_db_jadwal = st.file_uploader("Upload Jadwal (.xlsx)", type=['xlsx'])
        if uploaded_db_jadwal:
            try:
                df_jadwal = pd.read_excel(uploaded_db_jadwal)
                df_jadwal['Label'] = df_jadwal.apply(lambda x: f"{x.get('Fasilitator', '')} - {x.get('Kode Kelas','?')} - {x.get('Mata Kuliah','?')}", axis=1)
                pilihan = st.selectbox("Pilih Kelas:", df_jadwal['Label'])
                row = df_jadwal[df_jadwal['Label'] == pilihan].iloc[0]
                defaults.update({
                    'matkul': str(row.get('Mata Kuliah', '')), 'dosen': str(row.get('Nama Dosen', '')),
                    'kode': str(row.get('Kode Kelas', '')), 'fasil': str(row.get('Fasilitator', '')),
                    'jam_mulai': str(row.get('Jam Mulai', '00.00')).replace(':', '.'),
                    'jam_full': f"{str(row.get('Jam Mulai', '')).replace('.', ':')} - Selesai"
                })
                defaults['tipe'] = [t.strip() for t in str(row.get('Tipe', 'Reguler')).split(',')]
            except: st.error("Database Error")

    st.markdown("---")
    st.subheader("📝 Edit Detail")
    
    inp_tgl = st.text_input("📅 Tanggal", datetime.now().strftime("%d %B %Y"))
    inp_fasil = st.text_input("👤 Fasilitator", value=defaults['fasil'])
    inp_matkul = st.text_input("📚 Mata Kuliah", value=defaults['matkul'])
    inp_dosen = st.text_input("👨‍🏫 Dosen", value=defaults['dosen'])
    
    c1, c2 = st.columns(2)
    inp_jam_dot = c1.text_input("⏰ Jam (08.00)", value=defaults['jam_mulai'])
    inp_kode = c2.text_input("🏷️ Kode (F4)", value=defaults['kode'])
    
    inp_tipe = st.multiselect("🎓 Tipe", ["Reguler", "Profesional", "Akselerasi"], default=defaults['tipe'])
    inp_tipe_str = " & ".join(inp_tipe) if inp_tipe else "Reguler"

    c3, c4 = st.columns(2)
    inp_pertemuan = c3.text_input("🔢 Sesi (1 & 2)", value="1")
    inp_fee = c4.number_input("💰 Fee", value=150000)

    info = {
        "tgl": inp_tgl, "matkul": inp_matkul, "dosen": inp_dosen, "jam_dot": inp_jam_dot, 
        "jam_full": defaults['jam_full'], "kode": inp_kode, 
        "pertemuan": inp_pertemuan, "tipe": inp_tipe_str, "fasil": inp_fasil, "lokasi": "Online"
    }

# --- BAGIAN 1: GENERATOR NAMA FILE ---
st.markdown('<div class="glass-container">', unsafe_allow_html=True)
st.subheader("1️⃣ Generator Nama File Otomatis")
pert_fmt = f"Pertemuan {inp_pertemuan}".replace("&", "dan")

col_h1, col_h2 = st.columns(2)
with col_h1:
    st.info("🅰️ Format Request Zoom")
    file_req = f"{inp_matkul}_{pert_fmt}_{inp_tgl}_{inp_tipe_str}_{inp_dosen}_{inp_jam_dot}"
    st.code(file_req, language="text")
with col_h2:
    st.success("🅱️ Format Bukti Foto (Rename Foto Anda!)")
    file_bukti = f"{inp_tgl}_{inp_dosen}_{inp_matkul}_{inp_fasil}_{pert_fmt}"
    st.code(f"{file_bukti}.jpg", language="text")
st.markdown('</div>', unsafe_allow_html=True)

# --- BAGIAN 2: UPLOAD AREA ---
st.markdown('<div class="glass-container">', unsafe_allow_html=True)
st.subheader("2️⃣ Upload Data Presensi")

col_ocr, col_file = st.columns(2)
with col_ocr:
    st.markdown("📸 **Screenshot Zoom (OCR)**")
    up_foto = st.file_uploader("Upload Foto Zoom (JPG/PNG)", type=['jpg','png'])
    ocr_res = ""
    if up_foto:
        with st.spinner("🔍 AI sedang membaca teks..."):
            try: ocr_res = extract_text_from_image(up_foto)
            except: pass
    txt_zoom = st.text_area("List Nama (Hasil Scan):", value=ocr_res, height=120)
    
    st.markdown("---")
    st.markdown("🙋‍♂️ **Peserta Onsite / Manual**")
    txt_onsite = st.text_area("Ketik nama manual disini (1 nama per baris):", height=80)

with col_file:
    st.markdown("📂 **Database & Feedback**")
    file_master = st.file_uploader("Master Mahasiswa (.xlsx/.csv)", type=['xlsx', 'csv'])
    file_feedback = st.file_uploader("Feedback (.csv)", type=['csv'])

st.markdown('</div>', unsafe_allow_html=True)

# --- PROCESS BUTTON ---
if st.button("🚀 PROSES ANALISIS & DOWNLOAD", type="primary"):
    if (not txt_zoom and not txt_onsite) or not file_master:
        st.error("❌ Data Peserta dan Master Mahasiswa belum lengkap!")
        st.stop()
    
    # 1. LOAD MASTER
    db = load_master_data(file_master)
    if db is None:
        st.error("Gagal membaca file Master. Coba file lain.")
        st.stop()
        
    col_nm = next((c for c in db.columns if 'nama' in c.lower()), None)
    if not col_nm:
        st.error("Kolom 'Nama' tidak ditemukan di Excel Master.")
        st.stop()
    db_names = db[col_nm].astype(str).tolist()
    
    # 2. MATCHING (ZOOM + ONSITE)
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
        try:
            df_fb = pd.read_csv(file_feedback)
            col_fb = [c for c in df_fb.columns if "Nama" in c][0]
            col_sesi = next((c for c in df_fb.columns if "Pertemuan" in c or "Sesi" in c), None)
            targets = [str(s) for s in get_session_list(inp_pertemuan)]
            for _, r in df_fb.iterrows():
                valid = True
                if col_sesi:
                    if not any(t in str(r[col_sesi]) for t in targets): valid = False
                if valid:
                    m, _ = get_best_match_info(str(r[col_fb]), db_names)
                    if m: final_fb.append(m)
        except: pass

    # 4. STATS
    uh, uf = list(set(final_hadir)), list(set(final_fb))
    fb_ok, fb_no = 0, []
    for h in uh:
        if h in uf: fb_ok += 1
        else: fb_no.append(h)
    
    stats = {'total':len(db), 'hadir':len(uh), 'fb_ok':fb_ok, 'fb_no':len(fb_no), 'pct':round(fb_ok/len(uh)*100,1) if uh else 0}

    # --- DASHBOARD HASIL ---
    st.markdown("---")
    st.subheader("📊 Hasil Analisis")
    
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Hadir", f"{len(uh)}/{len(db)}")
    m2.metric("Feedback OK", f"{fb_ok}")
    m3.metric("Tanpa Feedback", len(fb_no), delta_color="inverse")
    m4.metric("Estimasi Gaji", f"Rp {inp_fee * (len(get_session_list(inp_pertemuan)) or 1):,.0f}")

    if ambig: st.warning("⚠️ Perhatian: Ada nama ambigu yang dipilih sistem."); st.dataframe(ambig)
    if fb_no: 
        with st.expander("📢 Klik untuk lihat Mahasiswa Belum Feedback (Copy untuk Tagih)"):
            st.code("\n".join(fb_no))

    # --- DOWNLOAD SECTION ---
    st.markdown("### 📥 Download Laporan (Excel)")
    t1, t2, t3, t4 = st.tabs(["✅ Sprint Checklist", "📝 Laporan Fasil", "🎓 Presensi Fakultas", "💰 Tabungan Gaji"])
    
    with t1:
        df = generate_laporan_sprint(info, stats)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Checklist_Sprint.xlsx")
    with t2:
        df = generate_laporan_fasilitator(info, stats, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name=f"Laporan_{info['tgl']}.xlsx")
    with t3:
        df = generate_presensi(db, uh, uf, info)
        def color(v): return 'background-color:#c6efce' if v=='O' else 'background-color:#ffc7ce' if 'F' in str(v) else ''
        cols = [c for c in df.columns if "Nama" in c or "nama" in c] + [f'Sesi {i}' for i in get_session_list(inp_pertemuan)]
        st.dataframe(df[cols].style.applymap(color), use_container_width=True)
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Presensi.xlsx")
    with t4:
        df = generate_gaji(info, inp_fee, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Download Excel", data=to_excel_download(df), file_name="Gaji.xlsx")