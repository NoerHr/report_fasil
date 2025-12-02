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
# 1. UI CONFIGURATION (MODERN GLASS THEME)
# ==========================================
st.set_page_config(page_title="Admin Kelas Pro Max", layout="wide", page_icon="💎")

def local_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&display=swap');
    html, body, [class*="css"] { font-family: 'Outfit', sans-serif; }
    
    .stApp {
        background-color: #09090b;
        background-image: 
            radial-gradient(at 0% 0%, rgba(56, 189, 248, 0.15) 0, transparent 50%), 
            radial-gradient(at 100% 100%, rgba(139, 92, 246, 0.15) 0, transparent 50%);
        color: #f8fafc;
    }

    /* GLASS CARD */
    .glass-box {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 24px -1px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 20px;
    }

    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: rgba(10, 10, 15, 0.8);
        border-right: 1px solid rgba(255,255,255,0.05);
    }

    /* INPUTS */
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        background-color: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border-radius: 8px !important;
    }

    /* BUTTONS */
    div.stButton > button {
        background: linear-gradient(135deg, #0ea5e9 0%, #3b82f6 100%);
        color: white;
        font-weight: 700;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
        width: 100%;
        transition: 0.3s;
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(14, 165, 233, 0.4);
    }
    
    /* TABLES */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 8px;
    }
    </style>
    """, unsafe_allow_html=True)

local_css()

# ==========================================
# 2. FUNGSI LOGIKA (BACKEND)
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
    try:
        if uploaded_file.name.endswith('.csv'):
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=';')
                if len(df.columns) < 2: raise Exception
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, sep=',')
        else:
            df = pd.read_excel(uploaded_file)
        
        # Normalisasi Kolom: Title Case & Strip
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
    clean_str = str(pertemuan_str).replace('&', ',').replace('-', ',').replace('dan', ',')
    return [int(x) for x in clean_str.split(',') if x.strip().isdigit()]

def to_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for i, col in enumerate(df.columns):
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, width)
        
        # Format Wrap Text untuk kolom panjang (Feedback Summary)
        workbook = writer.book
        wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        if 'Form Feedback' in df.columns:
            idx_fb = df.columns.get_loc('Form Feedback')
            worksheet.set_column(idx_fb, idx_fb, 50, wrap_format)

    return output.getvalue()

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
        data['fasilitator'] = fasil_clean if fasil_clean else "Fasilitator"
        sisa = parts[1] if len(parts) > 1 else text
    else:
        data['jam_mulai'] = "00.00"; data['jam_full'] = "00:00 - 00:00"; data['fasilitator'] = "Fasilitator"; sisa = text

    kode_match = re.search(r'([A-Za-z]{2,}\d{1,3})', sisa)
    if kode_match:
        data['kode_kelas'] = kode_match.group(1)
        parts_kode = sisa.split(data['kode_kelas'])
        data['matkul'] = parts_kode[0].strip()
        raw_dosen = parts_kode[1]
        clean_dosen = re.split(r'(\d{2}[A-Za-z]|\d{2}\s|Reg|Pro|Sulawesi|Bali|Java|Sumatera|Papua|Pertemuan)', raw_dosen, flags=re.IGNORECASE)[0]
        data['dosen'] = clean_dosen.strip().strip(",").strip()
    else:
        data['kode_kelas'] = "KODE"; data['matkul'] = "Matkul"; data['dosen'] = "Dosen"

    pertemuan_match = re.search(r'Pertemuan\s?([\d\s&,-]+)', text, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "1"
    
    types = []
    if re.search(r'Reguler|Reg', text, re.IGNORECASE): types.append("Reguler")
    if re.search(r'Profesional|Pro', text, re.IGNORECASE): types.append("Profesional")
    data['tipe_str'] = " & ".join(types) if types else "Reguler"
    return data

# ==========================================
# 3. GENERATOR LAPORAN SESUAI FORMAT USER
# ==========================================

def generate_laporan_utama_sesuai_format(info, stats, filename, sks_val, role_val):
    """
    Menghasilkan laporan sesuai struktur 'format.xlsx'
    """
    # Format Text Summary untuk kolom "Form Feedback" (Penting!)
    text_summary = (
        f"Kelas : {info['kode']}\n"
        f"Jam : {info['jam_full']}\n"
        f"Pertemuan : {info['pertemuan']}\n"
        f"Tanggal : {info['tgl']}\n"
        f"1. Total Mahasiswa Terdaftar : {stats['total']}\n"
        f"2. Jumlah Mahasiswa Hadir : {stats['hadir']}\n"
        f"3. Jumlah Mengisi Form Feedback: {stats['fb_ok']}\n"
        f"4. Jumlah belum mengisi form feedback: {stats['fb_no']}\n"
        f"5. Jumlah tidak hadir : {stats['total'] - stats['hadir']}"
    )

    data = {
        "Tanggal Kehadiran": [info['tgl']],
        "Nama Dosen": [info['dosen']],
        "Nama Mata Kuliah": [info['matkul']],
        "Jam Perkuliahan": [info['jam_full']],
        "SKS": [sks_val], # User Input
        "Tipe Kelas": [info['tipe']],
        "Sesi Kelas": [info['pertemuan']],
        "Peran": [role_val], # Default: Fasilitator Kelas
        "Bukti Kehadiran\n(link ss zoom kelas)": [filename],
        "Validasi Bukti Hadir": ["Valid"],
        "Reminder H-1": ["Done"],
        "Reminder H-30 menit": ["Done"],
        "Form Feedback": [text_summary],
        "Waktu Report": [datetime.now().strftime("%H:%M")],
        "Waktu kirim PDF Feedback": ["-"],
        "Pemenuhan Feedback": [f"{stats['pct']}%"],
        "Final Absen": ["Ready"],
        "Jumlah Online": [stats['online_count']],
        "Jumlah On-Site": [stats['onsite_count']]
    }
    return pd.DataFrame(data)

def generate_checklist_sprint(info):
    """
    Menghasilkan checklist sesuai 'format (1).xlsx'
    """
    tasks = [
        "Reminder informasi kelas kepada mahasiswa",
        "Reminder informasi kelas kepada dosen",
        "Pastikan ke dosen apakah materi dishare sendiri atau dibantu",
        "Jika kelas cancel, jangan lupa update di Informasi Fasilitator",
        "Jika kelas direschedule, jangan lupa update di spreadsheet",
        "Jika fasil berhalangan, silakan infokan ke group",
        "Matikan proyektor, laptop, AC dan peralatan lainnya (Onsite)",
        "Kembalikan alat-alat pendukung ke bagian akademik (Onsite)",
        "Reminder presensi attendance dan form feedback saat kelas berjalan",
        "Update informasi terkait kelas (recording Zoom, dll)",
        "Mengisi rekap kehadiran fasilitator",
        "Update informasi jumlah pengisi form feedback",
        "Mengirimkan file PDF feedback dosen",
        "Rekap presensi kehadiran kelas",
        "Follow-up mahasiswa yang belum melakukan pengisian feedback"
    ]
    
    col_name = f"{info['kode']} - Pertemuan {info['pertemuan']}"
    
    df = pd.DataFrame({
        "Task": tasks,
        col_name: ["TRUE"] * len(tasks) # Default TRUE semua
    })
    return df

def generate_chat_templates(info):
    """
    Menghasilkan template chat sesuai 'format (2).xlsx'
    """
    templates = []
    
    # 1. Perkenalan
    templates.append({
        "Context": "Perkenalan",
        "Text": f"Halo Bapak/Ibu, selamat pagi/siang/sore.\n\nPerkenalkan, saya {info['fasil']} sebagai fasilitator untuk kelas {info['kode']}, salam kenal ya Bapak/Ibu.\n\nSekaligus saya izin mengingatkan Bapak/Ibu untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJika ada yang ingin ditanyakan, silakan yaa Bapak/Ibu. Terima kasih"
    })
    
    # 2. Reminder Kelas (Mhs)
    templates.append({
        "Context": "Reminder Kelas",
        "Text": f"Halo teman-teman, selamat pagi/siang/sore/malam.\n\nSehubungan dengan jadwal perkuliahan yang berjalan, saya izin mengingatkan untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJangan sampai terlambat dan sampai jumpa di kelas👋"
    })
    
    # 3. Reminder Dosen
    templates.append({
        "Context": "Reminder Dosen",
        "Text": f"Halo Bapak/Ibu, selamat pagi/siang/sore.\n\nSaya izin mengingatkan untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJIka ada yang bisa dibantu atau yang ingin ditanyakan terkait keberlangsungan kelas nanti, boleh ya Bapak/Ibu. Terima kasih🙏"
    })
    
    # 4. Kirim Feedback
    templates.append({
        "Context": "Kirim Feedback",
        "Text": f"Halo teman-teman,\n\njangan lupa untuk mengisi link feedback kelas sebagai bukti presensi pada link berikut yaa:\n\n🔗 [LINK FEEDBACK]\n\nTerima kasih dan sampai jumpa di kelas berikutya!"
    })
    
    return pd.DataFrame(templates)

def generate_presensi_real(db, hadir, fb, info):
    df = db.copy()
    # Cari kolom nama & nim dengan fleksibel
    col_nm = next((c for c in df.columns if 'nama' in str(c).lower()), df.columns[1])
    col_nim = next((c for c in df.columns if 'nim' in str(c).lower()), df.columns[0])
    
    # Buat DataFrame Baru yang bersih
    df_out = df[[col_nim, col_nm]].copy()
    df_out.columns = ['NIM', 'Nama Mahasiswa']
    
    target_sessions = get_session_list(info['pertemuan'])
    
    for idx, r in df_out.iterrows():
        n = str(r['Nama Mahasiswa'])
        code = "A"
        if n in hadir: code = "O" if n in fb else "OF" # O=Hadir+FB, OF=Hadir No FB
        
        # Isi Kolom Sesi (1-16)
        for s in target_sessions:
            if 1 <= int(s) <= 16:
                col_sesi = f"Sesi {s}"
                df_out.loc[idx, col_sesi] = code
                
    return df_out

def generate_gaji(info, fee, filename):
    jml = len(get_session_list(info['pertemuan'])) or 1
    return pd.DataFrame([{"Tanggal": info['tgl'], "Dosen": info['dosen'], "Matkul": info['matkul'], "Kode": info['kode'], "Sesi": info['pertemuan'], "Jml": jml, "Fee": fee, "Total": fee*jml, "Bukti": filename}])

# ==========================================
# 4. USER INTERFACE
# ==========================================

st.markdown('<h1 class="main-header">⚡ Admin Kelas Pro Max</h1>', unsafe_allow_html=True)
st.write("")

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("### 🎛️ Control Panel")
    
    mode_input = st.radio("Sumber Jadwal:", ["✍️ Manual Paste", "📂 Database Excel"])
    defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"], "fasil": "Fasil"}
    
    if mode_input == "✍️ Manual Paste":
        raw_template = st.text_area("Jadwal Raw (Gancet):", height=80, placeholder="Contoh: SeninMika08.30...")
        if raw_template:
            parsed = parse_data_template(raw_template)
            defaults.update({k:v for k,v in parsed.items() if k in defaults})
            if 'fasilitator' in parsed: defaults['fasil'] = parsed['fasilitator']
            
    elif mode_input == "📂 Database Excel":
        up_db = st.file_uploader("Upload Jadwal", type=['xlsx', 'xls', 'csv'])
        if up_db:
            df_db = load_data_smart(up_db)
            if df_db is not None:
                # Kolom fleksibel
                col_fasil = next((c for c in df_db.columns if 'fasil' in c.lower()), '')
                col_kode = next((c for c in df_db.columns if 'kode' in c.lower()), '')
                col_matkul = next((c for c in df_db.columns if 'mata' in c.lower()), '')
                
                col_label = next((c for c in df_db.columns if 'label' in c.lower()), '')
                if not col_label and col_kode:
                    df_db['Label_Gen'] = df_db.apply(lambda x: f"{x.get(col_fasil, '')} - {x.get(col_matkul,'?')}", axis=1)
                    col_label = 'Label_Gen'

                if col_label:
                    pilihan = st.selectbox("Pilih Kelas:", df_db[col_label])
                    row = df_db[df_db[col_label] == pilihan].iloc[0]
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
    
    # Input Tambahan Sesuai Format
    inp_sks = st.number_input("SKS", value=3)
    inp_role = st.text_input("Peran", value="Fasilitator Kelas")

    info = {
        "tgl": inp_tgl, "matkul": inp_matkul, "dosen": inp_dosen, "jam_dot": inp_jam_dot, 
        "jam_full": defaults['jam_full'], "kode": inp_kode, 
        "pertemuan": inp_pertemuan, "tipe": inp_tipe_str, "fasil": inp_fasil, "lokasi": "Online"
    }

# --- SECTION 1: GENERATOR ---
pert_fmt = f"Pertemuan {inp_pertemuan}".replace("&", "dan")
file_req = f"{inp_matkul}_{pert_fmt}_{inp_tgl}_{inp_tipe_str}_{inp_dosen}_{inp_jam_dot}"
file_bukti = f"{inp_tgl}_{inp_dosen}_{inp_matkul}_{inp_fasil}_{pert_fmt}"

st.markdown('<div class="glass-box">', unsafe_allow_html=True)
st.markdown("### 1️⃣ Generator Nama File")
c1, c2 = st.columns(2)
c1.info("🅰️ Req Zoom"); c1.code(file_req, language="text")
c2.success("🅱️ Bukti Foto"); c2.code(f"{file_bukti}.jpg", language="text")
st.markdown('</div>', unsafe_allow_html=True)

# --- SECTION 2: UPLOAD ---
st.markdown('<div class="glass-box">', unsafe_allow_html=True)
st.markdown("### 2️⃣ Upload Data")

col_ocr, col_file = st.columns([1, 1.2])

with col_ocr:
    st.markdown("**📸 Screenshot Zoom (AI OCR)**")
    up_foto = st.file_uploader("Upload Foto Zoom", type=['jpg','png'])
    ocr_res = ""
    if up_foto:
        with st.spinner("AI Reading..."):
            try: ocr_res = extract_text_from_image(up_foto)
            except: pass
    txt_zoom = st.text_area("List Online (Zoom):", value=ocr_res, height=120)
    st.markdown("**🙋‍♂️ Peserta Onsite / Manual**")
    txt_onsite = st.text_area("List Onsite (Manual):", height=80)

with col_file:
    st.markdown("**📂 Master & Feedback**")
    file_master = st.file_uploader("Master Mahasiswa", type=['xlsx', 'csv'])
    file_feedback = st.file_uploader("Feedback Data", type=['xlsx', 'xls', 'csv'])

st.markdown('</div>', unsafe_allow_html=True)

# --- PROCESS ---
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    process_btn = st.button("🚀 JALANKAN ANALISIS")

if process_btn:
    if (not txt_zoom and not txt_onsite) or not file_master:
        st.error("❌ Data tidak lengkap!"); st.stop()
    
    # 1. LOAD MASTER
    db = load_data_smart(file_master)
    if db is None: st.error("Gagal baca Master."); st.stop()
    
    col_nm = next((c for c in db.columns if 'nama' in str(c).lower()), None)
    if not col_nm: st.error("Kolom 'Nama' tidak ditemukan di Master."); st.stop()
    db_names = db[col_nm].astype(str).tolist()
    
    # 2. MATCHING (Separate Online vs Onsite for Report)
    list_online_raw = [clean_nama_zoom(x) for x in txt_zoom.split('\n') if len(x)>3]
    list_onsite_raw = [clean_nama_zoom(x) for x in txt_onsite.split('\n') if len(x)>3]
    
    # Identify unique people
    online_verified = set()
    onsite_verified = set()
    ambig = []

    for z in list_online_raw:
        best, conf = get_best_match_info(z, db_names)
        if best: 
            online_verified.add(best)
            if len(conf)>1: ambig.append({"Input":z, "Pilih":best})
            
    for z in list_onsite_raw:
        best, conf = get_best_match_info(z, db_names)
        if best: 
            onsite_verified.add(best)
    
    # Total Hadir (Union)
    all_hadir = online_verified.union(onsite_verified)

    # 3. FEEDBACK LOGIC
    final_fb = []
    if file_feedback:
        df_fb = load_data_smart(file_feedback)
        if df_fb is not None:
            col_fb = next((c for c in df_fb.columns if 'nama' in str(c).lower()), None)
            col_sesi = next((c for c in df_fb.columns if 'pertemuan' in str(c).lower() or 'sesi' in str(c).lower()), None)
            
            targets = [str(s) for s in get_session_list(inp_pertemuan)]
            if col_fb:
                for _, r in df_fb.iterrows():
                    valid = True
                    if col_sesi:
                        # Logic: Irisan angka antara CSV dan Target
                        csv_val = re.findall(r'\d+', str(r[col_sesi]))
                        if not set(targets).intersection(set(csv_val)): valid = False
                    
                    if valid:
                        m, _ = get_best_match_info(str(r[col_fb]), db_names)
                        if m: final_fb.append(m)

    fb_ok_count = 0
    fb_no_list = []
    for h in all_hadir:
        if h in final_fb: fb_ok_count += 1
        else: fb_no_list.append(h)
    
    stats = {
        'total': len(db), 
        'hadir': len(all_hadir), 
        'online_count': len(online_verified),
        'onsite_count': len(onsite_verified),
        'fb_ok': fb_ok_count, 
        'fb_no': len(fb_no_list), 
        'pct': round(fb_ok_count/len(all_hadir)*100,1) if all_hadir else 0
    }

    # DASHBOARD
    st.markdown("---")
    st.write("### 📊 Dashboard Hasil")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Hadir", f"{stats['hadir']}/{stats['total']}")
    m2.metric("Online / Onsite", f"{stats['online_count']} / {stats['onsite_count']}")
    m3.metric("Belum Feedback", stats['fb_no'], delta_color="inverse")
    m4.metric("Estimasi Gaji", f"Rp {inp_fee * (len(get_session_list(inp_pertemuan)) or 1):,.0f}")

    if ambig: st.warning("⚠️ Ambigu"); st.dataframe(ambig)
    if fb_no_list: 
        with st.expander("📢 Copy List Belum Feedback"): st.code("\n".join(fb_no_list))

    st.write("### 📥 Reports & Outputs")
    t1, t2, t3, t4, t5 = st.tabs(["📄 Laporan Utama", "✅ Checklist Sprint", "💬 Template Chat", "🎓 Presensi", "💰 Gaji"])
    
    with t1:
        st.write("Format: `format.xlsx`")
        df = generate_laporan_utama_sesuai_format(info, stats, f"{file_bukti}.jpg", inp_sks, inp_role)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.text_area("Copy:", value=df.to_csv(sep='\t', index=False), height=100, key="c1")
        st.download_button("Download Excel", data=to_excel_download(df), file_name=f"Laporan_Utama_{info['tgl']}.xlsx")

    with t2:
        st.write("Format: `format (1).xlsx`")
        df = generate_checklist_sprint(info)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.text_area("Copy:", value=df.to_csv(sep='\t', index=False), height=100, key="c2")
        st.download_button("Download Excel", data=to_excel_download(df), file_name="Checklist_Sprint.xlsx")

    with t3:
        st.write("Format: `format (2).xlsx`")
        df = generate_chat_templates(info)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.text_area("Copy:", value=df.to_csv(sep='\t', index=False), height=100, key="c3")
        
    with t4:
        df = generate_presensi_real(db, all_hadir, final_fb, info)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.text_area("Copy:", value=df.to_csv(sep='\t', index=False), height=150, key="c4")
        st.download_button("Download Excel", data=to_excel_download(df), file_name="Presensi.xlsx")
        
    with t5:
        df = generate_gaji(info, inp_fee, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.text_area("Copy:", value=df.to_csv(sep='\t', index=False), height=100, key="c5")
        st.download_button("Download Excel", data=to_excel_download(df), file_name="Gaji.xlsx")

st.markdown('</div>', unsafe_allow_html=True)