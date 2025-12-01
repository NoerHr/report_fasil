import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
import difflib

# ==========================================
# 1. CONFIG & CSS (GLASSMORPHISM UI)
# ==========================================
st.set_page_config(page_title="Admin Kelas Pro", layout="wide", page_icon="💎")

def local_css():
    st.markdown("""
    <style>
    /* 1. BACKGROUND GRADIENT UTAMA */
    .stApp {
        background: linear-gradient(120deg, #1e3c72 0%, #2a5298 100%);
        background-attachment: fixed;
        color: white;
    }

    /* 2. SIDEBAR GLASS EFFECT */
    section[data-testid="stSidebar"] {
        background-color: rgba(255, 255, 255, 0.05);
        backdrop-filter: blur(15px);
        border-right: 1px solid rgba(255, 255, 255, 0.1);
    }

    /* 3. GLASS CARDS (KOTAK KACA) */
    .glass-card {
        background: rgba(255, 255, 255, 0.1);
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border-radius: 15px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        padding: 20px;
        margin-bottom: 20px;
        color: white;
    }
    
    .glass-card h3 {
        color: #f0f2f6;
        font-weight: 600;
        margin-bottom: 5px;
    }
    
    .glass-card p {
        color: #e0e0e0;
        font-size: 0.9rem;
    }

    /* 4. CUSTOM BUTTONS */
    div.stButton > button {
        background: linear-gradient(90deg, #00d2ff 0%, #3a7bd5 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
        width: 100%;
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(0,0,0,0.3);
    }
    
    /* 5. INPUT FIELDS STYLING */
    .stTextInput input, .stTextArea textarea, .stSelectbox, .stNumberInput input {
        background-color: rgba(0, 0, 0, 0.3) !important;
        color: white !important;
        border: 1px solid rgba(255, 255, 255, 0.2) !important;
        border-radius: 8px;
    }
    
    /* 6. DATAFRAME & TABLES */
    .stDataFrame {
        background-color: rgba(255, 255, 255, 0.9);
        border-radius: 10px;
        padding: 10px;
        color: black;
    }
    
    /* JUDUL BESAR */
    h1 {
        text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
        font-weight: 800 !important;
    }
    
    </style>
    """, unsafe_allow_html=True)

local_css()

# ==========================================
# 2. FUNGSI LOGIKA (BACKEND)
# ==========================================

def parse_data_template(text):
    data = {}
    jam_match = re.search(r'(\d{1,2}[\.:]\d{2})\s?-\s?(\d{1,2}[\.:]\d{2})', text)
    if jam_match:
        data['jam_full'] = jam_match.group(0).replace('.', ':')
        parts = text.split(jam_match.group(0))
        data['fasilitator'] = parts[0].strip()
        sisa = parts[1] if len(parts) > 1 else text
    else:
        data['jam_full'] = "00:00 - 00:00"
        data['fasilitator'] = "Fasilitator"
        sisa = text

    pertemuan_match = re.search(r'Pertemuan\s?([\d\s&,-]+)', sisa, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "1"

    kode_match = re.search(r'([A-Z]{2,5}\d{1,3})', sisa)
    data['kode_kelas'] = kode_match.group(1) if kode_match else "KODE_KELAS"

    if data['kode_kelas'] in sisa:
        parts_matkul = sisa.split(data['kode_kelas'])
        data['matkul'] = parts_matkul[0].strip()
        data['dosen'] = parts_matkul[1].split(',')[0].strip()
        if "Pertemuan" in data['dosen']: data['dosen'] = "Nama Dosen"
    else:
        data['matkul'] = "Nama Mata Kuliah"
        data['dosen'] = "Nama Dosen"
    return data

def clean_nama_zoom(nama_raw):
    if not isinstance(nama_raw, str): return ""
    if any(x in nama_raw.lower() for x in ["fasil", "host", "admin", "co-host"]): return "IGNORE"
    nama = re.sub(r'^[\d\-\_\.]+\s*', '', nama_raw) 
    nama = re.sub(r'[\-\_]\s*[A-Z]{2,3}$', '', nama) 
    return nama.strip().title()

def get_best_match_info(nama_zoom, list_db_names):
    nama_zoom = nama_zoom.lower()
    db_lower_map = {name.lower(): name for name in list_db_names}
    
    for db_low, db_real in db_lower_map.items():
        if nama_zoom in db_low or db_low in nama_zoom:
            potential_conflicts = [db_lower_map[k] for k in db_lower_map if nama_zoom in k]
            if len(potential_conflicts) > 1:
                return db_real, potential_conflicts
            return db_real, []

    matches = difflib.get_close_matches(nama_zoom, list_db_names, n=3, cutoff=0.6)
    if not matches: return None, []
    
    best_match = matches[0]
    if len(matches) > 1: return best_match, matches
    return best_match, []

def get_session_list(pertemuan_str):
    clean_str = pertemuan_str.replace('&', ',').replace('-', ',').replace('dan', ',')
    return [int(x) for x in clean_str.split(',') if x.strip().isdigit()]

# ==========================================
# 3. GENERATOR LAPORAN
# ==========================================

def generate_laporan_sprint(info, stats):
    tasks = ["Reminder H-1/H-Jam", "Dosen Hadir & Materi Siap", "Host Zoom Claim", "Absensi & Screenshot", 
             "Recording Berjalan", "Upload GDrive", "Laporan Fasil (Sistem)", "Update Feedback"]
    return pd.DataFrame({"No": range(1, len(tasks)+1), "Task": tasks, "Status": ["v"]*len(tasks), "Ket": ["Done"]*len(tasks)})

def generate_laporan_fasilitator_batch3(info, stats, filename_bukti):
    text_summary = (f"Kelas: {info['kode']}\nJam: {info['jam']}\nSesi: {info['pertemuan']}\n"
                    f"Mhs Terdaftar: {stats['total_mhs']}\nHadir: {stats['hadir']}\n"
                    f"Feedback: {stats['sudah_fb']} | Belum: {stats['belum_fb']}")
    data = {"Dosen/Tgl": [f"{info['dosen']}\n{info['tgl']}"], "Matkul": [info['matkul']], "Jam": [info['jam']],
            "Tipe": [info['lokasi'].split()[0]], "Sesi": [info['pertemuan']], "Bukti": [filename_bukti],
            "Validasi": ["Valid"], "Feedback Summary": [text_summary], "Persentase FB": [f"{stats['persen_fb']}%"]}
    return pd.DataFrame(data)

def generate_presensi_mahasiswa(db_master, list_hadir, list_fb, info):
    df = db_master.copy()
    col_nama = [c for c in df.columns if "Nama" in c][0]
    for i in range(1, 17): df[f'Sesi {i}'] = ""
    mode = "O" if "Online" in info['lokasi'] else "S"
    target_sessions = get_session_list(info['pertemuan'])
    
    for idx, row in df.iterrows():
        nama = str(row[col_nama])
        hadir = nama in list_hadir
        fb = nama in list_fb
        code = "A"
        if hadir: code = mode if fb else f"{mode}F"
        for s in target_sessions:
            if 1 <= s <= 16: df.at[idx, f'Sesi {s}'] = code
    return df

def generate_rekap_gaji(info, fee, filename_bukti):
    jml_sesi = len(get_session_list(info['pertemuan'])) or 1
    return pd.DataFrame([{
        "Tanggal": info['tgl'], "Matkul": info['matkul'], "Sesi": info['pertemuan'], 
        "Jml Sesi": jml_sesi, "Fee/Sesi": fee, "Total": fee * jml_sesi, "Bukti": filename_bukti
    }])

# ==========================================
# 4. USER INTERFACE (GLASS UI)
# ==========================================

st.title("💎 Admin Kelas Pro")
st.markdown("---")

# --- SIDEBAR (SETTING KELAS) ---
with st.sidebar:
    st.header("⚙️ Setting Kelas")
    st.info("Paste jadwal acak di bawah ini untuk auto-fill:")
    raw_template = st.text_area("Template Jadwal:", height=80, help="Paste teks jadwal acak disini")
    
    defaults = {"jam": "00:00 - 00:00", "matkul": "", "dosen": "", "kode": "", "pertemuan": "1"}
    if raw_template:
        parsed = parse_data_template(raw_template)
        defaults.update({k: v for k, v in parsed.items() if k in defaults})
        if 'pertemuan_str' in parsed: defaults['pertemuan'] = parsed['pertemuan_str']

    # INPUT FORM DENGAN STYLE
    inp_tgl = st.text_input("📅 Tanggal", datetime.now().strftime("%d %B %Y"))
    inp_matkul = st.text_input("📚 Mata Kuliah", value=defaults['matkul'])
    inp_dosen = st.text_input("👨‍🏫 Nama Dosen", value=defaults['dosen'])
    c1, c2 = st.columns(2)
    inp_jam = c1.text_input("⏰ Jam", value=defaults['jam'])
    inp_kode = c2.text_input("🏷️ Kode Kelas", value=defaults['kode'])
    
    c3, c4 = st.columns(2)
    inp_pertemuan = c3.text_input("🔢 Sesi (ex: 1 & 2)", value=defaults['pertemuan'])
    inp_fee = c4.number_input("💰 Fee", value=150000)
    
    inp_lokasi = st.selectbox("📍 Mode", ["Online (Zoom)", "Onsite (Kampus)"])

    info_kelas = {"tgl": inp_tgl, "jam": inp_jam, "matkul": inp_matkul, "dosen": inp_dosen, "kode": inp_kode, "pertemuan": inp_pertemuan, "lokasi": inp_lokasi}

# --- BAGIAN 1: REQUEST FORMAT FILE ---
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.markdown("### 1️⃣ Request Format Nama File")
st.markdown("Gunakan format ini untuk rename screenshot Zoom Anda:")
nama_file_final = f"{inp_matkul}_Sesi {inp_pertemuan}_{inp_tgl}_{inp_kode}_{inp_dosen}".replace("/", "-")
st.code(f"{nama_file_final}.jpg", language="text")
st.markdown('</div>', unsafe_allow_html=True)

# --- BAGIAN 2: UPLOAD AREA ---
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.markdown("### 2️⃣ Upload Data")
col_up1, col_up2 = st.columns(2)
with col_up1:
    txt_zoom = st.text_area("📋 Paste Nama Peserta (Zoom)", height=150, placeholder="Ryan Ahmadi\nBudi Santoso...")
with col_up2:
    file_master = st.file_uploader("📂 Upload Master Mahasiswa (.xlsx)", type=['xlsx'])
    file_feedback = st.file_uploader("📂 Upload Feedback (.csv) - Opsional", type=['csv'])
st.markdown('</div>', unsafe_allow_html=True)

# --- TOMBOL PROSES ---
if st.button("🚀 PROSES DATA & GENERATE LAPORAN"):
    if not txt_zoom or not file_master:
        st.error("❌ Mohon lengkapi Data Peserta dan Master Mahasiswa!")
        st.stop()

    # --- LOGIC PROCESSING ---
    try:
        db_master = pd.read_excel(file_master)
        col_target_db = [c for c in db_master.columns if "Nama" in c or "Name" in c][0]
        db_names = db_master[col_target_db].astype(str).tolist()
    except:
        st.error("Gagal membaca file Excel. Pastikan ada kolom 'Nama'.")
        st.stop()

    list_hadir_raw = [clean_nama_zoom(x) for x in txt_zoom.split('\n') if clean_nama_zoom(x) not in ["IGNORE", ""]]
    
    final_attendance = []
    ambiguous = []
    
    for z in list_hadir_raw:
        best, conflicts = get_best_match_info(z, db_names)
        if best:
            final_attendance.append(best)
            if len(conflicts) > 1: ambiguous.append({"Input": z, "Match": best, "Konflik": conflicts})

    final_feedback = []
    if file_feedback:
        try:
            df_fb = pd.read_csv(file_feedback)
            col_nama_fb = [c for c in df_fb.columns if "Nama" in c][0]
            col_sesi_fb = next((c for c in df_fb.columns if "Pertemuan" in c or "Sesi" in c), None)
            targets = [str(s) for s in get_session_list(inp_pertemuan)]
            
            for idx, r in df_fb.iterrows():
                valid = True
                if col_sesi_fb and not any(t in str(r[col_sesi_fb]) for t in targets): valid = False
                if valid:
                    match, _ = get_best_match_info(str(r[col_nama_fb]), db_names)
                    if match: final_feedback.append(match)
        except: pass

    hadir_uniq = list(set(final_attendance))
    fb_uniq = list(set(final_feedback))
    
    sudah_fb = 0
    belum_fb_names = []
    for h in hadir_uniq:
        if h in fb_uniq: sudah_fb += 1
        else: belum_fb_names.append(h)

    stats = {'total_mhs': len(db_master), 'hadir': len(hadir_uniq), 
             'sudah_fb': sudah_fb, 'belum_fb': len(belum_fb_names), 
             'persen_fb': round((sudah_fb/len(hadir_uniq)*100),1) if hadir_uniq else 0}

    # --- HASIL DASHBOARD ---
    st.markdown("### 📊 Ringkasan Hasil")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Hadir", f"{len(hadir_uniq)}/{len(db_master)}")
    c2.metric("Sudah Feedback", f"{sudah_fb}")
    c3.metric("Belum Feedback", f"{len(belum_fb_names)}", delta_color="inverse")
    
    # Hitung Gaji Total
    sesi_count = len(get_session_list(inp_pertemuan)) or 1
    total_gaji = inp_fee * sesi_count
    c4.metric("Estimasi Gaji", f"Rp {total_gaji:,.0f}")

    if ambiguous:
        st.warning("⚠️ Ada nama ambigu (Sistem memilih yang paling mirip)")
        st.dataframe(pd.DataFrame(ambiguous), hide_index=True)
    
    if belum_fb_names:
        with st.expander("📢 Klik untuk lihat Mahasiswa Belum Feedback (Copy untuk Tagih)", expanded=True):
            st.code("\n".join(belum_fb_names), language="text")

    # --- TABS HASIL ---
    st.markdown("### 📥 Download Laporan")
    t1, t2, t3, t4 = st.tabs(["✅ Checklist", "📝 Fasil", "🎓 Presensi", "💰 Gaji"])

    with t1:
        df_s = generate_laporan_sprint(info_kelas, stats)
        st.dataframe(df_s, hide_index=True, use_container_width=True)
        st.text_area("Copy Checklist:", value=df_s.to_csv(sep='\t', index=False), height=150)

    with t2:
        df_f = generate_laporan_fasilitator_batch3(info_kelas, stats, f"{nama_file_final}.jpg")
        st.dataframe(df_f, hide_index=True)
        st.text_area("Copy Data Fasil:", value=df_f.to_csv(sep='\t', index=False), height=150)

    with t3:
        df_p = generate_presensi_mahasiswa(db_master, hadir_uniq, fb_uniq, info_kelas)
        def warnai(val):
            if val in ['S', 'O']: return 'background-color: #c6efce'
            if 'F' in str(val): return 'background-color: #ffc7ce'
            return ''
        cols = [c for c in df_p.columns if "Nama" in c or "NIM" in c] + [f'Sesi {i}' for i in range(1, 17)]
        st.dataframe(df_p[cols].style.applymap(warnai), use_container_width=True)
        st.text_area("Copy Presensi:", value=df_p[cols].to_csv(sep='\t', index=False), height=150)

    with t4:
        df_g = generate_rekap_gaji(info_kelas, inp_fee, f"{nama_file_final}.jpg")
        st.dataframe(df_g, hide_index=True)
        st.text_area("Copy Gaji:", value=df_g.to_csv(sep='\t', index=False), height=100)