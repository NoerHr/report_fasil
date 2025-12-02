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
# 1. UI CONFIGURATION (GLASSMORPHISM)
# ==========================================
st.set_page_config(page_title="Admin Kelas Pro Max", layout="wide", page_icon="💎")

def local_css():
    st.markdown("""
    <style>
    .stApp {
        background: radial-gradient(circle at 10% 20%, rgb(30, 60, 114) 0%, rgb(42, 82, 152) 90%);
        color: #ffffff;
    }
    section[data-testid="stSidebar"] {
        background-color: rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(10px);
        border-right: 1px solid rgba(255,255,255,0.1);
    }
    .glass-card {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 16px;
        box-shadow: 0 4px 30px rgba(0, 0, 0, 0.1);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255, 255, 255, 0.15);
        padding: 20px;
        margin-bottom: 20px;
    }
    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        color: white !important;
    }
    div.stButton > button {
        background: linear-gradient(90deg, #00C9FF 0%, #92FE9D 100%);
        color: #004852;
        font-weight: bold;
        border: none;
        border-radius: 8px;
        padding: 12px;
        width: 100%;
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

def parse_data_template(text):
    data = {}
    jam_match = re.search(r'(\d{1,2}[\.:]\d{2})\s?-\s?(\d{1,2}[\.:]\d{2})', text)
    if jam_match:
        data['jam_mulai'] = jam_match.group(1).replace(':', '.') 
        data['jam_full'] = jam_match.group(0).replace('.', ':')
        parts = text.split(jam_match.group(0))
        sisa = parts[1] if len(parts) > 1 else text
    else:
        data['jam_mulai'] = "00.00"
        data['jam_full'] = "00:00 - 00:00"
        sisa = text

    pertemuan_match = re.search(r'Pertemuan\s?([\d\s&,-]+)', sisa, re.IGNORECASE)
    data['pertemuan_str'] = pertemuan_match.group(1).strip() if pertemuan_match else "1"

    kode_match = re.search(r'([A-Za-z]{1,5}\d{1,3})', sisa)
    data['kode_kelas'] = kode_match.group(1) if kode_match else "KODE"

    # Parsing Tipe Kelas
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

# ==========================================
# 3. GENERATOR OUTPUT
# ==========================================
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
# 4. USER INTERFACE (SIDEBAR DUAL MODE)
# ==========================================

st.title("💎 Admin Kelas Pro Max")

# --- SIDEBAR CONFIGURATION ---
with st.sidebar:
    st.header("⚙️ Mode Input Data")
    
    # 1. PILIH MODE (Ini yang sebelumnya hilang)
    mode_input = st.radio("Sumber Data:", ["✍️ Manual / Paste", "📂 Database Excel"])
    
    defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"]}
    
    # LOGIKA MODE 1: PARSING MANUAL
    if mode_input == "✍️ Manual / Paste":
        raw_template = st.text_area("Paste Jadwal Gancet:", height=70, help="Contoh: Hafiz18.30 - ...")
        if raw_template:
            parsed = parse_data_template(raw_template)
            defaults.update({k:v for k,v in parsed.items() if k in defaults})
            
    # LOGIKA MODE 2: DATABASE EXCEL
    elif mode_input == "📂 Database Excel":
        uploaded_db_jadwal = st.file_uploader("Upload Database_Jadwal.xlsx", type=['xlsx'])
        if uploaded_db_jadwal:
            try:
                df_jadwal = pd.read_excel(uploaded_db_jadwal)
                # Buat Label Dropdown
                df_jadwal['Label'] = df_jadwal.apply(lambda x: f"{x.get('Kode Kelas','?')} - {x.get('Mata Kuliah','?')} - {x.get('Nama Dosen','?')}", axis=1)
                
                pilihan_kelas = st.selectbox("Pilih Kelas:", df_jadwal['Label'])
                
                # Ambil data
                row = df_jadwal[df_jadwal['Label'] == pilihan_kelas].iloc[0]
                
                # Update Defaults
                defaults['matkul'] = str(row.get('Mata Kuliah', ''))
                defaults['dosen'] = str(row.get('Nama Dosen', ''))
                defaults['kode'] = str(row.get('Kode Kelas', ''))
                jam_raw = str(row.get('Jam Mulai', '00.00'))
                defaults['jam_mulai'] = jam_raw.replace(':', '.')
                defaults['jam_full'] = f"{jam_raw.replace('.', ':')} - Selesai"
                tipe_raw = str(row.get('Tipe', 'Reguler'))
                defaults['tipe'] = [t.strip() for t in tipe_raw.split(',')]
                
            except Exception as e:
                st.error(f"Gagal baca database: {e}")

    st.markdown("---")
    st.subheader("📝 Detail Kelas (Edit)")
    
    inp_tgl = st.text_input("📅 Tanggal", datetime.now().strftime("%d %B %Y"))
    inp_matkul = st.text_input("📚 Mata Kuliah", value=defaults['matkul'])
    inp_dosen = st.text_input("👨‍🏫 Nama Dosen", value=defaults['dosen'])
    
    c1, c2 = st.columns(2)
    inp_jam_dot = c1.text_input("⏰ Jam (08.00)", value=defaults['jam_mulai'])
    inp_kode = c2.text_input("🏷️ Kode (F4)", value=defaults['kode'])
    
    inp_tipe = st.multiselect("🎓 Tipe", ["Reguler", "Profesional", "Akselerasi"], default=[x for x in defaults['tipe'] if x in ["Reguler", "Profesional", "Akselerasi"]])
    inp_tipe_str = " & ".join(inp_tipe) if inp_tipe else "Reguler"

    c3, c4 = st.columns(2)
    inp_pertemuan = c3.text_input("🔢 Sesi (ex: 1 & 2)", value="1")
    inp_fee = c4.number_input("💰 Fee", value=150000)

    info = {
        "tgl": inp_tgl, "matkul": inp_matkul, "dosen": inp_dosen, "jam_dot": inp_jam_dot, 
        "jam_full": defaults['jam_full'], "kode": inp_kode, 
        "pertemuan": inp_pertemuan, "tipe": inp_tipe_str, "lokasi": "Online"
    }

# --- MAIN CONTENT ---

# 1. GENERATOR NAMA FILE
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.subheader("1️⃣ Generator Nama File")
pert_fmt = f"Pertemuan {inp_pertemuan}".replace("&", "dan")

file_req_zoom = f"{inp_matkul}_{pert_fmt}_{inp_tgl}_{inp_tipe_str}_{inp_dosen}_{inp_jam_dot}"
file_bukti = f"{inp_tgl}_{inp_dosen}_{inp_matkul}_{pert_fmt}"

col_h1, col_h2 = st.columns(2)
with col_h1:
    st.markdown("**🅰️ Req Zoom (Tanggal belakangan)**")
    st.code(file_req_zoom, language="text")
with col_h2:
    st.markdown("**🅱️ Bukti Foto (Tanggal duluan)**")
    st.code(f"{file_bukti}.jpg", language="text")
st.markdown('</div>', unsafe_allow_html=True)

# 2. UPLOAD DATA
st.markdown('<div class="glass-card">', unsafe_allow_html=True)
st.subheader("2️⃣ Upload Data Presensi")

col_ocr, col_file = st.columns(2)
with col_ocr:
    st.info("📸 **Screenshot Zoom**")
    up_foto = st.file_uploader("Upload Foto", type=['jpg','png'])
    ocr_res = ""
    if up_foto:
        with st.spinner("AI Sedang Membaca..."):
            try: ocr_res = extract_text_from_image(up_foto)
            except: pass
    txt_zoom = st.text_area("List Zoom (Scan):", value=ocr_res, height=100)
    
    st.markdown("---")
    st.info("🙋‍♂️ **Peserta Onsite / Tambahan**")
    txt_onsite = st.text_area("List Onsite (Manual):", height=80, placeholder="Budi\nSiti")

with col_file:
    st.info("📂 **Data Mahasiswa & Feedback**")
    file_master = st.file_uploader("Master Excel (.xlsx)", type=['xlsx'], help="Wajib ada kolom Nama & NIM")
    file_feedback = st.file_uploader("Feedback CSV (.csv)", type=['csv'])

# --- PROCESS ---
if st.button("🚀 PROSES & GENERATE EXCEL", type="primary"):
    if (not txt_zoom and not txt_onsite) or not file_master:
        st.error("❌ Data Peserta dan Master Mahasiswa Wajib Diisi!")
        st.stop()
    
    try:
        db = pd.read_excel(file_master)
        col_nm = [c for c in db.columns if "Nama" in c][0]
        db_names = db[col_nm].astype(str).tolist()
    except:
        st.error("Gagal baca file Master.")
        st.stop()
        
    list_zoom = [x for x in txt_zoom.split('\n') if len(x)>3]
    list_onsite = [x for x in txt_onsite.split('\n') if len(x)>3]
    raw_names = list_zoom + list_onsite
    
    final_hadir, ambig = [], []
    for z in raw_names:
        best, conf = get_best_match_info(clean_nama_zoom(z), db_names)
        if best:
            final_hadir.append(best)
            if len(conf)>1: ambig.append({"Input":z, "Pilih":best})

    final_fb = []
    if file_feedback:
        try:
            df_fb = pd.read_csv(file_feedback)
            col_fb = [c for c in df_fb.columns if "Nama" in c][0]
            col_sesi = next((c for c in df_fb.columns if "Pertemuan" in c or "Sesi" in c), None)
            
            # --- FIX BUG FILTER SESI (LOGIKA BARU) ---
            targets = [str(s) for s in get_session_list(inp_pertemuan)]
            for _, r in df_fb.iterrows():
                valid = True
                if col_sesi:
                    # Ambil semua angka dari kolom pertemuan di CSV
                    sesi_csv_list = re.findall(r'\d+', str(r[col_sesi]))
                    # Cek apakah ada irisan antara target dan sesi csv
                    if not any(t in sesi_csv_list for t in targets):
                        valid = False
                
                if valid:
                    m, _ = get_best_match_info(str(r[col_fb]), db_names)
                    if m: final_fb.append(m)
        except: pass

    uh, uf = list(set(final_hadir)), list(set(final_fb))
    fb_ok, fb_no = 0, []
    for h in uh:
        if h in uf: fb_ok += 1
        else: fb_no.append(h)
    
    stats = {'total':len(db), 'hadir':len(uh), 'fb_ok':fb_ok, 'fb_no':len(fb_no), 'pct':round(fb_ok/len(uh)*100,1) if uh else 0}

    # DASHBOARD
    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Hadir", f"{len(uh)}/{len(db)}")
    m2.metric("Feedback", f"{fb_ok}")
    m3.metric("Tanpa Feedback", len(fb_no), delta_color="inverse")
    m4.metric("Gaji", f"Rp {inp_fee * (len(get_session_list(inp_pertemuan)) or 1):,.0f}")

    if ambig: st.warning("⚠️ Nama Ambigu"); st.dataframe(ambig)
    if fb_no: 
        with st.expander("📢 Lihat Belum Feedback"): st.code("\n".join(fb_no))

    # TABS
    t1, t2, t3, t4 = st.tabs(["✅ Sprint", "📝 Fasil", "🎓 Presensi", "💰 Gaji"])
    
    with t1:
        df = generate_laporan_sprint(info, stats)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Excel", data=to_excel_download(df), file_name="Sprint.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with t2:
        df = generate_laporan_fasilitator(info, stats, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Excel", data=to_excel_download(df), file_name=f"Laporan_{info['tgl']}.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with t3:
        df = generate_presensi(db, uh, uf, info)
        def color(v): return 'background-color:#c6efce' if v=='O' else 'background-color:#ffc7ce' if 'F' in str(v) else ''
        cols = [c for c in df.columns if "Nama" in c] + [f'Sesi {i}' for i in get_session_list(inp_pertemuan)]
        st.dataframe(df[cols].style.applymap(color), use_container_width=True)
        st.download_button("📥 Excel", data=to_excel_download(df), file_name="Presensi.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    with t4:
        df = generate_gaji(info, inp_fee, f"{file_bukti}.jpg")
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button("📥 Excel", data=to_excel_download(df), file_name="Gaji.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

st.markdown('</div>', unsafe_allow_html=True)
