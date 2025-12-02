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
# 1. UI CONFIGURATION (MODERN GLASS)
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

    .glass-box {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 24px -1px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 20px;
    }

    section[data-testid="stSidebar"] {
        background-color: rgba(10, 10, 15, 0.9);
        border-right: 1px solid rgba(255,255,255,0.05);
    }

    .stTextInput input, .stTextArea textarea, .stNumberInput input, .stSelectbox div[data-testid="stMarkdownContainer"] {
        background-color: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        color: white !important;
        border-radius: 8px !important;
    }

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
    try:
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
    except: return ""

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
# 3. GENERATOR LAPORAN SESUAI FORMAT.XLSX
# ==========================================

def generate_laporan_utama_sesuai_format(info, stats, filename, sks_val, role_val):
    text_summary = (
        f"Kelas : {info['kode']}\n"
        f"Jam : {info['jam_full']}\n"
        f"Pertemuan : {info['pertemuan']}\n"
        f"Tanggal : {info['tgl']}\n"
        f"1. Total Mahasiswa Terdaftar : {stats['total']}\n"
        f"2. Jumlah Mahasiswa Hadir : {stats['hadir_valid']}\n"
        f"3. Jumlah Mengisi Form Feedback: {stats['fb_ok']}\n"
        f"4. Jumlah belum mengisi form feedback: {stats['fb_no']}\n"
        f"5. Jumlah tidak hadir : {stats['total'] - stats['hadir_valid']}"
    )

    data = {
        "Tanggal Kehadiran": [info['tgl']],
        "Nama Dosen": [info['dosen']],
        "Nama Mata Kuliah": [info['matkul']],
        "Jam Perkuliahan": [info['jam_full']],
        "SKS": [sks_val],
        "Tipe Kelas": [info['tipe']],
        "Sesi Kelas": [info['pertemuan']],
        "Peran": [role_val],
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
    tasks = [
        "Reminder informasi kelas kepada mahasiswa", "Reminder informasi kelas kepada dosen",
        "Pastikan ke dosen apakah materi dishare sendiri atau dibantu", "Jika kelas cancel, jangan lupa update di Informasi Fasilitator",
        "Jika kelas direschedule, jangan lupa update di spreadsheet", "Jika fasil berhalangan, silakan infokan ke group",
        "Matikan proyektor, laptop, AC dan peralatan lainnya (Onsite)", "Kembalikan alat-alat pendukung ke bagian akademik (Onsite)",
        "Reminder presensi attendance dan form feedback saat kelas berjalan", "Update informasi terkait kelas (recording Zoom, dll)",
        "Mengisi rekap kehadiran fasilitator", "Update informasi jumlah pengisi form feedback",
        "Mengirimkan file PDF feedback dosen", "Rekap presensi kehadiran kelas",
        "Follow-up mahasiswa yang belum melakukan pengisian feedback"
    ]
    col_name = f"{info['kode']} - Pertemuan {info['pertemuan']}"
    return pd.DataFrame({"Task": tasks, col_name: ["TRUE"] * len(tasks)})

def generate_chat_templates(info):
    templates = [
        {"Context": "Perkenalan", "Text": f"Halo Bapak/Ibu, selamat pagi/siang/sore.\n\nPerkenalkan, saya {info['fasil']} sebagai fasilitator untuk kelas {info['kode']}, salam kenal ya Bapak/Ibu.\n\nSekaligus saya izin mengingatkan Bapak/Ibu untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJika ada yang ingin ditanyakan, silakan yaa Bapak/Ibu. Terima kasih"},
        {"Context": "Reminder Kelas", "Text": f"Halo teman-teman, selamat pagi/siang/sore/malam.\n\nSehubungan dengan jadwal perkuliahan yang berjalan, saya izin mengingatkan untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJangan sampai terlambat dan sampai jumpa di kelas👋"},
        {"Context": "Reminder Dosen", "Text": f"Halo Bapak/Ibu, selamat pagi/siang/sore.\n\nSaya izin mengingatkan untuk kelas yang akan berlangsung pada:\n\n📚 Judul Materi/Mata Kuliah: {info['matkul']}\n🗓️ Hari/tanggal: {info['tgl']}\n⏰ Jadwal: {info['jam_full']}\n\nJIka ada yang bisa dibantu atau yang ingin ditanyakan terkait keberlangsungan kelas nanti, boleh ya Bapak/Ibu. Terima kasih🙏"},
        {"Context": "Kirim Feedback", "Text": f"Halo teman-teman,\n\njangan lupa untuk mengisi link feedback kelas sebagai bukti presensi pada link berikut yaa:\n\n🔗 [LINK FEEDBACK]\n\nTerima kasih dan sampai jumpa di kelas berikutya!"}
    ]
    return pd.DataFrame(templates)

def generate_presensi_real(db, hadir_valid_set, list_feedback, info):
    df = db.copy()
    col_nm = next((c for c in df.columns if 'nama' in str(c).lower()), df.columns[1])
    col_nim = next((c for c in df.columns if 'nim' in str(c).lower()), df.columns[0])
    
    df_out = df[[col_nim, col_nm]].copy()
    df_out.columns = ['NIM', 'Nama Mahasiswa']
    
    target_sessions = get_session_list(info['pertemuan'])
    
    for idx, r in df_out.iterrows():
        n = str(r['Nama Mahasiswa'])
        code = "A"
        if n in hadir_valid_set: 
            code = "O" if n in list_feedback else "OF"
        
        for s in target_sessions:
            if 1 <= int(s) <= 16:
                df_out.loc[idx, f"Sesi {s}"] = code
                
    return df_out

def generate_gaji(info, fee, filename):
    jml = len(get_session_list(info['pertemuan'])) or 1
    return pd.DataFrame([{"Tanggal": info['tgl'], "Dosen": info['dosen'], "Matkul": info['matkul'], "Kode": info['kode'], "Sesi": info['pertemuan'], "Jml": jml, "Fee": fee, "Total": fee*jml, "Bukti": filename}])

# ==========================================
# 4. CORE LOGIC (SINGLE & BATCH)
# ==========================================

def run_analysis(info, txt_zoom, txt_onsite, db_names, df_fb):
    # Parsing Input
    list_zoom = [clean_nama_zoom(x) for x in str(txt_zoom).split('\n') if len(x)>3]
    list_onsite = [clean_nama_zoom(x) for x in str(txt_onsite).split('\n') if len(x)>3]
    
    hadir_zoom = set()
    hadir_onsite = set()
    ambig = []

    # Matching Zoom
    for z in list_zoom:
        best, conf = get_best_match_info(z, db_names)
        if best:
            hadir_zoom.add(best)
            if len(conf)>1: ambig.append({"Input":z, "Pilih":best})
    
    # Matching Onsite
    for z in list_onsite:
        best, conf = get_best_match_info(z, db_names)
        if best: hadir_onsite.add(best)

    # Feedback Logic
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
                        if m not in hadir_zoom and m not in hadir_onsite:
                            ghosts.append(m)

    # Hitungan Hadir (Valid hanya Zoom/Onsite)
    real_attendees = hadir_zoom.union(hadir_onsite)
    
    fb_ok_count = 0
    fb_no_list = []
    for h in real_attendees:
        if h in final_fb: fb_ok_count += 1
        else: fb_no_list.append(h)
        
    stats = {
        'total': len(db_names), 
        'hadir_valid': len(real_attendees), 
        'online_count': len(hadir_zoom),
        'onsite_count': len(hadir_onsite),
        'fb_ok': fb_ok_count, 
        'fb_no': len(fb_no_list), 
        'pct': round(fb_ok_count/len(real_attendees)*100,1) if real_attendees else 0,
        'ghosts': ghosts,
        'ambig': ambig,
        'fb_no_list': fb_no_list
    }
    
    return stats, real_attendees, final_fb

# ==========================================
# 5. USER INTERFACE
# ==========================================

st.markdown('<h1 class="main-header">⚡ Admin Kelas Pro Max</h1>', unsafe_allow_html=True)
st.write("")

# --- SIDEBAR: MODE SELECTION ---
with st.sidebar:
    st.markdown("### 🎛️ Control Panel")
    app_mode = st.radio("Pilih Mode:", ["👤 Single Process (1 Kelas)", "🚀 Batch Process (Banyak)"])
    
    st.markdown("---")
    inp_sks = st.number_input("SKS Default", value=3)
    inp_fee = st.number_input("Fee Default", value=150000)
    inp_role = st.text_input("Peran", value="Fasilitator Kelas")

# ==========================================
# MODE 1: SINGLE PROCESS
# ==========================================
if app_mode == "👤 Single Process (1 Kelas)":
    
    with st.expander("📝 Input Data Kelas (Klik untuk Buka/Tutup)", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            inp_input_method = st.radio("Sumber:", ["Manual", "DB Excel"], horizontal=True)
            defaults = {"jam_mulai":"00.00", "jam_full":"00:00-00:00", "matkul":"", "dosen":"", "kode":"", "tipe":["Reguler"], "fasil": "Fasil"}
            
            if inp_input_method == "Manual":
                raw = st.text_area("Paste Jadwal:", height=70)
                if raw: 
                    parsed = parse_data_template(raw)
                    defaults.update(parsed)
            else:
                up_db = st.file_uploader("Upload DB Jadwal", type=['xlsx', 'xls', 'csv'])
                if up_db:
                    df_db = load_data_smart(up_db)
                    if df_db is not None:
                        col_label = next((c for c in df_db.columns if 'label' in c.lower()), None)
                        if not col_label: # Fallback label
                            col_kode = next((c for c in df_db.columns if 'kode' in c.lower()), df_db.columns[0])
                            df_db['Label_Gen'] = df_db[col_kode].astype(str)
                            col_label = 'Label_Gen'
                        
                        pilihan = st.selectbox("Pilih Kelas:", df_db[col_label])
                        row = df_db[df_db[col_label] == pilihan].iloc[0]
                        # Mapping basic fields (Simplified)
                        col_matkul = next((c for c in df_db.columns if 'mata' in c.lower()), '')
                        if col_matkul: defaults['matkul'] = str(row[col_matkul])
                        # (Assume other fields exist for brevity, logic same as previous)
        
        with c2:
            i_tgl = st.text_input("Tanggal", datetime.now().strftime("%d %B %Y"))
            i_matkul = st.text_input("Matkul", value=defaults['matkul'])
            i_dosen = st.text_input("Dosen", value=defaults['dosen'])
            i_fasil = st.text_input("Fasil", value=defaults['fasil'])
            i_kode = st.text_input("Kode", value=defaults['kode'])
            i_jam = st.text_input("Jam", value=defaults['jam_full'])
            i_tipe = st.text_input("Tipe", value=defaults['tipe'][0])
            i_sesi = st.text_input("Sesi", value="1")

            info = {"tgl":i_tgl, "matkul":i_matkul, "dosen":i_dosen, "kode":i_kode, "jam_full":i_jam, "tipe":i_tipe, "pertemuan":i_sesi, "fasil":i_fasil}

    c_ocr, c_file = st.columns(2)
    with c_ocr:
        up_img = st.file_uploader("📸 Bukti Foto (OCR)", type=['jpg','png'])
        txt_ocr = extract_text_from_image(up_img) if up_img else ""
        txt_zoom = st.text_area("List Zoom:", value=txt_ocr, height=100)
        txt_onsite = st.text_area("List Onsite:", height=80)
    with c_file:
        file_master = st.file_uploader("📂 Master Mhs", type=['xlsx', 'csv'])
        file_fb = st.file_uploader("📂 Feedback", type=['xlsx', 'csv'])

    if st.button("🚀 PROSES SINGLE"):
        if not file_master: st.error("Master Wajib Ada"); st.stop()
        
        db = load_data_smart(file_master)
        col_nm = next((c for c in db.columns if 'nama' in str(c).lower()), None)
        db_names = db[col_nm].astype(str).tolist()
        
        df_fb_data = load_data_smart(file_fb) if file_fb else None
        
        stats, real_attendees, final_fb = run_analysis(info, txt_zoom, txt_onsite, db_names, df_fb_data)
        
        file_bukti = f"{info['tgl']}_{info['dosen']}_{info['matkul']}_{info['fasil']}_Pertemuan {info['pertemuan']}"
        
        st.info(f"Hadir: {stats['hadir_valid']} | Feedback: {stats['fb_ok']} | Gaji: Rp {inp_fee * len(get_session_list(info['pertemuan'])):,.0f}")
        
        if stats['ghosts']:
            st.warning(f"⚠️ {len(stats['ghosts'])} Mahasiswa Isi Feedback Tapi TIDAK HADIR (Hantu):")
            st.write(", ".join(stats['ghosts']))

        t1, t2, t3, t4, t5 = st.tabs(["Laporan", "Sprint", "Chat", "Presensi", "Gaji"])
        with t1:
            df = generate_laporan_utama_sesuai_format(info, stats, f"{file_bukti}.jpg", inp_sks, inp_role)
            st.dataframe(df)
            st.download_button("Download", to_excel_download(df), f"Laporan_{info['kode']}.xlsx")
        with t2:
            df = generate_checklist_sprint(info)
            st.dataframe(df)
            st.download_button("Download", to_excel_download(df), "Checklist.xlsx")
        with t3:
            df = generate_chat_templates(info)
            st.dataframe(df)
        with t4:
            df = generate_presensi_real(db, real_attendees, final_fb, info)
            st.dataframe(df)
            st.download_button("Download", to_excel_download(df), "Presensi.xlsx")
        with t5:
            df = generate_gaji(info, inp_fee, f"{file_bukti}.jpg")
            st.dataframe(df)
            st.download_button("Download", to_excel_download(df), "Gaji.xlsx")

# ==========================================
# MODE 2: BATCH PROCESS
# ==========================================
elif app_mode == "🚀 Batch Process (Banyak)":
    st.info("Upload File Batch Excel: Kolom Wajib -> Tanggal, Fasilitator, Jam, Kode Kelas, Mata Kuliah, Nama Dosen, Tipe, Sesi, Nama File Foto")
    
    up_batch = st.file_uploader("1. Upload 'Jadwal_Batch.xlsx'", type=['xlsx'])
    up_master = st.file_uploader("2. Upload Master Mahasiswa", type=['xlsx'])
    up_fb = st.file_uploader("3. Upload File Feedback", type=['xlsx', 'csv'])
    up_imgs = st.file_uploader("4. Upload SEMUA FOTO BUKTI", type=['jpg','png'], accept_multiple_files=True)
    
    if st.button("🚀 JALANKAN BATCH"):
        if not (up_batch and up_master and up_imgs):
            st.error("File belum lengkap!")
            st.stop()
            
        db = load_data_smart(up_master)
        col_nm = next((c for c in db.columns if 'nama' in str(c).lower()), None)
        db_names = db[col_nm].astype(str).tolist()
        
        df_batch = load_data_smart(up_batch)
        df_fb_data = load_data_smart(up_fb) if up_fb else None
        
        img_map = {f.name: f for f in up_imgs}
        
        res_laporan = []
        res_gaji = []
        
        progress = st.progress(0)
        
        for idx, row in df_batch.iterrows():
            info = {
                "tgl": str(row['Tanggal']), "matkul": str(row['Mata Kuliah']), 
                "dosen": str(row['Nama Dosen']), "kode": str(row['Kode Kelas']),
                "jam_full": str(row['Jam']), "pertemuan": str(row['Sesi']),
                "fasil": str(row['Fasilitator']), "tipe": str(row['Tipe'])
            }
            target_img_name = str(row['Nama File Foto'])
            
            txt_zoom = ""
            if target_img_name in img_map:
                txt_zoom = extract_text_from_image(img_map[target_img_name])
            
            stats, _, _ = run_analysis(info, txt_zoom, "", db_names, df_fb_data)
            
            df_lap = generate_laporan_utama_sesuai_format(info, stats, target_img_name, inp_sks, inp_role)
            df_gaj = generate_gaji(info, inp_fee, target_img_name)
            
            res_laporan.append(df_lap)
            res_gaji.append(df_gaj)
            progress.progress((idx + 1) / len(df_batch))
            
        final_laporan = pd.concat(res_laporan)
        final_gaji = pd.concat(res_gaji)
        
        st.success("✅ Batch Selesai!")
        c1, c2 = st.columns(2)
        c1.download_button("📥 Download Laporan Gabungan", to_excel_download(final_laporan), "Batch_Laporan.xlsx")
        c2.download_button("📥 Download Rekap Gaji", to_excel_download(final_gaji), "Batch_Gaji.xlsx")

st.markdown('</div>', unsafe_allow_html=True)