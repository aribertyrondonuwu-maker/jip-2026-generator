"""
app.py — Antarmuka Streamlit untuk JIP 2026 Article Auto-Generator
"""
import os
import re
import streamlit as st
from platax_builder import Naskah, Penulis, bangun

# ── Nama file template sesuai OJS (tidak boleh diubah) ──────────────────────
TEMPLATE_ID_FINAL  = "Template_Artikel_PLATAX_2026_ID.docx"
TEMPLATE_EN_FINAL  = "Template_Article_PLATAX_2026_EN.docx"
TEMPLATE_ID_BLIND  = "Template_Naskah_BlindReview_JIP_2026.docx"
TEMPLATE_EN_BLIND  = "Template_Manuscript_BlindReview_JIP_2026.docx"

st.set_page_config(page_title="JIP 2026 Article Auto-Generator", page_icon="📄", layout="wide")

def as_records(tabel):
    return tabel.to_dict("records") if hasattr(tabel, "to_dict") else list(tabel)

def susun_penulis(tabel_penulis, daftar_afiliasi, peta=None):
    penulis, catatan = [], []
    for baris in as_records(tabel_penulis):
        nama = str(baris.get("Nama", " ") or " ").strip()
        if not nama:
            continue
        mentah = [int(x) for x in re.findall(r"\d+", str(baris.get("Afiliasi") or " "))]
        if not mentah:
            catatan.append(f"Penulis \"{nama}\" belum diberi nomor afiliasi.")
        nomor = [peta.get(x, x) if peta else x for x in mentah]
        penulis.append(Penulis(
            nama=nama,
            afiliasi_ids=[x for x in nomor if 1 <= x <= len(daftar_afiliasi)],
            is_corresp=bool(baris.get("Korespondensi", False)),
            email=str(baris.get("Email", " ") or " ").strip(),
            orcid=str(baris.get("ORCID", " ") or " ").strip(),
        ))
    if not penulis:
        catatan.append("Belum ada penulis yang diisi.")
    n_koresp = sum(1 for p in penulis if p.is_corresp)
    if n_koresp == 0:
        catatan.append("Belum ada penulis korespondensi yang ditandai.")
    elif n_koresp > 1:
        catatan.append("Lebih dari satu penulis korespondensi ditandai.")
    for i, teks in enumerate(daftar_afiliasi, 1):
        if not teks.strip():
            catatan.append(f"Kotak Afiliasi {i} masih kosong.")
    return penulis, catatan

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
st.sidebar.header("⚙️ Pengaturan")
bahasa_label = st.sidebar.selectbox("🌐 Bahasa naskah:", ["Bahasa Indonesia", "English"], index=0)
BAHASA = "en" if bahasa_label == "English" else "id"
EN = BAHASA == "en"

def L(teks_id: str, teks_en: str) -> str:
    return teks_en if EN else teks_id

versi = st.sidebar.radio("Versi dokumen:", ["Blind Review", "Final (Lengkap)"], index=0)
is_blind = versi == "Blind Review"
st.sidebar.info("📌 Blind Review: Identitas disembunyikan otomatis.\nFinal: Lengkap.")
st.sidebar.markdown("---")

# Pilih template aktif berdasarkan bahasa dan mode
if is_blind:
    template_aktif = TEMPLATE_EN_BLIND if EN else TEMPLATE_ID_BLIND
else:
    template_aktif = TEMPLATE_EN_FINAL if EN else TEMPLATE_ID_FINAL

# Pilih nama file output yang diunduh
if is_blind:
    nama_file_output = TEMPLATE_EN_BLIND if EN else TEMPLATE_ID_BLIND
else:
    nama_file_output = TEMPLATE_EN_FINAL if EN else TEMPLATE_ID_FINAL

template_path = None
if os.path.exists(template_aktif):
    template_path = template_aktif
    st.sidebar.success(f"✅ {template_aktif}")
else:
    st.sidebar.warning(f"⚠️ Template tidak ditemukan: {template_aktif}")

st.title("📄 JIP 2026 Article Auto-Generator")

# ── NAMA HEADING BAB (bilingual) ─────────────────────────────────────────────
H = {
    "bab1":       L("1. Pendahuluan",                    "1. Introduction"),
    "bab2":       L("2. Bahan dan Metode",               "2. Materials and Methods"),
    "met21":      L("2.1. Waktu dan lokasi penelitian",  "2.1. Study period and location"),
    "met22":      L("2.2. Pengumpulan data",             "2.2. Data collection"),
    "met221":     L("2.2.1. Analisis laboratorium",      "2.2.1. Laboratory analysis"),
    "met23":      L("2.3. Analisis data",                "2.3. Data analysis"),
    "bab3":       L("3. Hasil",                          "3. Results"),
    "bab4":       L("4. Pembahasan",                     "4. Discussion"),
    "bab5":       L("5. Simpulan",                       "5. Conclusion"),
}

# ── FORMULIR ─────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📌 Identitas & Judul", "📝 Abstract & Abstrak",
    "📄 Bab 1–5", "🗂 Tabel & Gambar", "🤝 Pernyataan & Pustaka"
])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        judul_id = st.text_input(L("Judul Bahasa Indonesia", "Judul Bahasa Indonesia"), "Keanekaragaman Ikan Gobi di Muara Sungai Tondano")
    with c2:
        judul_en = st.text_input(L("Title in English", "Title in English"), "Diversity of Goby Fish in Tondano River Estuary")
    running_title = st.text_input(
        L("Running title (maks. 60 karakter)", "Running title (max. 60 characters)"),
        "Keanekaragaman Ikan Gobi Muara Tondano", max_chars=60
    )
    st.caption(f"{len(running_title)}/60 {L('karakter', 'characters')}")
    st.markdown("---")
    st.subheader(L("Informasi Penulis", "Author Information"))
    tabel_penulis = st.data_editor(
        [{"Nama": "Febry S. I. Menajang", "Afiliasi": "1", "Korespondensi": True,
          "ORCID": "0000-0000-0000-0000", "Email": "febry@unsrat.ac.id"}],
        num_rows="dynamic", width="stretch", key="tabel_penulis"
    )
    nomor_dirujuk = sorted({
        int(x) for baris in as_records(tabel_penulis)
        if str(baris.get("Nama", " ")).strip()
        for x in re.findall(r"\d+", str(baris.get("Afiliasi") or " "))
    })
    CONTOH_AFILIASI = ["Program Studi Ilmu Kelautan, FPIK, UNSRAT, Manado, 95115, Indonesia"]
    if not nomor_dirujuk:
        daftar_afiliasi = []; peta = None
    else:
        isian = {}
        for nomor in nomor_dirujuk:
            contoh = CONTOH_AFILIASI[nomor - 1] if nomor <= len(CONTOH_AFILIASI) else " "
            isian[nomor] = st.text_area(
                L(f"Afiliasi {nomor}", f"Affiliation {nomor}"),
                value=contoh, height=68, key=f"afiliasi_{nomor}"
            )
        peta = {lama: baru for baru, lama in enumerate(nomor_dirujuk, 1)}
        daftar_afiliasi = [isian[lama].strip() for lama in nomor_dirujuk]
    penulis_list, catatan_penulis = susun_penulis(tabel_penulis, daftar_afiliasi, peta=peta if nomor_dirujuk else None)
    for c in catatan_penulis:
        st.warning(c)
    c3, c4 = st.columns(2)
    with c3:
        telepon = st.text_input(L("Telepon korespondensi", "Corresponding author phone"), "+62-8XXXXXXXXXX")
    with c4:
        email = st.text_input(L("Surel korespondensi", "Corresponding author e-mail"), "nama@unsrat.ac.id")

with tab2:
    abstract_en = st.text_area("Abstract (English)", height=170, value="This study analysed...")
    keywords_en = st.text_input("Keywords (comma-separated)", "Goby fish, Diversity, Estuary")
    abstrak_id = st.text_area("Abstrak (Bahasa Indonesia)", height=170, value="Penelitian ini menganalisis...")
    kata_kunci_id = st.text_input("Kata kunci (dipisah koma)", "Ikan gobi, Keanekaragaman, Estuari")

with tab3:
    bab1     = st.text_area(H["bab1"],   height=140, value=L(
        "Muara Sungai Tondano memiliki peranan ekologis penting (Carpenter & Niem, 1998).",
        "Tondano River Estuary has important ecological roles (Carpenter & Niem, 1998)."))
    metode_21  = st.text_area(H["met21"], height=80, value=L(
        "Penelitian dilaksanakan Januari–Maret 2026...",
        "The study was conducted from January to March 2026..."))
    metode_22  = st.text_area(H["met22"], height=80, value=L(
        "Pengambilan sampel menggunakan jaring insang...",
        "Sampling was carried out using gill nets..."))
    metode_221 = st.text_area(H["met221"], height=80, value=L(
        "Identifikasi morfometrik dilakukan...",
        "Morphometric identification was conducted..."))
    metode_23  = st.text_area(H["met23"], height=80, value=L(
        "Indeks Shannon-Wiener dihitung...",
        "The Shannon-Wiener index was calculated..."))
    bab3     = st.text_area(H["bab3"],   height=120, value=L(
        "Tercatat 12 jenis dari tiga stasiun (Tabel 1).",
        "Twelve species were recorded across three stations (Table 1)."))
    bab4     = st.text_area(H["bab4"],   height=140, value=L(
        "Tingginya keanekaragaman diduga berkaitan...",
        "The high diversity is presumably related to..."))
    bab5     = st.text_area(H["bab5"],   height=90, value=L(
        "Keanekaragaman ikan gobi tergolong sedang...",
        "Goby fish diversity was classified as moderate..."))
    st.markdown("---")
    st.subheader(L("📐 Persamaan Matematika", "📐 Mathematical Equations"))
    st.caption(L("Gunakan sintaks LaTeX. Contoh: `C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}`",
                 "Use LaTeX syntax. Example: `C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}`"))
    n_eq = st.number_input(L("Jumlah persamaan", "Number of equations"), min_value=0, max_value=20, value=0, key="n_eq")
    eq_inputs = {}
    if int(n_eq) > 0:
        for i in range(int(n_eq)):
            eq_num = i + 1
            latex = st.text_area(
                L(f"Persamaan {eq_num} (LaTeX)", f"Equation {eq_num} (LaTeX)"),
                placeholder="Contoh / Example: C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}",
                height=80, key=f"eq_{eq_num}"
            )
            if latex.strip():
                eq_inputs[f'eq_{eq_num}'] = latex

with tab4:
    # ── TABEL ────────────────────────────────────────────────────────────────
    n_tabel = st.number_input(L("Jumlah Tabel", "Number of Tables"),
                               min_value=0, max_value=20, value=1, key="n_tabel")
    daftar_tabel = []
    for i in range(int(n_tabel)):
        with st.expander(f"{L('Tabel', 'Table')} {i+1}", expanded=(i == 0)):
            # Hanya SATU caption sesuai bahasa naskah (tidak bilingual)
            cap_label = L(f"Keterangan Tabel {i+1}", f"Table {i+1} caption")
            cap = st.text_input(cap_label,
                                value=L("Parameter kualitas air.", "Water quality parameters."),
                                key=f"cap_{i}")
            t_data = st.text_area(
                L(f"Isi Tabel {i+1} (pisah kolom dengan ;)", f"Table {i+1} data (separate columns with ;)"),
                value=L("Parameter;Stasiun 1\nSuhu;28,4", "Parameter;Station 1\nTemperature;28.4"),
                key=f"t_data_{i}"
            )
            cat = st.text_input(
                L(f"Catatan kaki Tabel {i+1}", f"Table {i+1} footnote"),
                value=L("Nilai rerata ± SD.", "Mean values ± SD."),
                key=f"cat_{i}"
            )
            daftar_tabel.append({
                "nomor": i + 1,
                "cap": cap,       # satu caption, sesuai bahasa
                "data": t_data,
                "catatan": cat
            })

    # ── GAMBAR ───────────────────────────────────────────────────────────────
    n_gambar = st.number_input(L("Jumlah Gambar", "Number of Figures"),
                                min_value=0, max_value=20, value=1, key="n_gambar")
    daftar_gambar = []
    for j in range(int(n_gambar)):
        with st.expander(f"{L('Gambar', 'Figure')} {j+1}", expanded=(j == 0)):
            berkas = st.file_uploader(
                L(f"Unggah Gambar {j+1}", f"Upload Figure {j+1}"),
                type=["png", "jpg", "jpeg"], key=f"berkas_{j}"
            )
            # Hanya SATU caption sesuai bahasa naskah (tidak bilingual)
            cap_g_label = L(f"Keterangan Gambar {j+1}", f"Figure {j+1} caption")
            cap_g = st.text_input(cap_g_label, key=f"cap_g_{j}")
            lebar = st.slider(
                L(f"Lebar Gambar {j+1} (cm)", f"Figure {j+1} width (cm)"),
                4.0, 7.6, 7.6, 0.2, key=f"lebar_{j}"
            )
            daftar_gambar.append({
                "nomor": j + 1,
                "blob": berkas.getvalue() if berkas else b"",
                "cap": cap_g,     # satu caption, sesuai bahasa
                "lebar": lebar
            })

with tab5:
    konflik = st.text_area(
        L("Konflik kepentingan", "Competing interests"),
        value=L("Penulis menyatakan tidak ada konflik kepentingan.",
                "The authors declare no competing interests relevant to this article.")
    )
    dana = st.text_area(
        L("Sumber dana", "Funding sources"),
        value=L("Tidak berlaku / Not applicable.", "Not applicable.")
    )
    ucapan = st.text_area(
        L("Ucapan terima kasih", "Acknowledgements"),
        value=L("Penulis berterima kasih kepada Laboratorium Biologi Laut.",
                "The authors thank the Marine Biology Laboratory.")
    )
    kontribusi = st.text_area(
        L("Kontribusi penulis (CRediT)", "Author contributions (CRediT)"),
        value="FSIM: conceptualization; ABR: formal analysis."
    )
    data_avail = st.text_area(
        L("Ketersediaan data", "Availability of data and materials"),
        value=L("Dataset tersedia dari penulis korespondensi.",
                "The dataset is available from the corresponding author upon reasonable request.")
    )
    etik = st.text_area(
        L("Persetujuan etik", "Ethics approval and consent to participate"),
        value=L("Tidak berlaku / Not applicable.", "Not applicable.")
    )
    orcid_list = st.text_area(
        "ORCID",
        value="Febry S. I. Menajang — https://orcid.org/0000-0000-0000-0000"
    )
    daftar_pustaka = st.text_area(
        L("Daftar Pustaka (satu entri per baris)", "References (one entry per line)"),
        height=200,
        value="Rondonuwu, A. B. (2025). Diversity. FAS, 28(10), 667–676."
    )

# ── UNDUH ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader(f"📥 {L('Unduh Naskah', 'Download Manuscript')} ({versi})")

tbl_1 = daftar_tabel[0] if daftar_tabel else {"cap": " ", "data": " ", "catatan": " "}
gbr_1 = daftar_gambar[0] if daftar_gambar else {"cap": " ", "blob": b"", "lebar": 7.6}

naskah = Naskah(
    judul_id=judul_id, judul_en=judul_en, running_title=running_title,
    penulis_list=penulis_list, afiliasi_list=daftar_afiliasi,
    telepon=telepon, email_korespondensi=email,
    abstrak_id=abstrak_id, abstrak_en=abstract_en,
    kata_kunci_id=kata_kunci_id, kata_kunci_en=keywords_en,
    bab1=bab1, metode_21=metode_21, metode_22=metode_22,
    metode_221=metode_221, metode_23=metode_23,
    bab3=bab3, bab4=bab4, bab5=bab5,
    tabel_list=daftar_tabel, gambar_list=daftar_gambar,
    tbl_1=tbl_1, gbr_1=gbr_1,
    konflik=konflik, dana=dana, ucapan=ucapan, kontribusi=kontribusi,
    data_avail=data_avail, etik=etik, orcid_list=orcid_list,
    daftar_pustaka=daftar_pustaka,
    bahasa=BAHASA, blind=is_blind, **eq_inputs
)

if not template_path:
    st.error(f"❌ Template tidak ditemukan: {template_aktif}. Pastikan file ada di direktori yang sama.")
    st.stop()

try:
    berkas = bangun(naskah, template_path)
    st.download_button(
        label=f"📄 {L('Unduh Dokumen', 'Download Document')} — {nama_file_output}",
        data=berkas,
        file_name=nama_file_output,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
except Exception as e:
    st.error(f"Gagal menyusun dokumen: {e}")
