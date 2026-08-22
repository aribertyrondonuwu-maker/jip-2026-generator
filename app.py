"""
app.py — Antarmuka Streamlit untuk JIP 2026 Article Auto-Generator
"""
import os
import re
import streamlit as st
from platax_builder import Naskah, Penulis, bangun

TEMPLATE_DEFAULT = "Template_Artikel_PLATAX_2026_FINAL_OJS.docx"

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
    n_koresp = sum(1 for p in penulis if p.is_corresp)  # FIXED: was p.koresp
    if n_koresp == 0:
        catatan.append("Belum ada penulis korespondensi yang ditandai.")
    elif n_koresp > 1:
        catatan.append("Lebih dari satu penulis korespondensi ditandai.")
    for i, teks in enumerate(daftar_afiliasi, 1):
        if not teks.strip():
            catatan.append(f"Kotak Afiliasi {i} masih kosong.")
    return penulis, catatan

# SIDEBAR
st.sidebar.header("⚙️ Pengaturan")
bahasa_label = st.sidebar.selectbox("🌐 Bahasa naskah:", ["Bahasa Indonesia", "English"], index=0)
BAHASA = "en" if bahasa_label == "English" else "id"
EN = BAHASA == "en"

def L(teks_id: str, teks_en: str) -> str:
    return teks_en if EN else teks_id

versi = st.sidebar.radio("Versi dokumen:", ["Blind Review", "Final (Lengkap)"], index=0)
is_blind = versi == "Blind Review"
st.sidebar.info("📌 Blind Review: Identitas disembunyikan otomatis.\n Final: Lengkap.")
st.sidebar.markdown("---")

template_path = None
if os.path.exists(TEMPLATE_DEFAULT):
    template_path = TEMPLATE_DEFAULT
    st.sidebar.success(f"✅ {TEMPLATE_DEFAULT}")
else:
    st.sidebar.warning("Template resmi tidak ditemukan.")

st.title("📄 JIP 2026 Article Auto-Generator")

# FORMULIR
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📌 Identitas & Judul", "📝 Abstract & Abstrak", "📄 Bab 1–5", " Tabel & Gambar", "🤝 Pernyataan & Pustaka"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        judul_id = st.text_input(L("Judul Bahasa Indonesia", "Indonesian title"), "Keanekaragaman Ikan Gobi di Muara Sungai Tondano")
    with c2:
        judul_en = st.text_input(L("Title in English", "English title"), "Diversity of Goby Fish in Tondano River Estuary")
    running_title = st.text_input("Running title (maks. 60 karakter)", "Keanekaragaman Ikan Gobi Muara Tondano", max_chars=60)
    st.caption(f"{len(running_title)}/60 karakter")
    st.markdown("---")
    st.subheader("Informasi Penulis")
    tabel_penulis = st.data_editor(
        [{"Nama": "Febry S. I. Menajang", "Afiliasi": "1", "Korespondensi": True, "ORCID": "0000-0000-0000-0000", "Email": "febry@unsrat.ac.id"}],
        num_rows="dynamic", width="stretch", key="tabel_penulis"
    )
    nomor_dirujuk = sorted({int(x) for baris in as_records(tabel_penulis) if str(baris.get("Nama", " ")).strip() for x in re.findall(r"\d+", str(baris.get("Afiliasi") or " "))})
    CONTOH_AFILIASI = ["Program Studi Ilmu Kelautan, FPIK, UNSRAT, Manado, 95115, Indonesia"]
    if not nomor_dirujuk:
        daftar_afiliasi = []; peta = None
    else:
        isian = {}
        for nomor in nomor_dirujuk:
            contoh = CONTOH_AFILIASI[nomor - 1] if nomor <= len(CONTOH_AFILIASI) else " "
            isian[nomor] = st.text_area(f"Afiliasi {nomor}", value=contoh, height=68, key=f"afiliasi_{nomor}")
        peta = {lama: baru for baru, lama in enumerate(nomor_dirujuk, 1)}
        daftar_afiliasi = [isian[lama].strip() for lama in nomor_dirujuk]
    penulis_list, catatan_penulis = susun_penulis(tabel_penulis, daftar_afiliasi, peta=peta if nomor_dirujuk else None)
    for c in catatan_penulis:
        st.warning(c)
    c3, c4 = st.columns(2)
    with c3:
        telepon = st.text_input("Telepon korespondensi", "+62-8XXXXXXXXXX")
    with c4:
        email = st.text_input("Surel korespondensi", "nama@unsrat.ac.id")

with tab2:
    abstract_en = st.text_area("Abstract (English)", height=170, value="This study analysed...")
    keywords_en = st.text_input("Keywords (comma-separated)", "Goby fish, Diversity, Estuary")
    abstrak_id = st.text_area("Abstrak (Bahasa Indonesia)", height=170, value="Penelitian ini menganalisis...")
    kata_kunci_id = st.text_input("Kata kunci (dipisah koma)", "Ikan gobi, Keanekaragaman, Estuari")

with tab3:
    bab1 = st.text_area("1. Pendahuluan", height=140, value="Muara Sungai Tondano memiliki peranan ekologis penting (Carpenter & Niem, 1998).")
    metode_21 = st.text_area("2.1. Waktu dan lokasi penelitian", height=80, value="Penelitian dilaksanakan Januari–Maret 2026...")
    metode_22 = st.text_area("2.2. Pengumpulan data", height=80, value="Pengambilan sampel menggunakan jaring insang...")
    metode_221 = st.text_area("2.2.1. Analisis laboratorium", height=80, value="Identifikasi morfometrik dilakukan...")
    metode_23 = st.text_area("2.3. Analisis data", height=80, value="Indeks Shannon-Wiener dihitung...")
    bab3 = st.text_area("3. Hasil", height=120, value="Tercatat 12 jenis dari tiga stasiun (Tabel 1).")
    bab4 = st.text_area("4. Pembahasan", height=140, value="Tingginya keanekaragaman diduga berkaitan...")
    bab5 = st.text_area("5. Simpulan", height=90, value="Keanekaragaman ikan gobi tergolong sedang...")
    st.markdown("---")
    st.subheader("📐 Persamaan Matematika")
    st.caption("Gunakan sintaks LaTeX. Contoh: `C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}`")
    n_eq = st.number_input("Jumlah persamaan", min_value=0, max_value=20, value=0, key="n_eq")
    eq_inputs = {}
    if int(n_eq) > 0:
        for i in range(int(n_eq)):
            eq_num = i + 1
            latex = st.text_area(f"Persamaan {eq_num} (LaTeX)", placeholder="Contoh: C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}", height=80, key=f"eq_{eq_num}")
            if latex.strip():
                eq_inputs[f'eq_{eq_num}'] = latex

with tab4:
    n_tabel = st.number_input("Jumlah Tabel", min_value=0, max_value=20, value=1, key="n_tabel")
    daftar_tabel = []
    for i in range(int(n_tabel)):
        with st.expander(f"Tabel {i+1}", expanded=(i==0)):
            cap_id = st.text_input(f"Keterangan Tabel {i+1} (ID)", value="Parameter kualitas air.", key=f"cap_id_{i}")
            cap_en = st.text_input(f"Table {i+1} caption (EN)", value="Water quality parameters.", key=f"cap_en_{i}")
            t_data = st.text_area(f"Isi Tabel {i+1} (pisah dengan titik koma ;)", value="Parameter;Stasiun 1\nSuhu;28,4", key=f"t_data_{i}")
            cat = st.text_input(f"Catatan kaki Tabel {i+1}", value="Nilai rerata ± SD.", key=f"cat_{i}")
            daftar_tabel.append({"nomor": i+1, "cap_id": cap_id, "cap_en": cap_en, "data": t_data, "catatan": cat})
    n_gambar = st.number_input("Jumlah Gambar", min_value=0, max_value=20, value=1, key="n_gambar")
    daftar_gambar = []
    for j in range(int(n_gambar)):
        with st.expander(f"Gambar {j+1}", expanded=(j==0)):
            berkas = st.file_uploader(f"Unggah Gambar {j+1}", type=["png", "jpg", "jpeg"], key=f"berkas_{j}")
            cap_g_id = st.text_input(f"Keterangan Gambar {j+1} (ID)", key=f"cap_g_id_{j}")
            cap_g_en = st.text_input(f"Figure {j+1} caption (EN)", key=f"cap_g_en_{j}")
            lebar = st.slider(f"Lebar Gambar {j+1} (cm)", 4.0, 7.6, 7.6, 0.2, key=f"lebar_{j}")
            daftar_gambar.append({"nomor": j+1, "blob": berkas.getvalue() if berkas else b"", "cap_id": cap_g_id, "cap_en": cap_g_en, "lebar": lebar})

with tab5:
    konflik = st.text_area("Konflik kepentingan", value="Penulis menyatakan tidak ada konflik kepentingan.")
    dana = st.text_area("Sumber dana", value="Tidak berlaku / Not applicable.")
    ucapan = st.text_area("Ucapan terima kasih", value="Penulis berterima kasih kepada Laboratorium Biologi Laut.")
    kontribusi = st.text_area("Kontribusi penulis (CRediT)", value="FSIM: conceptualization; ABR: formal analysis.")
    data_avail = st.text_area("Ketersediaan data", value="Dataset tersedia dari penulis korespondensi.")
    etik = st.text_area("Persetujuan etik", value="Tidak berlaku / Not applicable.")
    orcid_list = st.text_area("ORCID", value="Febry S. I. Menajang — https://orcid.org/0000-0000-0000-0000")
    daftar_pustaka = st.text_area("Daftar Pustaka (satu entri per baris)", height=200, value="Rondonuwu, A. B. (2025). Diversity. FAS, 28(10), 667–676.")

# UNDUH
st.markdown("---")
st.subheader(f"📥 Unduh Naskah ({versi})")
tbl_1 = daftar_tabel[0] if daftar_tabel else {"cap_id": " ", "cap_en": " ", "data": " ", "catatan": " "}
gbr_1 = daftar_gambar[0] if daftar_gambar else {"cap_id": " ", "cap_en": " ", "blob": b"", "lebar": 7.6}

naskah = Naskah(
    judul_id=judul_id, judul_en=judul_en, running_title=running_title,
    penulis_list=penulis_list, afiliasi_list=daftar_afiliasi, telepon=telepon,
    email_korespondensi=email, abstrak_id=abstrak_id, abstrak_en=abstract_en,
    kata_kunci_id=kata_kunci_id, kata_kunci_en=keywords_en,
    bab1=bab1, metode_21=metode_21, metode_22=metode_22, metode_221=metode_221, metode_23=metode_23,
    bab3=bab3, bab4=bab4, bab5=bab5,
    tabel_list=daftar_tabel, gambar_list=daftar_gambar,
    tbl_1=tbl_1, gbr_1=gbr_1,
    konflik=konflik, dana=dana, ucapan=ucapan, kontribusi=kontribusi,
    data_avail=data_avail, etik=etik, orcid_list=orcid_list, daftar_pustaka=daftar_pustaka,
    bahasa=BAHASA, blind=is_blind, **eq_inputs
)

if not template_path:
    st.error("Template resmi belum tersedia.")
    st.stop()

try:
    berkas = bangun(naskah, template_path)
    st.download_button(f"📄 Unduh Dokumen ({versi})", data=berkas, file_name="Naskah_PLATAX_2026.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
except Exception as e:
    st.error(f"Gagal menyusun dokumen: {e}")
