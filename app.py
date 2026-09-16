"""
app.py — Antarmuka Streamlit untuk JIP 2026 Article Auto-Generator
Versi lengkap dengan Tab Halaman Judul (Title Page)
"""

import os
import re
import streamlit as st
from platax_builder import Naskah, Penulis, bangun
from title_page_builder import bangun_title_page, get_credit_roles

# ── Nama file template sesuai OJS (tidak boleh diubah) ──────────────────────
TEMPLATE_ID_FINAL      = "Template_Artikel_PLATAX_2026_ID.docx"
TEMPLATE_EN_FINAL      = "Template_Article_PLATAX_2026_EN.docx"
TEMPLATE_ID_BLIND      = "Template_Naskah_BlindReview_JIP_2026.docx"
TEMPLATE_EN_BLIND      = "Template_Manuscript_BlindReview_JIP_2026.docx"
TEMPLATE_TITLE_PAGE    = "Template_Title_Page_JIP_2026.docx"

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

def auto_singkatan(nama: str) -> str:
    """'Febry S. I. Menajang' → 'F.S.I.M.'  (inisial semua kata + titik)"""
    parts = [p.strip(".") for p in nama.strip().split() if p.strip(".")]
    if not parts:
        return ""
    return ".".join(p[0].upper() for p in parts) + "."

versi = st.sidebar.radio("Versi dokumen:", ["Blind Review", "Final (Lengkap)"], index=0)
is_blind = versi == "Blind Review"
st.sidebar.info("📌 Blind Review: Identitas disembunyikan otomatis.\nFinal: Lengkap.")
st.sidebar.markdown("---")

if is_blind:
    template_aktif = TEMPLATE_EN_BLIND if EN else TEMPLATE_ID_BLIND
else:
    template_aktif = TEMPLATE_EN_FINAL if EN else TEMPLATE_ID_FINAL

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

# Title Page template status
if os.path.exists(TEMPLATE_TITLE_PAGE):
    st.sidebar.success(f"✅ {TEMPLATE_TITLE_PAGE}")
else:
    st.sidebar.warning(f"⚠️ Template tidak ditemukan: {TEMPLATE_TITLE_PAGE}")

st.title("📄 JIP 2026 Article Auto-Generator")

# ── NAMA HEADING BAB (bilingual) ─────────────────────────────────────────────
H = {
    "bab1":   L("1. Pendahuluan", "1. Introduction"),
    "bab2":   L("2. Bahan dan Metode", "2. Materials and Methods"),
    "met21":  L("2.1. Waktu dan lokasi penelitian", "2.1. Study period and location"),
    "met22":  L("2.2. Pengumpulan data", "2.2. Data collection"),
    "met221": L("2.2.1. Analisis laboratorium", "2.2.1. Laboratory analysis"),
    "met23":  L("2.3. Analisis data", "2.3. Data analysis"),
    "bab3":   L("3. Hasil", "3. Results"),
    "bab4":   L("4. Pembahasan", "4. Discussion"),
    "bab5":   L("5. Simpulan", "5. Conclusion"),
}

# ── FORMULIR ─────────────────────────────────────────────────────────────────
tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏷️ Halaman Judul",
    "📌 Identitas & Judul", "📝 Abstract & Abstrak",
    "📄 Bab 1–5", "🗂 Tabel & Gambar", "🤝 Pernyataan & Pustaka"
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1: IDENTITAS & JUDUL (harus sebelum tab0 karena tab0 memakai variabel ini)
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns(2)
    with c1:
        judul_id = st.text_input(L("Judul Bahasa Indonesia", "Judul Bahasa Indonesia"),
                                 "Keanekaragaman Ikan Gobi di Muara Sungai Tondano")
    with c2:
        judul_en = st.text_input(L("Title in English", "Title in English"),
                                 "Diversity of Goby Fish in Tondano River Estuary")
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
    penulis_list, catatan_penulis = susun_penulis(tabel_penulis, daftar_afiliasi,
                                                   peta=peta if nomor_dirujuk else None)
    for c in catatan_penulis:
        st.warning(c)
    c3, c4 = st.columns(2)
    with c3:
        telepon = st.text_input(L("Telepon korespondensi", "Corresponding author phone"), "+62-8XXXXXXXXXX")
    with c4:
        email = st.text_input(L("Surel korespondensi", "Corresponding author e-mail"), "nama@unsrat.ac.id")
    ucapan = st.text_area(L("Ucapan terima kasih", "Acknowledgements"),
                          value=L("Penulis berterima kasih kepada Laboratorium Biologi Laut.",
                                  "The authors thank the Marine Biology Laboratory."),
                          height=60, key="ucapan_main")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 0: HALAMAN JUDUL / TITLE PAGE
# ═══════════════════════════════════════════════════════════════════════════════
with tab0:
    st.markdown(
        """
        <div style="background:#E8F4FB;border-left:4px solid #007BB8;
                    padding:10px 14px;border-radius:4px;margin-bottom:1rem;">
        <b>ℹ️ Halaman Judul (Title Page)</b> adalah berkas <em>terpisah</em> yang
        diunggah sebagai <em>Supplementary File</em> di OJS — bukan bagian naskah utama.
        Isi data di bawah, lalu unduh file <code>.docx</code>-nya.
        </div>
        """, unsafe_allow_html=True
    )

    st.subheader("0. ID Naskah OJS / OJS Submission ID")
    tp_ojs_id = st.text_input("ID Naskah OJS", value="PLATAX-2026-XXXX",
                               key="tp_ojs_id",
                               help="Contoh: PLATAX-2026-0042")

    st.subheader("1. Judul / Title")
    tp_judul_id  = st.text_input("Judul (Bahasa Indonesia)", value=judul_id, key="tp_judul_id")
    tp_judul_en  = st.text_input("Title (English)", value=judul_en, key="tp_judul_en")
    tp_running   = st.text_input(
        "Judul Singkat / Running Title (maks. 60 karakter)",
        value=running_title, max_chars=60, key="tp_running"
    )
    st.caption(f"{len(tp_running)}/60 karakter")

    st.markdown("---")
    st.subheader("2. Penulis dan Afiliasi / Authors and Affiliations")

    # Bangun nilai awal dari tab1
    _tp_init_rows = []
    for p in penulis_list:
        _nama = p.nama if hasattr(p, "nama") else ""
        _tp_init_rows.append({
            "Nama": _nama,
            "Singkatan Nama": auto_singkatan(_nama),
            "Afiliasi (huruf: a, b, ...)": "a",
            "Email": p.email if hasattr(p, "email") else "",
            "Korespondensi ★": p.is_corresp if hasattr(p, "is_corresp") else False,
            "ORCID iD": p.orcid if hasattr(p, "orcid") else "",
        })
    if not _tp_init_rows:
        _tp_init_rows = [{
            "Nama": "", "Singkatan Nama": "",
            "Afiliasi (huruf: a, b, ...)": "a", "Email": "",
            "Korespondensi ★": True, "ORCID iD": "0000-0000-0000-0000",
        }]

    tp_tabel_penulis = st.data_editor(
        _tp_init_rows,
        num_rows="dynamic", width="stretch", key="tp_tabel_penulis",
        column_config={
            "Korespondensi ★": st.column_config.CheckboxColumn(),
            "Singkatan Nama": st.column_config.TextColumn(
                help="Inisial semua kata + titik, misal: F.S.I.M. — otomatis, bisa diedit manual"
            ),
        }
    )

    huruf_aff_dirujuk = sorted({
        h.strip().lower()
        for baris in as_records(tp_tabel_penulis)
        if str(baris.get("Nama", "")).strip()
        for h in re.split(r"[,\s]+", str(baris.get("Afiliasi (huruf: a, b, ...)", "") or ""))
        if h.strip() and h.strip().isalpha()
    })

    tp_afiliasi = {}
    if huruf_aff_dirujuk:
        st.markdown("**Keterangan Afiliasi / Affiliation Details:**")
        for huruf in huruf_aff_dirujuk:
            default_aff = ""
            if huruf == "a" and daftar_afiliasi:
                default_aff = daftar_afiliasi[0]
            elif huruf == "b" and len(daftar_afiliasi) > 1:
                default_aff = daftar_afiliasi[1]
            tp_afiliasi[huruf] = st.text_area(
                f"({huruf})", value=default_aff, height=60, key=f"tp_aff_{huruf}"
            )

    st.markdown("---")
    st.subheader("3. Penulis Korespondensi / Corresponding Author")

    _koresp_obj = next((p for p in penulis_list
                        if hasattr(p, "is_corresp") and p.is_corresp), None)

    c_kol1, c_kol2 = st.columns(2)
    with c_kol1:
        tp_koresp_nama   = st.text_input("Nama Lengkap / Full Name",
                                          value=_koresp_obj.nama if _koresp_obj else "",
                                          key="tp_koresp_nama")
        tp_koresp_aff    = st.text_input("Afiliasi / Affiliation",
                                          value=(list(tp_afiliasi.values())[0] if tp_afiliasi else ""),
                                          key="tp_koresp_aff")
        tp_koresp_alamat = st.text_input("Alamat / Address", value="", key="tp_koresp_alamat")
    with c_kol2:
        tp_koresp_telp  = st.text_input("Telepon / Phone", value=telepon, key="tp_koresp_telp")
        tp_koresp_email = st.text_input("E-mail", value=email, key="tp_koresp_email")
        tp_koresp_orcid = st.text_input(
            "ORCID iD",
            value=(_koresp_obj.orcid if _koresp_obj else "https://orcid.org/0000-0000-0000-0000"),
            key="tp_koresp_orcid"
        )

    st.markdown("---")
    st.subheader("4. Kontribusi Penulis / Author Contributions (CRediT)")
    st.caption("Pilih peran untuk setiap penulis.")

    CREDIT_ROLES = get_credit_roles()
    tp_kontribusi = {}

    # Bangun list penulis + singkatan dari tabel editor
    penulis_tp_rows = [
        {
            "nama": str(b.get("Nama", "")).strip(),
            "singkatan": (
                str(b.get("Singkatan Nama", "")).strip()
                or auto_singkatan(str(b.get("Nama", "")).strip())
            ),
        }
        for b in as_records(tp_tabel_penulis)
        if str(b.get("Nama", "")).strip()
    ]

    for idx_p, p_row in enumerate(penulis_tp_rows):
        nama_p      = p_row["nama"]
        singkatan_p = p_row["singkatan"]

        # Label expander: nama lengkap — singkatan
        exp_label = f"📝  {nama_p}"
        if singkatan_p:
            exp_label += f"  —  {singkatan_p}"

        with st.expander(exp_label, expanded=(idx_p == 0)):

            # ── Singkatan nama (editable) ──────────────────────────────────
            col_sk, col_del = st.columns([11, 1])
            singkatan_edit = col_sk.text_input(
                "Singkatan Nama / Abbreviated Name",
                value=singkatan_p,
                key=f"tp_singkatan_{idx_p}",
                placeholder="misal: F.S.I.M.",
                help=(
                    "Inisial semua kata dalam nama + titik. "
                    "Contoh: 'Febry S. I. Menajang' → 'F.S.I.M.' · "
                    "'Vera O. I. Kumaat' → 'V.O.I.K.'"
                ),
            )
            # Gunakan hasil edit; fallback ke auto
            singkatan_p = singkatan_edit or auto_singkatan(nama_p)

            st.markdown("---")

            # ── CRediT checkboxes (3 kolom) ────────────────────────────────
            cols_credit = st.columns(3)
            selected_roles = []
            for idx_r, role in enumerate(CREDIT_ROLES):
                col_idx = idx_r % 3
                if cols_credit[col_idx].checkbox(role, key=f"tp_credit_{idx_p}_{idx_r}"):
                    selected_roles.append(role)

            extra_roles = st.text_input(
                "Tambahan (pisah koma) / Additional roles",
                key=f"tp_credit_extra_{idx_p}",
                placeholder="misal: Project administration",
            )
            all_roles = selected_roles + [r.strip() for r in extra_roles.split(",") if r.strip()]

            # Simpan dengan kunci = singkatan (sesuai format dokumen)
            key_out = singkatan_p or nama_p
            tp_kontribusi[key_out] = ", ".join(all_roles) if all_roles else ""

            # ── Pratinjau output ───────────────────────────────────────────
            if all_roles:
                st.caption(
                    f"**Output:** {key_out}: "
                    + ", ".join(r.lower() for r in all_roles) + "."
                )

    st.markdown("---")
    st.subheader("5. Ucapan Terima Kasih / Acknowledgements")
    tp_ucapan = st.text_area("Ucapan terima kasih", value=ucapan, height=80, key="tp_ucapan")

    st.subheader("6. Pernyataan Pendanaan / Funding Statement")
    tp_pendanaan_mode = st.radio(
        "Jenis pendanaan:",
        ["Hibah / Grant-funded", "Dana mandiri / Self-funded", "Kustom / Custom"],
        horizontal=True, key="tp_pendanaan_mode"
    )
    if tp_pendanaan_mode == "Dana mandiri / Self-funded":
        tp_pendanaan = ("Penelitian ini menggunakan dana mandiri penulis. / "
                        "This research was self-funded by the authors.")
        st.info(tp_pendanaan)
    elif tp_pendanaan_mode == "Hibah / Grant-funded":
        _nama_lembaga   = st.text_input("Nama Lembaga / Agency Name", key="tp_lembaga")
        _nomor_kontrak  = st.text_input("Nomor Kontrak / Contract Number", key="tp_kontrak")
        tp_pendanaan = (
            f"Penelitian ini didanai oleh {_nama_lembaga} dengan nomor kontrak {_nomor_kontrak}. / "
            f"This study was funded by {_nama_lembaga} under contract number {_nomor_kontrak}."
        )
    else:
        tp_pendanaan = st.text_area("Pernyataan pendanaan", value="", height=80,
                                    key="tp_pendanaan_custom")

    st.subheader("7. Konflik Kepentingan / Conflict of Interest")
    tp_konflik_mode = st.radio(
        "Status konflik:",
        ["Tidak ada konflik / No conflict", "Ada konflik / Conflict exists"],
        horizontal=True, key="tp_konflik_mode"
    )
    if tp_konflik_mode == "Tidak ada konflik / No conflict":
        tp_konflik = ("Para penulis menyatakan tidak ada konflik kepentingan yang relevan "
                      "dengan artikel ini. / The authors declare no conflict of interest "
                      "relevant to this article.")
        st.info(tp_konflik)
    else:
        tp_konflik = st.text_area("Uraikan konflik kepentingan", value="", height=80,
                                   key="tp_konflik_custom")

    # ── Tombol Download Title Page ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📥 Unduh Halaman Judul / Download Title Page")

    if not os.path.exists(TEMPLATE_TITLE_PAGE):
        st.error(f"❌ Template Title Page tidak ditemukan: {TEMPLATE_TITLE_PAGE}")
    else:
        _tp_penulis = [
            {
                "nama":       str(b.get("Nama", "")).strip(),
                "singkatan":  (
                    str(b.get("Singkatan Nama", "")).strip()
                    or auto_singkatan(str(b.get("Nama", "")).strip())
                ),
                "aff":        str(b.get("Afiliasi (huruf: a, b, ...)", "")).strip(),
                "email":      str(b.get("Email", "")).strip(),
                "is_corresp": bool(b.get("Korespondensi ★", False)),
                "orcid":      str(b.get("ORCID iD", "")).strip() or "—",
            }
            for b in as_records(tp_tabel_penulis)
            if str(b.get("Nama", "")).strip()
        ]

        _tp_data = {
            "ojs_id":          tp_ojs_id,
            "judul_id":        tp_judul_id,
            "judul_en":        tp_judul_en,
            "running_title":   tp_running,
            "penulis":         _tp_penulis,
            "afiliasi":        tp_afiliasi,
            "koresp_nama":     tp_koresp_nama,
            "koresp_afiliasi": tp_koresp_aff,
            "koresp_alamat":   tp_koresp_alamat,
            "koresp_telepon":  tp_koresp_telp,
            "koresp_email":    tp_koresp_email,
            "koresp_orcid":    tp_koresp_orcid,
            "kontribusi":      tp_kontribusi,
            "ucapan":          tp_ucapan,
            "pendanaan":       tp_pendanaan,
            "konflik":         tp_konflik,
        }

        try:
            _tp_bytes = bangun_title_page(_tp_data, template_path=TEMPLATE_TITLE_PAGE)
            st.download_button(
                label="📄 Unduh Title Page — Template_Title_Page_JIP_2026.docx",
                data=_tp_bytes,
                file_name="Template_Title_Page_JIP_2026.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
            st.success("✅ Title Page siap diunduh. Unggah sebagai **Supplementary File** di OJS.")
        except Exception as _e:
            st.error(f"Gagal menyusun Title Page: {_e}")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2: ABSTRACT & ABSTRAK
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    abstract_en = st.text_area("Abstract (English)", height=170, value="This study analysed...")
    keywords_en = st.text_input("Keywords (comma-separated)", "Goby fish, Diversity, Estuary")
    abstrak_id  = st.text_area("Abstrak (Bahasa Indonesia)", height=170, value="Penelitian ini menganalisis...")
    kata_kunci_id = st.text_input("Kata kunci (dipisah koma)", "Ikan gobi, Keanekaragaman, Estuari")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3: BAB 1–5
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    bab1       = st.text_area(H["bab1"],   height=140, value=L(
        "Muara Sungai Tondano memiliki peranan ekologis penting (Carpenter & Niem, 1998).",
        "Tondano River Estuary has important ecological roles (Carpenter & Niem, 1998)."))
    metode_21  = st.text_area(H["met21"],  height=80,  value=L(
        "Penelitian dilaksanakan Januari–Maret 2026...",
        "The study was conducted from January to March 2026..."))
    metode_22  = st.text_area(H["met22"],  height=80,  value=L(
        "Pengambilan sampel menggunakan jaring insang...",
        "Sampling was carried out using gill nets..."))
    metode_221 = st.text_area(H["met221"], height=80,  value=L(
        "Identifikasi morfometrik dilakukan...",
        "Morphometric identification was conducted..."))
    metode_23  = st.text_area(H["met23"],  height=80,  value=L(
        "Indeks Shannon-Wiener dihitung...",
        "The Shannon-Wiener index was calculated..."))
    bab3       = st.text_area(H["bab3"],   height=120, value=L(
        "Tercatat 12 jenis dari tiga stasiun (Tabel 1).",
        "Twelve species were recorded across three stations (Table 1)."))
    bab4       = st.text_area(H["bab4"],   height=140, value=L(
        "Tingginya keanekaragaman diduga berkaitan...",
        "The high diversity is presumably related to..."))
    bab5       = st.text_area(H["bab5"],   height=90,  value=L(
        "Keanekaragaman ikan gobi tergolong sedang...",
        "Goby fish diversity was classified as moderate..."))

    st.markdown("---")
    st.subheader(L("📐 Persamaan Matematika", "📐 Mathematical Equations"))
    st.caption(L("Gunakan sintaks LaTeX. Contoh: `C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}`",
                 "Use LaTeX syntax. Example: `C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}`"))
    n_eq = st.number_input(L("Jumlah persamaan", "Number of equations"),
                           min_value=0, max_value=20, value=0, key="n_eq")
    eq_inputs = {}
    if int(n_eq) > 0:
        for i in range(int(n_eq)):
            eq_num = i + 1
            latex  = st.text_area(
                L(f"Persamaan {eq_num} (LaTeX)", f"Equation {eq_num} (LaTeX)"),
                placeholder="Contoh / Example: C_b = \\frac{\\nabla}{L \\cdot B \\cdot d}",
                height=80, key=f"eq_{eq_num}"
            )
            if latex.strip():
                eq_inputs[f'eq_{eq_num}'] = latex

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4: TABEL & GAMBAR
# ═══════════════════════════════════════════════════════════════════════════════
with tab4:
    n_tabel = st.number_input(L("Jumlah Tabel", "Number of Tables"),
                              min_value=0, max_value=20, value=1, key="n_tabel")
    daftar_tabel = []
    for i in range(int(n_tabel)):
        with st.expander(f"{L('Tabel', 'Table')} {i+1}", expanded=(i == 0)):
            cap = st.text_input(L(f"Keterangan Tabel {i+1}", f"Table {i+1} caption"),
                                value=L("Parameter kualitas air.", "Water quality parameters."),
                                key=f"cap_{i}")
            t_data = st.text_area(
                L(f"Isi Tabel {i+1} (pisah kolom dengan ;)", f"Table {i+1} data (separate columns with ;)"),
                value=L("Parameter;Stasiun 1\nSuhu;28,4", "Parameter;Station 1\nTemperature;28.4"),
                key=f"t_data_{i}"
            )
            cat = st.text_input(L(f"Catatan kaki Tabel {i+1}", f"Table {i+1} footnote"),
                                value=L("Nilai rerata ± SD.", "Mean values ± SD."), key=f"cat_{i}")
            daftar_tabel.append({"nomor": i+1, "cap": cap, "data": t_data, "catatan": cat})

    n_gambar = st.number_input(L("Jumlah Gambar", "Number of Figures"),
                               min_value=0, max_value=20, value=1, key="n_gambar")
    daftar_gambar = []
    for j in range(int(n_gambar)):
        with st.expander(f"{L('Gambar', 'Figure')} {j+1}", expanded=(j == 0)):
            berkas = st.file_uploader(L(f"Unggah Gambar {j+1}", f"Upload Figure {j+1}"),
                                      type=["png", "jpg", "jpeg"], key=f"berkas_{j}")
            cap_g  = st.text_input(L(f"Keterangan Gambar {j+1}", f"Figure {j+1} caption"),
                                   key=f"cap_g_{j}")
            lebar  = st.slider(L(f"Lebar Gambar {j+1} (cm)", f"Figure {j+1} width (cm)"),
                               4.0, 7.6, 7.6, 0.2, key=f"lebar_{j}")
            daftar_gambar.append({"nomor": j+1, "blob": berkas.getvalue() if berkas else b"",
                                  "cap": cap_g, "lebar": lebar})

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5: PERNYATAAN & PUSTAKA
# ═══════════════════════════════════════════════════════════════════════════════
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

# ── UNDUH NASKAH ─────────────────────────────────────────────────────────────
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
