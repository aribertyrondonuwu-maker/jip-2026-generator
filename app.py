"""
JIP 2026 Article Auto-Generator — Jurnal Ilmiah PLATAX

Antarmuka Streamlit. Seluruh penyusunan dokumen dikerjakan platax_builder.py,
yang mengisi Template_Artikel_PLATAX_2026_FINAL_OJS.docx secara langsung.

Letakkan berkas berikut dalam satu folder:
    app.py
    platax_builder.py
    Template_Artikel_PLATAX_2026_FINAL_OJS.docx
"""

import os
import re

import streamlit as st

from platax_builder import Naskah, Penulis, bangun

TEMPLATE_DEFAULT = "Template_Artikel_PLATAX_2026_FINAL_OJS.docx"

st.set_page_config(page_title="JIP 2026 Article Auto-Generator", page_icon="📄", layout="wide")


# =========================================================
# UTILITAS PENULIS
# =========================================================
def as_records(tabel):
    return tabel.to_dict("records") if hasattr(tabel, "to_dict") else list(tabel)


def susun_penulis(tabel_penulis, daftar_afiliasi, peta=None):
    """
    Ubah tabel input menjadi daftar Penulis + catatan validasi.

    `peta` memetakan nomor yang diketik penulis ke nomor rapat 1..n pada
    `daftar_afiliasi`, sehingga penulis boleh menulis 1 dan 3 tanpa membuat
    lubang penomoran di naskah.
    """
    penulis, catatan = [], []
    for baris in as_records(tabel_penulis):
        nama = str(baris.get("Nama", "") or "").strip()
        if not nama:
            continue
        mentah = [int(x) for x in re.findall(r"\d+", str(baris.get("Afiliasi") or ""))]
        if not mentah:
            catatan.append(f"Penulis “{nama}” belum diberi nomor afiliasi.")
        nomor = [peta.get(x, x) if peta else x for x in mentah]
        penulis.append(Penulis(
            nama=nama,
            afil=[x for x in nomor if 1 <= x <= len(daftar_afiliasi)],
            koresp=bool(baris.get("Korespondensi", False)),
            orcid=str(baris.get("ORCID", "") or "").strip(),
        ))

    if not penulis:
        catatan.append("Belum ada penulis yang diisi.")
    n_koresp = sum(1 for p in penulis if p.koresp)
    if n_koresp == 0:
        catatan.append("Belum ada penulis korespondensi yang ditandai.")
    elif n_koresp > 1:
        catatan.append("Lebih dari satu penulis korespondensi ditandai.")

    for i, teks in enumerate(daftar_afiliasi, 1):
        if not teks.strip():
            catatan.append(f"Kotak Afiliasi {i} masih kosong.")
    return penulis, catatan


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("⚙️ Pengaturan")
bahasa_label = st.sidebar.selectbox(
    "🌐 Bahasa naskah / Manuscript language:",
    ["Bahasa Indonesia", "English"], index=0,
    help="Menentukan bahasa judul bagian pada dokumen Word. Judul, abstrak, dan kata "
         "kunci tetap wajib dwibahasa apa pun pilihannya.")
BAHASA = "en" if bahasa_label == "English" else "id"
EN = BAHASA == "en"
st.sidebar.caption(
    "Naskah Inggris: judul bagian menjadi Introduction, Results, References, dan seterusnya; "
    "baris terjemahan pada keterangan tabel/gambar dihapus (Petunjuk §1)."
    if EN else
    "Naskah Indonesia: keterangan tabel dan gambar wajib disertai terjemahan Inggris "
    "pada baris kedua (Petunjuk §1).")
st.sidebar.markdown("---")


def L(teks_id: str, teks_en: str) -> str:
    """Pilih teks antarmuka sesuai bahasa naskah."""
    return teks_en if EN else teks_id


versi = st.sidebar.radio("Versi dokumen keluaran:",
                         ["Blind Review", "Final (Lengkap)"], index=0,
                         help="Blind Review untuk pengiriman awal; "
                              "Final hanya setelah naskah dinyatakan ACCEPTED.")
is_blind = versi == "Blind Review"
st.sidebar.info("📌 **Blind Review**: nama penulis, afiliasi, surel, ucapan terima kasih, "
                "ORCID, dan metadata Word dikosongkan otomatis (Petunjuk §9).\n\n"
                "📌 **Final**: seluruh identitas dimunculkan.")

st.sidebar.markdown("---")
st.sidebar.caption("Template resmi")
unggahan_template = None
template_path = None
if os.path.exists(TEMPLATE_DEFAULT):
    template_path = TEMPLATE_DEFAULT
    st.sidebar.success(f"✅ {TEMPLATE_DEFAULT}")
else:
    st.sidebar.warning("Template resmi tidak ditemukan di folder aplikasi.")
    unggahan_template = st.sidebar.file_uploader("Unggah template .docx", type=["docx"])

st.title("📄 JIP 2026 Article Auto-Generator")
st.write(L("Aplikasi mengisi **template resmi PLATAX 2026** secara langsung, sehingga masthead, "
           "logo, kotak abstrak, margin, header, footer, dan seluruh tipografi mengikuti "
           "template apa adanya.",
           "The app fills the **official PLATAX 2026 template** directly, so the masthead, logo, "
           "abstract boxes, margins, header, footer, and all typography follow the template "
           "exactly as issued."))

# =========================================================
# FORMULIR
# =========================================================
tab1, tab2, tab3, tab5, tab4 = st.tabs(
    ["📌 Identitas & Judul", "📝 Abstract & Abstrak", "📄 Bab 1–5",
     "📊 Tabel & Gambar", "🤝 Pernyataan & Pustaka"] if not EN else
    ["📌 Title & Authors", "📝 Abstracts", "📄 Sections 1–5",
     "📊 Tables & Figures", "🤝 Statements & References"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        judul_id = st.text_input(L("Judul Bahasa Indonesia (maks. 20 kata)",
                                   "Indonesian title (max. 20 words)"),
                                 "Keanekaragaman Ikan Gobi di Muara Sungai Tondano, Sulawesi Utara")
    with c2:
        judul_en = st.text_input(L("Title in English (max. 20 words)",
                                   "English title (max. 20 words)"),
                                 "Diversity of Goby Fish in Tondano River Estuary, North Sulawesi")
    running_title = st.text_input(
        "Running title (max. 60 characters)" if EN
        else "Judul singkat / Running title (maks. 60 karakter)",
        "Goby Diversity in Tondano River Estuary" if EN
        else "Keanekaragaman Ikan Gobi Muara Tondano",
        max_chars=60,
        help="Ditulis dalam bahasa naskah — satu versi saja, karena dipakai pada header "
             "halaman genap yang hanya menyediakan satu ruas teks (Petunjuk §4).")
    st.caption(f"{len(running_title)}/60 " + ("characters" if EN else "karakter")
               + (" · ikuti bahasa naskah, tidak perlu dwibahasa" if not EN else ""))

    st.markdown("---")
    st.subheader(L("Informasi Penulis", "Author Information"))
    if is_blind:
        st.info("ℹ️ Mode Blind Review aktif: data di bawah tetap boleh diisi, tetapi tidak "
                "dimunculkan dalam berkas Word. Gunakan untuk menyusun Title Page terpisah.")

    st.markdown(L("**1. Daftar penulis** — kolom *No. afiliasi* cukup diisi angka urut "
                  "(1, 2, 3 …). Penulis dengan dua afiliasi ditulis `1,2`. Kotak isian "
                  "afiliasinya muncul otomatis di bawah.",
                  "**1. Author list** — the *Affiliation no.* column takes plain numbers "
                  "(1, 2, 3 …). An author with two affiliations is written `1,2`. The "
                  "affiliation fields appear automatically below."))
    tabel_penulis = st.data_editor(
        [
            {"Nama": "Febry S. I. Menajang", "Afiliasi": "1", "Korespondensi": True,
             "ORCID": "https://orcid.org/0000-0000-0000-0000"},
            {"Nama": "Ari B. Rondonuwu", "Afiliasi": "1", "Korespondensi": False, "ORCID": ""},
            {"Nama": "John L. Tombokan", "Afiliasi": "2", "Korespondensi": False, "ORCID": ""},
            {"Nama": "Veibe Warouw", "Afiliasi": "1,2", "Korespondensi": False, "ORCID": ""},
        ],
        num_rows="dynamic", width="stretch", key="tabel_penulis",
        column_config={
            "Nama": st.column_config.TextColumn(
                L("Nama penulis", "Author name"), required=True, width="large",
                help=L("Nama lengkap tanpa gelar. Jangan ketik angka superskrip di sini.",
                       "Full name without titles. Do not type superscript numbers here.")),
            "Afiliasi": st.column_config.TextColumn(
                L("No. afiliasi", "Affiliation no."), width="small",
                help=L("Contoh: 1 atau 1,2", "Example: 1 or 1,2")),
            "Korespondensi": st.column_config.CheckboxColumn(
                L("Koresp.", "Corresp."), width="small"),
            "ORCID": st.column_config.TextColumn("ORCID iD", width="medium"),
        })

    # Nomor afiliasi yang benar-benar dirujuk penulis menentukan jumlah kotak isian
    nomor_dirujuk = sorted({
        int(x)
        for baris in as_records(tabel_penulis)
        if str(baris.get("Nama", "") or "").strip()
        for x in re.findall(r"\d+", str(baris.get("Afiliasi") or ""))
    })

    st.markdown(L("**2. Afiliasi** — satu kotak per nomor yang dirujuk di atas. "
                  "Penulis se-institusi cukup memakai nomor yang sama.",
                  "**2. Affiliations** — one field per number referenced above. Authors from "
                  "the same institution simply share the same number."))
    CONTOH_AFILIASI = [
        "Program Studi Ilmu Kelautan, Fakultas Perikanan dan Ilmu Kelautan, "
        "Universitas Sam Ratulangi, Manado, 95115, Indonesia",
        "Program Studi Manajemen Sumberdaya Perairan, Fakultas Perikanan dan Ilmu Kelautan, "
        "Universitas Sam Ratulangi, Manado, 95115, Indonesia",
    ]
    if not nomor_dirujuk:
        st.caption(L("Belum ada nomor afiliasi pada tabel penulis.",
                     "No affiliation number entered in the author table yet."))
        daftar_afiliasi = []
    else:
        if nomor_dirujuk != list(range(1, len(nomor_dirujuk) + 1)):
            st.warning(f"Nomor afiliasi sebaiknya berurutan mulai dari 1. "
                       f"Saat ini terpakai: {nomor_dirujuk}.")
        isian = {}
        for nomor in nomor_dirujuk:
            contoh = CONTOH_AFILIASI[nomor - 1] if nomor <= len(CONTOH_AFILIASI) else ""
            isian[nomor] = st.text_area(
                L(f"Afiliasi {nomor}  (Departemen, Fakultas, Universitas, Kota, Kode Pos, Negara)",
                  f"Affiliation {nomor}  (Department, Faculty, University, City, Postcode, Country)"),
                value=contoh, height=68, key=f"afiliasi_{nomor}")
        # Daftar dipadatkan menjadi 1..n agar penomoran pada naskah selalu rapat
        peta = {lama: baru for baru, lama in enumerate(nomor_dirujuk, 1)}
        daftar_afiliasi = [isian[lama].strip() for lama in nomor_dirujuk]

    penulis_list, catatan_penulis = susun_penulis(tabel_penulis, daftar_afiliasi,
                                                  peta=peta if nomor_dirujuk else None)
    for c in catatan_penulis:
        st.warning(c)
    if penulis_list:
        st.caption(L("Pratinjau baris penulis: ", "Author line preview: ") + ", ".join(
            p.nama + ("^" + ",".join(map(str, p.afil)) if p.afil else "")
            + ("*" if p.koresp else "") for p in penulis_list))

    c3, c4 = st.columns(2)
    with c3:
        telepon = st.text_input(L("Telepon korespondensi", "Corresponding author phone"),
                                "+62-8XXXXXXXXXX")
    with c4:
        email = st.text_input(L("Surel korespondensi", "Corresponding author e-mail"),
                              "nama@unsrat.ac.id")

with tab2:
    st.caption(L("Urutan template: Abstract (EN) lebih dahulu, baru Abstrak (ID). "
                 "150–250 kata per bahasa, satu paragraf naratif, tanpa sitasi.",
                 "Template order: Abstract (EN) first, then Abstrak (ID). 150–250 words per "
                 "language, a single narrative paragraph, no citations."))
    abstract_en = st.text_area("Abstract (English)", height=170,
                               value="This study analysed the diversity of goby fish...")
    keywords_en = st.text_input(L("Keywords (5, pisahkan dengan koma)",
                                  "Keywords (5, comma-separated)"),
                                "Goby fish, Diversity, Tondano River, Estuary, North Sulawesi")
    st.caption(f"Abstract: {len(abstract_en.split())} " + L("kata", "words"))
    abstrak_id = st.text_area("Abstrak (Bahasa Indonesia)", height=170,
                              value="Penelitian ini menganalisis keanekaragaman ikan gobi...")
    kata_kunci_id = st.text_input(L("Kata kunci (5, pisahkan dengan koma)",
                                    "Kata kunci / Indonesian keywords (5, comma-separated)"),
                                  "Ikan gobi, Keanekaragaman, Sungai Tondano, Estuari, Sulawesi Utara")
    st.caption(f"Abstrak: {len(abstrak_id.split())} " + L("kata", "words"))

with tab3:
    st.caption(L("Satu baris kosong = paragraf baru. Sitiran seperti (Froese & Pauly, 2024) "
                 "atau Rondonuwu et al. (2025) otomatis diwarnai biru 007BB8.",
                 "One line break = new paragraph. Citations such as (Froese & Pauly, 2024) or "
                 "Rondonuwu et al. (2025) are automatically coloured blue 007BB8."))
    bab1 = st.text_area(
        L("1. Pendahuluan", "1. Introduction"), height=140,
        value=L("Muara Sungai Tondano memiliki peranan ekologis penting "
                "(Carpenter & Niem, 1998).\n"
                "Rondonuwu et al. (2025) mencatat kebaruan jenis di kawasan ini.\n"
                "Penelitian ini bertujuan menilai keanekaragaman ikan gobi.",
                "The Tondano River estuary plays an important ecological role "
                "(Carpenter & Niem, 1998).\n"
                "Rondonuwu et al. (2025) reported a new record for the area.\n"
                "This study aims to assess goby diversity in the estuary."))
    st.markdown(L("**2. Bahan dan Metode**", "**2. Materials and Methods**"))
    metode_21 = st.text_area(
        L("2.1. Waktu dan lokasi penelitian", "2.1. Study period and location"), height=80,
        value=L("Penelitian dilaksanakan Januari–Maret 2026 pada tiga stasiun...",
                "The study was conducted from January to March 2026 at three stations..."))
    metode_22 = st.text_area(
        L("2.2. Pengumpulan data", "2.2. Data collection"), height=80,
        value=L("Pengambilan sampel menggunakan jaring insang dengan tiga ulangan...",
                "Sampling used gill nets with three replicates..."))
    metode_221 = st.text_area(
        L("2.2.1. Analisis laboratorium", "2.2.1. Laboratory analysis"), height=80,
        value=L("Identifikasi morfometrik dilakukan di laboratorium...",
                "Morphometric identification was carried out in the laboratory..."))
    metode_23 = st.text_area(
        L("2.3. Analisis data", "2.3. Data analysis"), height=80,
        value=L("Indeks Shannon-Wiener dihitung menggunakan PAST v4.03...",
                "The Shannon-Wiener index was computed using PAST v4.03..."))
    bab3 = st.text_area(
        L("3. Hasil", "3. Results"), height=120,
        value=L("Tercatat 12 jenis dari tiga stasiun pengamatan (Tabel 1).",
                "Twelve species were recorded across the three stations (Table 1)."))
    bab4 = st.text_area(
        L("4. Pembahasan", "4. Discussion"), height=140,
        value=L("Tingginya keanekaragaman diduga berkaitan dengan variasi salinitas...",
                "The high diversity is likely related to salinity variation..."))
    bab5 = st.text_area(
        L("5. Simpulan (maks. 150 kata)", "5. Conclusions (max. 150 words)"), height=90,
        value=L("Keanekaragaman ikan gobi tergolong sedang...",
                "Goby diversity in the estuary is moderate..."))
    st.caption(L(f"Simpulan: {len(bab5.split())} kata",
                 f"Conclusions: {len(bab5.split())} words"))

# =========================================================
# TAB 5: TABEL & GAMBAR (DINAMIS multi-tabel & multi-gambar)
# =========================================================
with tab5:
    st.caption(L("Keterangan tabel di ATAS tabel, keterangan gambar di BAWAH gambar. "
                 "Naskah Indonesia WAJIB menyertakan terjemahan Inggris pada baris kedua.",
                 "Table captions go ABOVE the table, figure captions BELOW the figure. "
                 "English manuscripts do not need the Indonesian translation line."))
    if EN:
        st.caption("Manuscript language is English: only the English caption is used; "
                   "the Indonesian field may be left blank.")

    # -----------------------------------------------------
    # BAGIAN TABEL
    # -----------------------------------------------------
    st.subheader(L("📊 Data Tabel", "📊 Tables"))
    col_tabel_cnt, _ = st.columns([1, 2])
    with col_tabel_cnt:
        n_tabel = st.number_input(
            L("Jumlah Tabel dalam artikel", "Number of Tables in article"),
            min_value=0, max_value=20, value=1, step=1, key="n_tabel"
        )

    daftar_tabel = []
    for i in range(int(n_tabel)):
        idx = i + 1
        with st.expander(L(f"📌 Tabel {idx}", f"📌 Table {idx}"), expanded=(i == 0)):
            cap_id = st.text_input(
                L(f"Keterangan Tabel {idx} (Indonesia)", f"Table {idx} caption (Indonesian — optional)"),
                value="Parameter kualitas air pada dua stasiun pengamatan." if idx == 1 else "",
                key=f"cap_tabel_id_{i}"
            )
            cap_en = st.text_input(
                L(f"Table {idx} caption (English)", f"Table {idx} caption (English)"),
                value="Water quality parameters at two observation stations." if idx == 1 else "",
                key=f"cap_tabel_en_{i}"
            )
            t_data = st.text_area(
                L(f"Isi Tabel {idx} — kolom dipisah tab atau titik koma; baris pertama = kepala tabel",
                  f"Table {idx} content — columns separated by tab or semicolon; first row = header"),
                height=110,
                value=L("Parameter;Stasiun 1;Stasiun 2\n"
                        "Suhu (°C);28,4 ± 0,3;29,1 ± 0,5\n"
                        "Salinitas (‰);32,1 ± 0,4;31,8 ± 0,6",
                        "Parameter;Station 1;Station 2\n"
                        "Temperature (°C);28.4 ± 0.3;29.1 ± 0.5\n"
                        "Salinity (‰);32.1 ± 0.4;31.8 ± 0.6") if idx == 1 else "",
                key=f"tabel_data_{i}"
            )
            cat = st.text_input(
                L(f"Keterangan kaki Tabel {idx}", f"Table {idx} footnote"),
                value=L("Keterangan: nilai merupakan rerata ± simpangan baku.",
                        "values are means ± standard deviation.") if idx == 1 else "",
                key=f"cat_tabel_{i}"
            )
            daftar_tabel.append({
                "nomor": idx,
                "cap_id": cap_id,
                "cap_en": cap_en,
                "data": t_data,
                "catatan": cat
            })

    st.markdown("---")

    # -----------------------------------------------------
    # BAGIAN GAMBAR
    # -----------------------------------------------------
    st.subheader(L("🖼️ Data Gambar", "🖼️ Figures"))
    col_gambar_cnt, _ = st.columns([1, 2])
    with col_gambar_cnt:
        n_gambar = st.number_input(
            L("Jumlah Gambar dalam artikel", "Number of Figures in article"),
            min_value=0, max_value=20, value=1, step=1, key="n_gambar"
        )

    daftar_gambar = []
    for j in range(int(n_gambar)):
        idx = j + 1
        with st.expander(L(f"🖼️ Gambar {idx}", f"Figure {idx}"), expanded=(j == 0)):
            berkas = st.file_uploader(
                L(f"Unggah Gambar {idx} (≥300 dpi; TIFF/PNG/JPEG)",
                  f"Upload Figure {idx} (≥300 dpi; TIFF/PNG/JPEG)"),
                type=["png", "jpg", "jpeg", "tif", "tiff"],
                key=f"berkas_gambar_{j}"
            )
            cap_g_id = st.text_input(
                L(f"Keterangan Gambar {idx} (Indonesia)", f"Figure {idx} caption (Indonesian — optional)"),
                value="Peta lokasi penelitian di Muara Sungai Tondano." if idx == 1 else "",
                key=f"cap_gambar_id_{j}"
            )
            cap_g_en = st.text_input(
                L(f"Figure {idx} caption (English)", f"Figure {idx} caption (English)"),
                value="Map of the study site at the Tondano River Estuary." if idx == 1 else "",
                key=f"cap_gambar_en_{j}"
            )
            lebar = st.slider(
                L(f"Lebar Gambar {idx} (cm) — maks. 7,6 cm untuk satu kolom",
                  f"Figure {idx} width (cm) — max. 7.6 cm for a single column"),
                4.0, 7.6, 7.6, 0.2,
                key=f"lebar_gambar_{j}"
            )
            daftar_gambar.append({
                "nomor": idx,
                "blob": berkas.getvalue() if berkas else b"",
                "cap_id": cap_g_id,
                "cap_en": cap_g_en,
                "lebar": lebar
            })

with tab4:
    st.caption(L("Tujuh pernyataan akhir wajib, urutan FAS sudah ditetapkan template. "
                 "Bila tidak relevan tulis 'Tidak berlaku / Not applicable'.",
                 "Seven closing statements are mandatory; FAS order is fixed by the template. "
                 "If not applicable, write 'Not applicable'."))
    konflik = st.text_area(
        L("Konflik kepentingan (Competing interests)", "Competing interests"), height=68,
        value=L("Penulis menyatakan tidak ada konflik kepentingan yang relevan dengan "
                "artikel ini.",
                "The authors declare no competing interests relevant to this article."))
    dana = st.text_area(L("Sumber dana (Funding sources)", "Funding sources"), height=68,
                        value=L("Tidak berlaku / Not applicable.", "Not applicable."))
    ucapan = st.text_area(
        L("Ucapan terima kasih (Acknowledgements)", "Acknowledgements"), height=68,
        value=L("Penulis berterima kasih kepada Laboratorium Biologi Laut FPIK UNSRAT.",
                "The authors thank the Marine Biology Laboratory, FPIK UNSRAT."))
    if is_blind:
        st.caption(L("🔒 Otomatis disembunyikan pada versi Blind Review.",
                     "🔒 Automatically hidden in the Blind Review version."))
    kontribusi = st.text_area(
        L("Kontribusi penulis (CRediT)", "Authors’ contributions (CRediT)"), height=68,
        value="FSIM: conceptualization, methodology; ABR: formal analysis; JLT: supervision.")
    data_avail = st.text_area(
        L("Ketersediaan data", "Availability of data and materials"), height=68,
        value=L("Dataset tersedia dari penulis korespondensi berdasarkan permintaan "
                "yang wajar.",
                "The dataset is available from the corresponding author on reasonable request."))
    etik = st.text_area(
        L("Persetujuan etik", "Ethics approval and consent to participate"), height=68,
        value=L("Tidak berlaku / Not applicable.", "Not applicable."))
    st.caption(L("ℹ️ ORCID diisi pada kolom ORCID di tabel penulis (tab Identitas & Judul).",
                 "ℹ️ ORCID is entered in the ORCID column of the author table (Title & Authors)."))
    daftar_pustaka = st.text_area(
        L("Daftar Pustaka (APA 7, satu entri per baris, min. 20)",
          "References (APA 7, one entry per line, min. 20)"),
        height=200,
        value="Rondonuwu, A. B., Kepel, R. C., & Tombokan, J. L. (2025). Diversity and "
              "biogeography of goby. Fisheries and Aquatic Sciences, 28(10), 667–676. "
              "https://doi.org/10.47853/FAS.2025.e56\n"
              "Nei, M., & Kumar, S. (2000). Molecular evolution and phylogenetics. "
              "Oxford University Press.")
    n_ref = len([r for r in daftar_pustaka.split("\n") if r.strip()])
    st.caption(L(f"{n_ref} rujukan" + ("" if n_ref >= 20 else " — minimal 20 disyaratkan."),
                 f"{n_ref} references" + ("" if n_ref >= 20 else " — at least 20 required.")))

# =========================================================
# DAFTAR PERIKSA & UNDUH
# =========================================================
st.markdown("---")
st.subheader(L(f"📥 Unduh Naskah ({versi})", f"📥 Download Manuscript ({versi})"))

peringatan = list(catatan_penulis)
if len(running_title) > 60:
    peringatan.append("Judul singkat melebihi 60 karakter.")
if not 150 <= len(abstract_en.split()) <= 250:
    peringatan.append("Abstract (EN) di luar rentang 150–250 kata.")
if not 150 <= len(abstrak_id.split()) <= 250:
    peringatan.append("Abstrak (ID) di luar rentang 150–250 kata.")
if len(bab5.split()) > 150:
    peringatan.append("Simpulan melebihi 150 kata.")
if n_ref < 20:
    peringatan.append("Daftar pustaka kurang dari 20 rujukan.")
if len(judul_id.split()) > 20 or len(judul_en.split()) > 20:
    peringatan.append("Judul melebihi 20 kata.")

# Validasi koma/titik desimal pada seluruh tabel
for t in daftar_tabel:
    if t["data"].strip():
        _desimal = re.findall(r"\d[.,]\d", t["data"])
        if EN and any("," in d for d in _desimal):
            peringatan.append(f"Tabel {t['nomor']}: Naskah Inggris memakai titik desimal (28.4), "
                              "tetapi isi tabel masih memakai koma.")
        if not EN and any("." in d for d in _desimal):
            peringatan.append(f"Tabel {t['nomor']}: Naskah Indonesia memakai koma desimal (28,4), "
                              "tetapi isi tabel masih memakai titik.")

if peringatan:
    st.warning(L("Daftar periksa pra-submit:", "Pre-submission checklist:")
               + "\n\n" + "\n".join(f"- {w}" for w in peringatan))

sumber_template = template_path or unggahan_template
if not sumber_template:
    st.error("Template resmi belum tersedia. Unggah berkas template pada panel kiri.")
    st.stop()

# Menyiapkan variabel fallback tunggal untuk kompatibilitas mundur jika platax_builder masih versi lama
tbl_1 = daftar_tabel[0] if daftar_tabel else {"cap_id": "", "cap_en": "", "data": "", "catatan": ""}
gbr_1 = daftar_gambar[0] if daftar_gambar else {"cap_id": "", "cap_en": "", "blob": b"", "lebar": 7.6}

naskah = Naskah(
    judul_id=judul_id, judul_en=judul_en, running_title=running_title,
    penulis=penulis_list, afiliasi=daftar_afiliasi, telepon=telepon, email=email,
    abstract_en=abstract_en, keywords_en=keywords_en,
    abstrak_id=abstrak_id, kata_kunci_id=kata_kunci_id,
    bab1=bab1, metode_21=metode_21, metode_22=metode_22,
    metode_221=metode_221, metode_23=metode_23,
    bab3=bab3, bab4=bab4, bab5=bab5,
    # Parameter lama (kompatibilitas mundur)
    cap_tabel_id=tbl_1["cap_id"], cap_tabel_en=tbl_1["cap_en"],
    tabel_data=tbl_1["data"], cat_tabel=tbl_1["catatan"],
    gambar_blob=gbr_1["blob"], lebar_gambar=gbr_1["lebar"],
    cap_gambar_id=gbr_1["cap_id"], cap_gambar_en=gbr_1["cap_en"],
    # Parameter baru (multi tabel & multi gambar)
    tabel_list=daftar_tabel,
    gambar_list=daftar_gambar,
    konflik=konflik, dana=dana, ucapan=ucapan, kontribusi=kontribusi,
    data_avail=data_avail, etik=etik, daftar_pustaka=daftar_pustaka,
    bahasa=BAHASA, blind=is_blind,
)

try:
    berkas = bangun(naskah, sumber_template)
    st.download_button(
        L(f"📄 Unduh Dokumen ({versi})", f"📄 Download Document ({versi})"), data=berkas,
        file_name="Naskah_Blind_Review_PLATAX_2026.docx" if is_blind
        else "Naskah_Final_PLATAX_2026.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        width="stretch")
except Exception as e:  # noqa: BLE001
    st.error(f"Gagal menyusun dokumen: {e}")

st.caption("Setelah dibuka di Word: tekan Ctrl+A lalu F9 untuk memperbarui nomor halaman. "
           "Volume, nomor terbitan, rentang halaman, DOI, dan tanggal terima/revisi/setuju "
           "diisi redaksi pada tahap layout.")
