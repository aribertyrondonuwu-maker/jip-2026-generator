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
st.write("Aplikasi mengisi **template resmi PLATAX 2026** secara langsung, sehingga masthead, "
         "logo, kotak abstrak, margin, header, footer, dan seluruh tipografi mengikuti "
         "template apa adanya.")

# =========================================================
# FORMULIR
# =========================================================
tab1, tab2, tab3, tab5, tab4 = st.tabs(
    ["📌 Identitas & Judul", "📝 Abstract & Abstrak", "📄 Bab 1–5",
     "📊 Tabel & Gambar", "🤝 Pernyataan & Pustaka"])

with tab1:
    c1, c2 = st.columns(2)
    with c1:
        judul_id = st.text_input("Judul Bahasa Indonesia (maks. 20 kata)",
                                 "Keanekaragaman Ikan Gobi di Muara Sungai Tondano, Sulawesi Utara")
    with c2:
        judul_en = st.text_input("Title in English (max. 20 words)",
                                 "Diversity of Goby Fish in Tondano River Estuary, North Sulawesi")
    running_title = st.text_input("Judul singkat / Running title (maks. 60 karakter)",
                                  "Keanekaragaman Ikan Gobi Muara Tondano", max_chars=60)
    st.caption(f"{len(running_title)}/60 karakter")

    st.markdown("---")
    st.subheader("Informasi Penulis")
    if is_blind:
        st.info("ℹ️ Mode Blind Review aktif: data di bawah tetap boleh diisi, tetapi tidak "
                "dimunculkan dalam berkas Word. Gunakan untuk menyusun Title Page terpisah.")

    st.markdown("**1. Daftar penulis** — kolom *No. afiliasi* cukup diisi angka urut "
                "(1, 2, 3 …). Penulis dengan dua afiliasi ditulis `1,2`. Kotak isian "
                "afiliasinya muncul otomatis di bawah.")
    tabel_penulis = st.data_editor(
        [
            {"Nama": "Febry S. I. Menajang", "Afiliasi": "1", "Korespondensi": True,
             "ORCID": "https://orcid.org/0000-0000-0000-0000"},
            {"Nama": "Ari B. Rondonuwu", "Afiliasi": "1", "Korespondensi": False, "ORCID": ""},
            {"Nama": "John L. Tombokan", "Afiliasi": "2", "Korespondensi": False, "ORCID": ""},
            {"Nama": "Veibe Warouw", "Afiliasi": "1,2", "Korespondensi": False, "ORCID": ""},
        ],
        num_rows="dynamic", use_container_width=True, key="tabel_penulis",
        column_config={
            "Nama": st.column_config.TextColumn(
                "Nama penulis", required=True, width="large",
                help="Nama lengkap tanpa gelar. Jangan ketik angka superskrip di sini."),
            "Afiliasi": st.column_config.TextColumn("No. afiliasi", width="small",
                                                    help="Contoh: 1 atau 1,2"),
            "Korespondensi": st.column_config.CheckboxColumn("Koresp.", width="small"),
            "ORCID": st.column_config.TextColumn("ORCID iD", width="medium"),
        })

    # Nomor afiliasi yang benar-benar dirujuk penulis menentukan jumlah kotak isian
    nomor_dirujuk = sorted({
        int(x)
        for baris in as_records(tabel_penulis)
        if str(baris.get("Nama", "") or "").strip()
        for x in re.findall(r"\d+", str(baris.get("Afiliasi") or ""))
    })

    st.markdown("**2. Afiliasi** — satu kotak per nomor yang dirujuk di atas. "
                "Penulis se-institusi cukup memakai nomor yang sama.")
    CONTOH_AFILIASI = [
        "Program Studi Ilmu Kelautan, Fakultas Perikanan dan Ilmu Kelautan, "
        "Universitas Sam Ratulangi, Manado, 95115, Indonesia",
        "Program Studi Manajemen Sumberdaya Perairan, Fakultas Perikanan dan Ilmu Kelautan, "
        "Universitas Sam Ratulangi, Manado, 95115, Indonesia",
    ]
    if not nomor_dirujuk:
        st.caption("Belum ada nomor afiliasi pada tabel penulis.")
        daftar_afiliasi = []
    else:
        if nomor_dirujuk != list(range(1, len(nomor_dirujuk) + 1)):
            st.warning(f"Nomor afiliasi sebaiknya berurutan mulai dari 1. "
                       f"Saat ini terpakai: {nomor_dirujuk}.")
        isian = {}
        for nomor in nomor_dirujuk:
            contoh = CONTOH_AFILIASI[nomor - 1] if nomor <= len(CONTOH_AFILIASI) else ""
            isian[nomor] = st.text_area(
                f"Afiliasi {nomor}  (Departemen, Fakultas, Universitas, Kota, Kode Pos, Negara)",
                value=contoh, height=68, key=f"afiliasi_{nomor}")
        # Daftar dipadatkan menjadi 1..n agar penomoran pada naskah selalu rapat
        peta = {lama: baru for baru, lama in enumerate(nomor_dirujuk, 1)}
        daftar_afiliasi = [isian[lama].strip() for lama in nomor_dirujuk]

    penulis_list, catatan_penulis = susun_penulis(tabel_penulis, daftar_afiliasi,
                                                  peta=peta if nomor_dirujuk else None)
    for c in catatan_penulis:
        st.warning(c)
    if penulis_list:
        st.caption("Pratinjau baris penulis: " + ", ".join(
            p.nama + ("^" + ",".join(map(str, p.afil)) if p.afil else "")
            + ("*" if p.koresp else "") for p in penulis_list))

    c3, c4 = st.columns(2)
    with c3:
        telepon = st.text_input("Telepon korespondensi", "+62-8XXXXXXXXXX")
    with c4:
        email = st.text_input("Surel korespondensi", "nama@unsrat.ac.id")

with tab2:
    st.caption("Urutan template: Abstract (EN) lebih dahulu, baru Abstrak (ID). "
               "150–250 kata per bahasa, satu paragraf naratif, tanpa sitasi.")
    abstract_en = st.text_area("Abstract (English)", height=170,
                               value="This study analysed the diversity of goby fish...")
    keywords_en = st.text_input("Keywords (5, pisahkan dengan koma)",
                                "Goby fish, Diversity, Tondano River, Estuary, North Sulawesi")
    st.caption(f"Abstract: {len(abstract_en.split())} kata")
    abstrak_id = st.text_area("Abstrak (Bahasa Indonesia)", height=170,
                              value="Penelitian ini menganalisis keanekaragaman ikan gobi...")
    kata_kunci_id = st.text_input("Kata kunci (5, pisahkan dengan koma)",
                                  "Ikan gobi, Keanekaragaman, Sungai Tondano, Estuari, Sulawesi Utara")
    st.caption(f"Abstrak: {len(abstrak_id.split())} kata")

with tab3:
    st.caption("Satu baris kosong = paragraf baru. Sitiran seperti (Froese & Pauly, 2024) "
               "atau Rondonuwu et al. (2025) otomatis diwarnai biru 007BB8.")
    bab1 = st.text_area("1. Pendahuluan", height=140,
                        value="Muara Sungai Tondano memiliki peranan ekologis penting "
                              "(Carpenter & Niem, 1998).\n"
                              "Rondonuwu et al. (2025) mencatat kebaruan jenis di kawasan ini.\n"
                              "Penelitian ini bertujuan menilai keanekaragaman ikan gobi.")
    st.markdown("**2. Bahan dan Metode**")
    metode_21 = st.text_area("2.1. Waktu dan lokasi penelitian", height=80,
                             value="Penelitian dilaksanakan Januari–Maret 2026 pada tiga stasiun...")
    metode_22 = st.text_area("2.2. Pengumpulan data", height=80,
                             value="Pengambilan sampel menggunakan jaring insang dengan tiga ulangan...")
    metode_221 = st.text_area("2.2.1. Analisis laboratorium", height=80,
                              value="Identifikasi morfometrik dilakukan di laboratorium...")
    metode_23 = st.text_area("2.3. Analisis data", height=80,
                             value="Indeks Shannon-Wiener dihitung menggunakan PAST v4.03...")
    bab3 = st.text_area("3. Hasil", height=120,
                        value="Tercatat 12 jenis dari tiga stasiun pengamatan (Tabel 1).")
    bab4 = st.text_area("4. Pembahasan", height=140,
                        value="Tingginya keanekaragaman diduga berkaitan dengan variasi salinitas...")
    bab5 = st.text_area("5. Simpulan (maks. 150 kata)", height=90,
                        value="Keanekaragaman ikan gobi tergolong sedang...")
    st.caption(f"Simpulan: {len(bab5.split())} kata")

with tab5:
    st.caption("Keterangan tabel di ATAS tabel, keterangan gambar di BAWAH gambar, "
               "masing-masing wajib disertai terjemahan Inggris pada baris kedua.")
    cap_tabel_id = st.text_input("Keterangan Tabel 1 (Indonesia)",
                                 "Parameter kualitas air pada dua stasiun pengamatan.")
    cap_tabel_en = st.text_input("Table 1 caption (English)",
                                 "Water quality parameters at two observation stations.")
    tabel_data = st.text_area("Isi tabel — kolom dipisah tab atau titik koma; "
                              "baris pertama = kepala tabel", height=110,
                              value="Parameter;Stasiun 1;Stasiun 2\n"
                                    "Suhu (°C);28,4 ± 0,3;29,1 ± 0,5\n"
                                    "Salinitas (‰);32,1 ± 0,4;31,8 ± 0,6")
    cat_tabel = st.text_input("Keterangan kaki tabel",
                              "Keterangan: nilai merupakan rerata ± simpangan baku.")
    st.markdown("---")
    berkas_gambar = st.file_uploader("Unggah Gambar 1 (≥300 dpi; TIFF/PNG/JPEG)",
                                     type=["png", "jpg", "jpeg", "tif", "tiff"])
    cap_gambar_id = st.text_input("Keterangan Gambar 1 (Indonesia)",
                                  "Peta lokasi penelitian di Muara Sungai Tondano.")
    cap_gambar_en = st.text_input("Figure 1 caption (English)",
                                  "Map of the study site at the Tondano River Estuary.")
    lebar_gambar = st.slider("Lebar gambar (cm) — maks. 7,6 cm untuk satu kolom",
                             4.0, 7.6, 7.6, 0.2)

with tab4:
    st.caption("Tujuh pernyataan akhir wajib, urutan FAS sudah ditetapkan template. "
               "Bila tidak relevan tulis 'Tidak berlaku / Not applicable'.")
    konflik = st.text_area("Konflik kepentingan (Competing interests)", height=68,
                           value="Penulis menyatakan tidak ada konflik kepentingan yang "
                                 "relevan dengan artikel ini.")
    dana = st.text_area("Sumber dana (Funding sources)", height=68,
                        value="Tidak berlaku / Not applicable.")
    ucapan = st.text_area("Ucapan terima kasih (Acknowledgements)", height=68,
                          value="Penulis berterima kasih kepada Laboratorium Biologi Laut "
                                "FPIK UNSRAT.")
    if is_blind:
        st.caption("🔒 Otomatis disembunyikan pada versi Blind Review.")
    kontribusi = st.text_area("Kontribusi penulis (CRediT)", height=68,
                              value="FSIM: conceptualization, methodology; "
                                    "ABR: formal analysis; JLT: supervision.")
    data_avail = st.text_area("Ketersediaan data", height=68,
                              value="Dataset tersedia dari penulis korespondensi berdasarkan "
                                    "permintaan yang wajar.")
    etik = st.text_area("Persetujuan etik", height=68, value="Tidak berlaku / Not applicable.")
    st.caption("ℹ️ ORCID diisi pada kolom ORCID di tabel penulis (tab Identitas & Judul).")
    daftar_pustaka = st.text_area("Daftar Pustaka (APA 7, satu entri per baris, min. 20)",
                                  height=200,
                                  value="Rondonuwu, A. B., Kepel, R. C., & Tombokan, J. L. (2025). "
                                        "Diversity and biogeography of goby. Fisheries and Aquatic "
                                        "Sciences, 28(10), 667–676. "
                                        "https://doi.org/10.47853/FAS.2025.e56\n"
                                        "Nei, M., & Kumar, S. (2000). Molecular evolution and "
                                        "phylogenetics. Oxford University Press.")
    n_ref = len([r for r in daftar_pustaka.split("\n") if r.strip()])
    st.caption(f"{n_ref} rujukan" + ("" if n_ref >= 20 else " — minimal 20 disyaratkan."))

# =========================================================
# DAFTAR PERIKSA & UNDUH
# =========================================================
st.markdown("---")
st.subheader(f"📥 Unduh Naskah ({versi})")

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
if peringatan:
    st.warning("Daftar periksa pra-submit:\n\n" + "\n".join(f"- {w}" for w in peringatan))

sumber_template = template_path or unggahan_template
if not sumber_template:
    st.error("Template resmi belum tersedia. Unggah berkas template pada panel kiri.")
    st.stop()

naskah = Naskah(
    judul_id=judul_id, judul_en=judul_en, running_title=running_title,
    penulis=penulis_list, afiliasi=daftar_afiliasi, telepon=telepon, email=email,
    abstract_en=abstract_en, keywords_en=keywords_en,
    abstrak_id=abstrak_id, kata_kunci_id=kata_kunci_id,
    bab1=bab1, metode_21=metode_21, metode_22=metode_22,
    metode_221=metode_221, metode_23=metode_23,
    bab3=bab3, bab4=bab4, bab5=bab5,
    cap_tabel_id=cap_tabel_id, cap_tabel_en=cap_tabel_en,
    tabel_data=tabel_data, cat_tabel=cat_tabel,
    gambar_blob=berkas_gambar.getvalue() if berkas_gambar else b"",
    lebar_gambar=lebar_gambar,
    cap_gambar_id=cap_gambar_id, cap_gambar_en=cap_gambar_en,
    konflik=konflik, dana=dana, ucapan=ucapan, kontribusi=kontribusi,
    data_avail=data_avail, etik=etik, daftar_pustaka=daftar_pustaka,
    blind=is_blind,
)

try:
    berkas = bangun(naskah, sumber_template)
    st.download_button(
        f"📄 Unduh Dokumen ({versi})", data=berkas,
        file_name="Naskah_Blind_Review_PLATAX_2026.docx" if is_blind
        else "Naskah_Final_PLATAX_2026.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True)
except Exception as e:  # noqa: BLE001
    st.error(f"Gagal menyusun dokumen: {e}")

st.caption("Setelah dibuka di Word: tekan Ctrl+A lalu F9 untuk memperbarui nomor halaman. "
           "Volume, nomor terbitan, rentang halaman, DOI, dan tanggal terima/revisi/setuju "
           "diisi redaksi pada tahap layout.")
