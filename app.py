import io
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
import streamlit as st

# ---------------------------------------------------------
# KONFIGURASI HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="JIP 2026 Article Auto-Generator",
    page_icon="📄",
    layout="wide",
)

st.title("📄 JIP 2026 Article Auto-Generator")
st.write(
    "Isi formulir di bawah ini untuk menghasilkan naskah `.docx` siap kirim"
    " sesuai aturan Jurnal Ilmiah PLATAX."
)

# ---------------------------------------------------------
# SIDEBAR - OPSI VERSI NASKAH
# ---------------------------------------------------------
st.sidebar.header("⚙️ Pengaturan Mode Naskah")
versi_naskah = st.sidebar.radio(
    "Pilih Versi Output Document:",
    ["Final (Lengkap)", "Blind Review"],
    help=(
        "Mode 'Blind Review' akan otomatis menyembunyikan identitas penulis,"
        " afiliasi, email, dan ucapan terima kasih pada file Word yang"
        " diunduh."
    ),
)

# ---------------------------------------------------------
# TAB FORMULIR INPUT
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📌 Identitas & Judul",
        "📝 Abstract & Abstrak",
        "📄 Bab 1–4",
        "🤝 Pernyataan & Pustaka",
    ]
)

with tab1:
    judul_id = st.text_input(
        "Judul Bahasa Indonesia (Maks 20 Kata)",
        value="Keanekaragaman Ikan Gobi di Muara Sungai Tondano, Sulawesi Utara",
    )
    judul_en = st.text_input(
        "Title in English (Max 20 Words)",
        value=(
            "Diversity of Goby Fish in Tondano River Estuary, North Sulawesi"
        ),
    )
    running_title = st.text_input(
        "Running Title (Maks 60 Karakter)",
        value="Keanekaragaman Ikan Gobi di Muara Sungai Tondano",
    )

    st.markdown("---")
    st.subheader("Informasi Penulis")

    if versi_naskah == "Blind Review":
        st.info(
            "ℹ️ **Mode Blind Review Aktif**: Data penulis di bawah ini tetap"
            " dapat diisi, namun **tidak akan dimunculkan** dalam file Word"
            " hasil unduhan."
        )

    nama_penulis = st.text_area(
        "Nama Penulis (Pisahkan dengan koma)",
        value="Ari Berty Rondonuwu*, John Doe, Jane Smith",
    )
    afiliasi = st.text_area(
        "Afiliasi / Instansi Penulis",
        value=(
            "Fakultas Perikanan dan Ilmu Kelautan, Universitas Sam"
            " Ratulangi, Manado"
        ),
    )
    email_korespondensi = st.text_input(
        "Email Korespondensi", value="aribertyrondonuwu@unsrat.ac.id"
    )

with tab2:
    abstrak_id = st.text_area(
        "Abstrak (Bahasa Indonesia)",
        height=150,
        value=(
            "Penelitian ini bertujuan untuk menganalisis keanekaragaman ikan"
            " Gobi di muara Sungai Tondano..."
        ),
    )
    kata_kunci_id = st.text_input(
        "Kata Kunci (Pisahkan dengan koma)",
        value="Ikan Gobi, Keanekaragaman, Sungai Tondano, Estuari",
    )
    abstract_en = st.text_area(
        "Abstract (English)",
        height=150,
        value=(
            "This study aims to analyze the diversity of Goby fish in the"
            " Tondano River estuary..."
        ),
    )
    keywords_en = st.text_input(
        "Keywords (Pisahkan dengan koma)",
        value="Goby fish, Diversity, Tondano River, Estuary",
    )

with tab3:
    bab1 = st.text_area(
        "Bab 1: Pendahuluan",
        height=150,
        value="Muara Sungai Tondano memiliki peranan penting secara ekologis...",
    )
    bab2 = st.text_area(
        "Bab 2: Metode Penelitian",
        height=150,
        value=(
            "Penelitian dilaksanakan pada bulan Januari hingga Maret 2026..."
        ),
    )
    bab3 = st.text_area(
        "Bab 3: Hasil dan Pembahasan",
        height=150,
        value=(
            "Berdasarkan hasil tangkapan di tiga stasiun pengamatan,"
            " ditemukan..."
        ),
    )
    bab4 = st.text_area(
        "Bab 4: Kesimpulan dan Saran",
        height=150,
        value="Keanekaragaman ikan Gobi di muara Sungai Tondano tergolong...",
    )

with tab4:
    ucapan_terima_kasih = st.text_area(
        "Ucapan Terima Kasih (Acknowledgements)",
        height=100,
        value=(
            "Penulis mengucapkan terima kasih kepada Laboratorium Biologi"
            " Laut..."
        ),
    )
    if versi_naskah == "Blind Review":
        st.caption(
            "🔒 *Catatan: Bagian Ucapan Terima Kasih akan otomatis"
            " disembunyikan/dianonimkan pada versi Blind Review.*"
        )

    daftar_pustaka = st.text_area(
        "Daftar Pustaka (APA Style)",
        height=200,
        value=(
            "Rondonuwu, A. B. (2026). Ikhtiofauna Perairan Tawar dan Payau."
            " Jurnal Ilmiah PLATAX, 14(1), 10-20."
        ),
    )


# ---------------------------------------------------------
# FUNGSI PENYUSUN DOKUMEN WORD (.DOCX)
# ---------------------------------------------------------
def create_docx():
    doc = Document()

    # Judul Bahasa Indonesia
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(judul_id.upper())
    run_title.bold = True
    run_title.font.size = Pt(14)

    # Judul Bahasa Inggris
    p_title_en = doc.add_paragraph()
    p_title_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title_en = p_title_en.add_run(f"({judul_en})")
    run_title_en.italic = True
    run_title_en.font.size = Pt(11)

    # Running Title
    p_rt = doc.add_paragraph()
    p_rt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rt = p_rt.add_run(f"Running Title: {running_title}")
    r_rt.font.size = Pt(9)
    r_rt.font.color.rgb = RGBColor(128, 128, 128)

    doc.add_paragraph()  # Spasi

    # LOGIKA BLIND REVIEW VS FINAL (PENULIS)
    if versi_naskah == "Final (Lengkap)":
        p_auth = doc.add_paragraph()
        p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_auth = p_auth.add_run(nama_penulis)
        r_auth.bold = True

        p_aff = doc.add_paragraph()
        p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_aff = p_aff.add_run(f"{afiliasi}\nEmail: {email_korespondensi}")
        r_aff.font.size = Pt(9)
        r_aff.font.italic = True
    else:
        p_anon = doc.add_paragraph()
        p_anon.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_anon = p_anon.add_run(
            "[INFORMASI PENULIS DIKOSONGKAN UNTUK PROSES BLIND REVIEW]"
        )
        r_anon.font.italic = True
        r_anon.font.color.rgb = RGBColor(180, 0, 0)

    doc.add_paragraph()  # Spasi

    # Abstrak Bahasa Indonesia
    doc.add_heading("ABSTRAK", level=2)
    doc.add_paragraph(abstrak_id)
    p_kw = doc.add_paragraph()
    r_kw_title = p_kw.add_run("Kata Kunci: ")
    r_kw_title.bold = True
    p_kw.add_run(kata_kunci_id)

    doc.add_paragraph()

    # Abstract Bahasa Inggris
    doc.add_heading("ABSTRACT", level=2)
    doc.add_paragraph(abstract_en)
    p_kw_en = doc.add_paragraph()
    r_kw_title_en = p_kw_en.add_run("Keywords: ")
    r_kw_title_en.bold = True
    p_kw_en.add_run(keywords_en)

    doc.add_paragraph()

    # Isi Bab Utama
    doc.add_heading("1. PENDAHULUAN", level=1)
    doc.add_paragraph(bab1)

    doc.add_heading("2. METODE PENELITIAN", level=1)
    doc.add_paragraph(bab2)

    doc.add_heading("3. HASIL DAN PEMBAHASAN", level=1)
    doc.add_paragraph(bab3)

    doc.add_heading("4. KESIMPULAN DAN SARAN", level=1)
    doc.add_paragraph(bab4)

    # LOGIKA BLIND REVIEW (UCAPAN TERIMA KASIH)
    doc.add_heading("UCAPAN TERIMA KASIH", level=1)
    if versi_naskah == "Final (Lengkap)":
        doc.add_paragraph(ucapan_terima_kasih)
    else:
        p_thx = doc.add_paragraph(
            "[Ucapan terima kasih disembunyikan untuk proses Blind Review]"
        )
        p_thx.runs[0].font.italic = True

    # Daftar Pustaka
    doc.add_heading("DAFTAR PUSTAKA", level=1)
    doc.add_paragraph(daftar_pustaka)

    # Simpan ke memory buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# TOMBOL GENERATE & DOWNLOAD
# ---------------------------------------------------------
st.markdown("---")
st.subheader(f"📥 Unduh Naskah ({versi_naskah})")

filename = (
    "Naskah_Blind_Review.docx"
    if versi_naskah == "Blind Review"
    else "Naskah_Final_Lengkap.docx"
)

st.download_button(
    label=f"📄 Download Document ({versi_naskah})",
    data=create_docx(),
    file_name=filename,
    mime=(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ),
    use_container_width=True,
)
