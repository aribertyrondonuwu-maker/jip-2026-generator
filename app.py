import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

st.set_page_config(page_title="JIP 2026 Paper Generator", page_icon="📄", layout="wide")

# Fungsi Bantuan Format DOCX
def set_margins(section):
    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(3.2)
    section.left_margin = Cm(1.3)
    section.right_margin = Cm(1.2)

def set_two_columns(section):
    sectPr = section._sectPr
    cols = sectPr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), '2')
        cols[0].set(qn('w:space'), '708')  # ~0.5 cm
    else:
        cols_elem = OxmlElement('w:cols')
        cols_elem.set(qn('w:num'), '2')
        cols_elem.set(qn('w:space'), '708')
        sectPr.append(cols_elem)

def generate_docx(data):
    doc = Document()
    
    # Set Margin Halaman Utama
    sec_1 = doc.sections[0]
    set_margins(sec_1)
    
    # Style Utama: Cambria
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Cambria'
    font.size = Pt(9)

    # 1. Header & Masthead
    p_head = doc.add_paragraph()
    p_head.paragraph_format.space_after = Pt(12)
    run_head = p_head.add_run("Jurnal Ilmiah PLATAX [Vol] ([Tahun]) | e-ISSN 2302-3589\nARTIKEL PENELITIAN / RESEARCH ARTICLE")
    run_head.font.size = Pt(8)
    run_head.font.color.rgb = RGBColor(85, 85, 85)

    # 2. Judul Indonesia
    p_title_id = doc.add_paragraph()
    p_title_id.paragraph_format.space_after = Pt(4)
    r_title_id = p_title_id.add_run(data['judul_id'])
    r_title_id.font.size = Pt(15)
    r_title_id.bold = True
    r_title_id.font.color.rgb = RGBColor(0, 123, 184) # Biru JIP #007BB8

    # 3. Judul Inggris
    p_title_en = doc.add_paragraph()
    p_title_en.paragraph_format.space_after = Pt(8)
    r_title_en = p_title_en.add_run(data['judul_en'])
    r_title_en.font.size = Pt(10)
    r_title_en.bold = True
    r_title_en.italic = True

    # 4. Running Title
    p_run = doc.add_paragraph()
    p_run.paragraph_format.space_after = Pt(12)
    r_run = p_run.add_run(f"Judul singkat (running title): {data['running_title']}")
    r_run.font.size = Pt(9)
    r_run.bold = True

    # 5. Penulis & Afiliasi
    p_author = doc.add_paragraph()
    p_author.paragraph_format.space_after = Pt(4)
    r_author = p_author.add_run(data['penulis_nama'])
    r_author.font.size = Pt(9)
    r_author.bold = True

    p_aff = doc.add_paragraph()
    p_aff.paragraph_format.space_after = Pt(16)
    r_aff = p_aff.add_run(data['penulis_afiliasi'])
    r_aff.font.size = Pt(8)
    r_aff.italic = True

    # 6. Abstract Inggris & Indonesia
    for lang, title, content, kw_label, kw_val in [
        ("en", "Abstract", data['abstract_en'], "Keywords:", data['keywords_en']),
        ("id", "Abstrak", data['abstrak_id'], "Kata kunci:", data['kata_kunci_id'])
    ]:
        p_abs = doc.add_paragraph()
        p_abs.paragraph_format.space_after = Pt(4)
        p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_abs_t = p_abs.add_run(f"{title}\n")
        r_abs_t.bold = True
        r_abs_t.font.size = Pt(8)
        r_abs_c = p_abs.add_run(f"{content}\n\n")
        r_abs_c.font.size = Pt(8)
        r_kw_l = p_abs.add_run(f"{kw_label} ")
        r_kw_l.bold = True
        r_kw_l.font.size = Pt(8)
        r_kw_v = p_abs.add_run(kw_val)
        r_kw_v.font.size = Pt(8)

    # --- BAGIAN 2 KOLOM (BAB 1 - 5) ---
    sec_2 = doc.add_section()
    set_margins(sec_2)
    set_two_columns(sec_2)

    bab_list = [
        ("1. Pendahuluan", data['pendahuluan']),
        ("2. Bahan dan Metode", data['metode']),
        ("3. Hasil", data['hasil']),
        ("4. Pembahasan", data['pembahasan']),
        ("5. Simpulan", data['simpulan'])
    ]

    for title, text in bab_list:
        p_h = doc.add_paragraph()
        p_h.paragraph_format.space_before = Pt(12)
        p_h.paragraph_format.space_after = Pt(4)
        r_h = p_h.add_run(title)
        r_h.bold = True
        r_h.font.size = Pt(8.5)

        p_b = doc.add_paragraph()
        p_b.paragraph_format.space_after = Pt(6)
        p_b.paragraph_format.first_line_indent = Cm(0.5)
        p_b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        r_b = p_b.add_run(text)
        r_b.font.size = Pt(9)

    # --- BAGIAN 1 KOLOM KEMBALI (STATEMENTS & REFS) ---
    sec_3 = doc.add_section()
    set_margins(sec_3)

    statements = [
        ("Konflik kepentingan (Competing interests)", data['konflik']),
        ("Sumber dana (Funding sources)", data['dana']),
        ("Ucapan terima kasih (Acknowledgements)", data['ucapan']),
        ("Kontribusi penulis (Authors’ contributions)", data['kontribusi']),
        ("Ketersediaan data (Availability of data and materials)", data['ketersediaan_data']),
        ("Persetujuan etik (Ethics approval and consent to participate)", data['etik']),
        ("ORCID", data['orcid'])
    ]

    for title, text in statements:
        p_st = doc.add_paragraph()
        p_st.paragraph_format.space_before = Pt(8)
        p_st.paragraph_format.space_after = Pt(2)
        r_st_t = p_st.add_run(title)
        r_st_t.bold = True
        r_st_t.font.size = Pt(8.5)
        
        p_st_b = doc.add_paragraph()
        p_st_b.paragraph_format.space_after = Pt(4)
        r_st_b = p_st_b.add_run(text)
        r_st_b.font.size = Pt(8.5)

    # Daftar Pustaka
    p_ref_h = doc.add_paragraph()
    p_ref_h.paragraph_format.space_before = Pt(14)
    p_ref_h.paragraph_format.space_after = Pt(6)
    r_ref_h = p_ref_h.add_run("Daftar Pustaka")
    r_ref_h.bold = True
    r_ref_h.font.size = Pt(8.5)

    for item in data['daftar_pustaka'].split('\n'):
        if item.strip():
            p_ref = doc.add_paragraph()
            p_ref.paragraph_format.space_after = Pt(4)
            p_ref.paragraph_format.left_indent = Cm(0.7)
            p_ref.paragraph_format.first_line_indent = Cm(-0.7)
            r_ref = p_ref.add_run(item.strip())
            r_ref.font.size = Pt(7.6)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- INTERFACE STREAMLIT ---
st.title("📄 JIP 2026 Article Auto-Generator")
st.write("Isi formulir di bawah ini untuk menghasilkan naskah `.docx` siap kirim sesuai aturan Jurnal Ilmiah PLATAX.")

tab1, tab2, tab3, tab4 = st.tabs(["📌 Identitas & Judul", "📝 Abstract & Abstrak", "📑 Bab 1–5", "🤝 Pernyataan & Pustaka"])

with tab1:
    judul_id = st.text_input("Judul Bahasa Indonesia (Maks 20 Kata)", "Keanekaragaman Ikan Gobi di Muara Sungai Tondano, Sulawesi Utara")
    judul_en = st.text_input("Title in English (Max 20 Words)", "Diversity of Goby Fish in Tondano River Estuary, North Sulawesi")
    running_title = st.text_input("Running Title (Maks 60 Karakter)", "Keanekaragaman Ikan Gobi Sungai Tondano")
    penulis_nama = st.text_input("Nama Penulis (Sesuai Urutan)", "A. B. Rondonuwu1*, R. C. Kepel2, J. L. Tombokan3")
    penulis_afiliasi = st.text_area("Afiliasi Penulis", "1Program Studi Manajemen Sumberdaya Perairan, FPIK, Universitas Sam Ratulangi, Manado, 95115, Indonesia\n2Program Studi Ilmu Kelautan, FPIK, Universitas Sam Ratulangi, Manado, 95115, Indonesia")

with tab2:
    abstract_en = st.text_area("Abstract (English, 150–250 kata)", "This study aims to evaluate the diversity of goby species...", height=150)
    keywords_en = st.text_input("Keywords (3–5 kata)", "Gobiidae, species richness, estuary, conservation, Manado Bay")
    abstrak_id = st.text_area("Abstrak (Indonesia, 150–250 kata)", "Penelitian ini bertujuan untuk mengevaluasi keanekaragaman spesies ikan gobi...", height=150)
    kata_kunci_id = st.text_input("Kata kunci (3–5 kata)", "Gobiidae, kekayaan spesies, muara sungai, konservasi, Teluk Manado")

with tab3:
    pendahuluan = st.text_area("1. Pendahuluan", "Ekosistem muara sungai memiliki peran penting... (Gunakan warna biru pada sitasi secara manual nanti di Word jika perlu)", height=120)
    metode = st.text_area("2. Bahan dan Metode", "Penelitian dilaksanakan pada bulan Maret hingga Mei 2026...", height=120)
    hasil = st.text_area("3. Hasil", "Hasil tangkapan menunjukkan total 12 spesies ikan gobi...", height=120)
    pembahasan = st.text_area("4. Pembahasan", "Tingginya keanekaragaman spesies diduga dipengaruhi oleh...", height=120)
    simpulan = st.text_area("5. Simpulan", "Ekosistem Muara Sungai Tondano memiliki keanekaragaman ikan gobi yang tinggi...", height=100)

with tab4:
    konflik = st.text_input("Konflik Kepentingan", "Penulis menyatakan tidak ada konflik kepentingan yang relevan dengan artikel ini.")
    dana = st.text_input("Sumber Dana", "Tidak berlaku / Not applicable")
    ucapan = st.text_area("Ucapan Terima Kasih", "Penulis mengucapkan terima kasih kepada Teknisi Laboratorium FPIK Unsrat...")
    kontribusi = st.text_input("Kontribusi Penulis (CRediT)", "ABR: conceptualization, methodology; RCK: formal analysis; JLT: supervision.")
    ketersediaan_data = st.text_input("Ketersediaan Data", "Dataset yang mendukung hasil penelitian ini tersedia dari penulis korespondensi berdasarkan permintaan yang wajar.")
    etik = st.text_input("Persetujuan Etik", "Tidak berlaku / Not applicable")
    orcid = st.text_area("ORCID Penulis", "A. B. Rondonuwu: https://orcid.org/0000-0002-1825-0097")
    daftar_pustaka = st.text_area("Daftar Pustaka (APA 7th, Pisahkan dengan Enter)", "Froese, R., & Pauly, D. (2024). FishBase. World Wide Web electronic publication.\nRondonuwu, A. B., Kepel, R. C., & Tombokan, J. L. (2025). Diversity and biogeography of goby. Fisheries and Aquatic Sciences, 28(10), 667–676.", height=150)

st.markdown("---")

data_input = {
    "judul_id": judul_id, "judul_en": judul_en, "running_title": running_title,
    "penulis_nama": penulis_nama, "penulis_afiliasi": penulis_afiliasi,
    "abstract_en": abstract_en, "keywords_en": keywords_en,
    "abstrak_id": abstrak_id, "kata_kunci_id": kata_kunci_id,
    "pendahuluan": pendahuluan, "metode": metode, "hasil": hasil,
    "pembahasan": pembahasan, "simpulan": simpulan,
    "konflik": konflik, "dana": dana, "ucapan": ucapan,
    "kontribusi": kontribusi, "ketersediaan_data": ketersediaan_data,
    "etik": etik, "orcid": orcid, "daftar_pustaka": daftar_pustaka
}

if st.button("🚀 Generate Dokumen Word (.docx)", type="primary", use_container_width=True):
    docx_file = generate_docx(data_input)
    st.success("Dokumen berhasil dibuat sesuai aturan JIP 2026!")
    st.download_button(
        label="📥 Download File .docx",
        data=docx_file,
        file_name="Naskah_JIP_2026.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )