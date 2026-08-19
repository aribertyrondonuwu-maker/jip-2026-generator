from dataclasses import dataclass, field
import io
import re
from typing import List, Dict, Any
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


@dataclass
class Naskah:
    judul_id: str = ""
    judul_en: str = ""
    penulis_list: List[Dict[str, Any]] = field(default_factory=list)
    afiliasi_list: List[str] = field(default_factory=list)
    email_korespondensi: str = ""
    abstrak_id: str = ""
    abstrak_en: str = ""
    kata_kunci_id: str = ""
    kata_kunci_en: str = ""
    pendahuluan: str = ""
    metode: str = ""
    hasil_pembahasan: str = ""
    kesimpulan: str = ""
    ucapan_terima_kasih: str = ""
    daftar_pustaka: str = ""
    bahasa: str = "id"
    tabel_list: List[Dict[str, Any]] = field(default_factory=list)
    gambar_list: List[Dict[str, Any]] = field(default_factory=list)


def build_docx(naskah: Naskah) -> io.BytesIO:
    doc = Document()

    # --- Pengaturan Halaman (A4 & Margin 2.5 cm) ---
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # Styling Font Dasar
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)

    def format_paragraph(p, space_before=0, space_after=4, line_spacing=1.15):
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing

    # --- Header Jurnal ---
    p_hdr = doc.add_paragraph()
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_hdr = p_hdr.add_run("Jurnal Ilmiah PLATAX")
    r_hdr.font.size = Pt(9)
    r_hdr.font.italic = True
    r_hdr.font.color.rgb = RGBColor(128, 128, 128)

    # --- Judul Naskah ---
    if naskah.judul_id:
        p_title_id = doc.add_paragraph()
        p_title_id.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_title_id, space_before=12, space_after=6)
        r = p_title_id.add_run(naskah.judul_id.upper())
        r.bold = True
        r.font.size = Pt(14)

    if naskah.judul_en:
        p_title_en = doc.add_paragraph()
        p_title_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_title_en, space_before=0, space_after=12)
        r = p_title_en.add_run(naskah.judul_en)
        r.italic = True
        r.bold = True
        r.font.size = Pt(12)

    # --- Penulis & Afiliasi ---
    if naskah.penulis_list:
        p_auth = doc.add_paragraph()
        p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_auth, space_before=6, space_after=4)
        
        for idx, p in enumerate(naskah.penulis_list):
            nama = p.get("nama", "")
            aff_ids = p.get("afiliasi_ids", [1])
            is_corresp = p.get("is_corresp", False)
            
            aff_str = ",".join(str(a) for a in aff_ids)
            if is_corresp:
                aff_str += "*"
                
            p_auth.add_run(nama).bold = True
            r_aff = p_auth.add_run(f"({aff_str})")
            r_aff.font.superscript = True
            
            if idx < len(naskah.penulis_list) - 1:
                p_auth.add_run(", ")

    if naskah.afiliasi_list:
        for idx, aff in enumerate(naskah.afiliasi_list, start=1):
            p_aff = doc.add_paragraph()
            p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p_aff, space_before=0, space_after=2)
            
            r_idx = p_aff.add_run(f"{idx} ")
            r_idx.font.superscript = True
            r_text = p_aff.add_run(aff)
            r_text.font.size = Pt(9.5)
            r_text.font.italic = True

    if naskah.email_korespondensi:
        p_email = doc.add_paragraph()
        p_email.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_email, space_before=2, space_after=16)
        r_email = p_email.add_run(f"*Email Korespondensi: {naskah.email_korespondensi}")
        r_email.font.size = Pt(9)
        r_email.font.italic = True

    # --- Abstrak Bahasa Indonesia ---
    if naskah.abstrak_id:
        p_abs_head = doc.add_paragraph()
        format_paragraph(p_abs_head, space_before=6, space_after=2)
        p_abs_head.add_run("ABSTRAK").bold = True

        p_abs = doc.add_paragraph(naskah.abstrak_id)
        p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_abs, space_before=0, space_after=4)
        p_abs.runs[0].font.size = Pt(10)

        if naskah.kata_kunci_id:
            p_kw = doc.add_paragraph()
            format_paragraph(p_kw, space_before=0, space_after=12)
            p_kw.add_run("Kata kunci: ").bold = True
            p_kw.add_run(naskah.kata_kunci_id).font.size = Pt(10)

    # --- Abstract Bahasa Inggris ---
    if naskah.abstrak_en:
        p_abs_head_en = doc.add_paragraph()
        format_paragraph(p_abs_head_en, space_before=6, space_after=2)
        p_abs_head_en.add_run("ABSTRACT").bold = True

        p_abs_en = doc.add_paragraph(naskah.abstrak_en)
        p_abs_en.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_abs_en, space_before=0, space_after=4)
        p_abs_en.runs[0].font.size = Pt(10)
        p_abs_en.runs[0].font.italic = True

        if naskah.kata_kunci_en:
            p_kw_en = doc.add_paragraph()
            format_paragraph(p_kw_en, space_before=0, space_after=16)
            r_lbl = p_kw_en.add_run("Keywords: ")
            r_lbl.bold = True
            r_lbl.font.italic = True
            r_lbl.font.size = Pt(10)
            r_val = p_kw_en.add_run(naskah.kata_kunci_en)
            r_val.font.size = Pt(10)
            r_val.font.italic = True

    def add_section_heading(title: str):
        p = doc.add_paragraph()
        format_paragraph(p, space_before=12, space_after=4)
        r = p.add_run(title.upper())
        r.bold = True
        r.font.size = Pt(11)

    # --- Bagian Utama ---
    sections = [
        ("PENDAHULUAN", naskah.pendahuluan),
        ("METODE PENELITIAN", naskah.metode),
        ("HASIL DAN PEMBAHASAN", naskah.hasil_pembahasan),
    ]

    for sec_title, sec_content in sections:
        if sec_content and sec_content.strip():
            add_section_heading(sec_title)
            p = doc.add_paragraph(sec_content)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            format_paragraph(p, space_before=0, space_after=6)

    # --- Rendering Multi-Tabel ---
    if naskah.tabel_list:
        for tbl in naskah.tabel_list:
            if not tbl.get("data", "").strip():
                continue

            p_cap = doc.add_paragraph()
            format_paragraph(p_cap, space_before=8, space_after=2)
            
            num = tbl.get("nomor", 1)
            cap_id = tbl.get("cap_id", "")
            cap_en = tbl.get("cap_en", "")
            
            p_cap.add_run(f"Tabel {num}. {cap_id}").bold = True
            if naskah.bahasa == "id" and cap_en:
                p_cap.add_run(f"\nTable {num}. {cap_en}").italic = True

            baris_raw = [b.strip() for b in tbl["data"].strip().split("\n") if b.strip()]
            if baris_raw:
                materi = [re.split(r"[\t;]", b) for b in baris_raw]
                n_rows = len(materi)
                n_cols = max(len(r) for r in materi)

                table = doc.add_table(rows=n_rows, cols=n_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.style = 'Table Grid'

                for r_idx, row_data in enumerate(materi):
                    for c_idx, val in enumerate(row_data):
                        if c_idx < n_cols:
                            cell = table.cell(r_idx, c_idx)
                            cell.text = val.strip()
                            p_cell = cell.paragraphs[0]
                            p_cell.paragraph_format.space_before = Pt(2)
                            p_cell.paragraph_format.space_after = Pt(2)
                            if r_idx == 0 and p_cell.runs:
                                p_cell.runs[0].bold = True

            if tbl.get("catatan"):
                p_cat = doc.add_paragraph(tbl["catatan"])
                format_paragraph(p_cat, space_before=2, space_after=8)
                if p_cat.runs:
                    p_cat.runs[0].font.size = Pt(9)

    # --- Rendering Multi-Gambar ---
    if naskah.gambar_list:
        for gbr in naskah.gambar_list:
            if not gbr.get("blob"):
                continue

            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p_img, space_before=8, space_after=4)

            stream = io.BytesIO(gbr["blob"])
            width_cm = gbr.get("lebar", 12.0)
            p_img.add_run().add_picture(stream, width=Cm(width_cm))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p_cap, space_before=2, space_after=8)

            num = gbr.get("nomor", 1)
            cap_id = gbr.get("cap_id", "")
            cap_en = gbr.get("cap_en", "")

            p_cap.add_run(f"Gambar {num}. {cap_id}").bold = True
            if naskah.bahasa == "id" and cap_en:
                p_cap.add_run(f"\nFigure {num}. {cap_en}").italic = True

    # --- Kesimpulan & Ucapan Terima Kasih ---
    if naskah.kesimpulan and naskah.kesimpulan.strip():
        add_section_heading("KESIMPULAN")
        p_kes = doc.add_paragraph(naskah.kesimpulan)
        p_kes.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_kes, space_before=0, space_after=6)

    if naskah.ucapan_terima_kasih and naskah.ucapan_terima_kasih.strip():
        add_section_heading("UCAPAN TERIMA KASIH")
        p_utk = doc.add_paragraph(naskah.ucapan_terima_kasih)
        p_utk.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_utk, space_before=0, space_after=6)

    # --- Daftar Pustaka ---
    if naskah.daftar_pustaka and naskah.daftar_pustaka.strip():
        add_section_heading("DAFTAR PUSTAKA")
        for line in naskah.daftar_pustaka.strip().split("\n"):
            if line.strip():
                p_dp = doc.add_paragraph(line.strip())
                p_dp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                format_paragraph(p_dp, space_before=0, space_after=4)
                p_dp.paragraph_format.left_indent = Cm(0.63)
                p_dp.paragraph_format.first_line_indent = Cm(-0.63)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output
