import io
import re
from typing import List, Dict, Any, Union
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT


class Penulis:
    def __init__(self, nama: str = "", afiliasi_ids: List[int] = None, is_corresp: bool = False, email: str = "", orcid: str = "", **kwargs):
        self.nama = nama
        self.afiliasi_ids = afiliasi_ids if afiliasi_ids is not None else []
        self.is_corresp = is_corresp
        self.email = email
        self.orcid = orcid
        for key, value in kwargs.items():
            setattr(self, key, value)


class Naskah:
    def __init__(
        self,
        judul_id: str = "",
        judul_en: str = "",
        running_title: str = "",
        penulis_list: List[Union[Penulis, Dict[str, Any]]] = None,
        afiliasi_list: List[str] = None,
        email_korespondensi: str = "",
        abstrak_id: str = "",
        abstrak_en: str = "",
        kata_kunci_id: str = "",
        kata_kunci_en: str = "",
        pendahuluan: str = "",
        metode: str = "",
        hasil_pembahasan: str = "",
        kesimpulan: str = "",
        ucapan_terima_kasih: str = "",
        daftar_pustaka: str = "",
        bahasa: str = "id",
        blind: bool = False,
        tabel_list: List[Dict[str, Any]] = None,
        gambar_list: List[Dict[str, Any]] = None,
        tbl_1: Dict[str, Any] = None,
        gbr_1: Dict[str, Any] = None,
        **kwargs
    ):
        self.judul_id = judul_id
        self.judul_en = judul_en
        self.running_title = running_title
        self.penulis_list = penulis_list if penulis_list is not None else []
        self.afiliasi_list = afiliasi_list if afiliasi_list is not None else []
        self.email_korespondensi = email_korespondensi
        self.abstrak_id = abstrak_id
        self.abstrak_en = abstrak_en
        self.kata_kunci_id = kata_kunci_id
        self.kata_kunci_en = kata_kunci_en
        self.pendahuluan = pendahuluan
        self.metode = metode
        self.hasil_pembahasan = hasil_pembahasan
        self.kesimpulan = kesimpulan
        self.ucapan_terima_kasih = ucapan_terima_kasih
        self.daftar_pustaka = daftar_pustaka
        self.bahasa = bahasa
        self.blind = blind
        self.tabel_list = tabel_list if tabel_list is not None else []
        self.gambar_list = gambar_list if gambar_list is not None else []
        self.tbl_1 = tbl_1 if tbl_1 is not None else {}
        self.gbr_1 = gbr_1 if gbr_1 is not None else {}

        # Menampung argumen opsional/tambahan lainnya
        for key, value in kwargs.items():
            setattr(self, key, value)


def bangun(naskah: Naskah) -> io.BytesIO:
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
    if getattr(naskah, "judul_id", ""):
        p_title_id = doc.add_paragraph()
        p_title_id.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_title_id, space_before=12, space_after=6)
        r = p_title_id.add_run(naskah.judul_id.upper())
        r.bold = True
        r.font.size = Pt(14)

    if getattr(naskah, "judul_en", ""):
        p_title_en = doc.add_paragraph()
        p_title_en.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p_title_en, space_before=0, space_after=12)
        r = p_title_en.add_run(naskah.judul_en)
        r.italic = True
        r.bold = True
        r.font.size = Pt(12)

    # Jika peninjauan blind (is_blind=True), sembunyikan identitas penulis
    if not getattr(naskah, "blind", False):
        # --- Penulis & Afiliasi ---
        if getattr(naskah, "penulis_list", None):
            p_auth = doc.add_paragraph()
            p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p_auth, space_before=6, space_after=4)
            
            for idx, p in enumerate(naskah.penulis_list):
                if isinstance(p, Penulis):
                    nama = getattr(p, "nama", "")
                    aff_ids = getattr(p, "afiliasi_ids", [1])
                    is_corresp = getattr(p, "is_corresp", False)
                elif isinstance(p, dict):
                    nama = p.get("nama", "")
                    aff_ids = p.get("afiliasi_ids", [1])
                    is_corresp = p.get("is_corresp", False)
                else:
                    continue
                
                aff_str = ",".join(str(a) for a in aff_ids)
                if is_corresp:
                    aff_str += "*"
                    
                p_auth.add_run(nama).bold = True
                r_aff = p_auth.add_run(f"({aff_str})")
                r_aff.font.superscript = True
                
                if idx < len(naskah.penulis_list) - 1:
                    p_auth.add_run(", ")

        if getattr(naskah, "afiliasi_list", None):
            for idx, aff in enumerate(naskah.afiliasi_list, start=1):
                p_aff = doc.add_paragraph()
                p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
                format_paragraph(p_aff, space_before=0, space_after=2)
                
                r_idx = p_aff.add_run(f"{idx} ")
                r_idx.font.superscript = True
                r_text = p_aff.add_run(aff)
                r_text.font.size = Pt(9.5)
                r_text.font.italic = True

        if getattr(naskah, "email_korespondensi", ""):
            p_email = doc.add_paragraph()
            p_email.alignment = WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p_email, space_before=2, space_after=16)
            r_email = p_email.add_run(f"*Email Korespondensi: {naskah.email_korespondensi}")
            r_email.font.size = Pt(9)
            r_email.font.italic = True

    # --- Abstrak Bahasa Indonesia ---
    if getattr(naskah, "abstrak_id", ""):
        p_abs_head = doc.add_paragraph()
        format_paragraph(p_abs_head, space_before=6, space_after=2)
        p_abs_head.add_run("ABSTRAK").bold = True

        p_abs = doc.add_paragraph(naskah.abstrak_id)
        p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_abs, space_before=0, space_after=4)
        p_abs.runs[0].font.size = Pt(10)

        if getattr(naskah, "kata_kunci_id", ""):
            p_kw = doc.add_paragraph()
            format_paragraph(p_kw, space_before=0, space_after=12)
            p_kw.add_run("Kata kunci: ").bold = True
            p_kw.add_run(naskah.kata_kunci_id).font.size = Pt(10)

    # --- Abstract Bahasa Inggris ---
    if getattr(naskah, "abstrak_en", ""):
        p_abs_head_en = doc.add_paragraph()
        format_paragraph(p_abs_head_en, space_before=6, space_after=2)
        p_abs_head_en.add_run("ABSTRACT").bold = True

        p_abs_en = doc.add_paragraph(naskah.abstrak_en)
        p_abs_en.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_abs_en, space_before=0, space_after=4)
        p_abs_en.runs[0].font.size = Pt(10)
        p_abs_en.runs[0].font.italic = True

        if getattr(naskah, "kata_kunci_en", ""):
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
        ("PENDAHULUAN", getattr(naskah, "pendahuluan", "")),
        ("METODE PENELITIAN", getattr(naskah, "metode", "")),
        ("HASIL DAN PEMBAHASAN", getattr(naskah, "hasil_pembahasan", "")),
    ]

    for sec_title, sec_content in sections:
        if sec_content and sec_content.strip():
            add_section_heading(sec_title)
            p = doc.add_paragraph(sec_content)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            format_paragraph(p, space_before=0, space_after=6)

    # --- Rendering Multi-Tabel ---
    tabel_items = getattr(naskah, "tabel_list", []) or ([getattr(naskah, "tbl_1", {})] if getattr(naskah, "tbl_1", None) else [])
    for tbl in tabel_items:
        if not tbl or not tbl.get("data", "").strip():
            continue

        p_cap = doc.add_paragraph()
        format_paragraph(p_cap, space_before=8, space_after=2)
        
        num = tbl.get("nomor", 1)
        cap_id = tbl.get("cap_id", "")
        cap_en = tbl.get("cap_en", "")
        
        p_cap.add_run(f"Tabel {num}. {cap_id}").bold = True
        if getattr(naskah, "bahasa", "id") == "id" and cap_en:
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
    gambar_items = getattr(naskah, "gambar_list", []) or ([getattr(naskah, "gbr_1", {})] if getattr(naskah, "gbr_1", None) else [])
    for gbr in gambar_items:
        if not gbr or not gbr.get("blob"):
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
        if getattr(naskah, "bahasa", "id") == "id" and cap_en:
            p_cap.add_run(f"\nFigure {num}. {cap_en}").italic = True

    # --- Kesimpulan & Ucapan Terima Kasih ---
    if getattr(naskah, "kesimpulan", "") and naskah.kesimpulan.strip():
        add_section_heading("KESIMPULAN")
        p_kes = doc.add_paragraph(naskah.kesimpulan)
        p_kes.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_kes, space_before=0, space_after=6)

    if getattr(naskah, "ucapan_terima_kasih", "") and naskah.ucapan_terima_kasih.strip():
        add_section_heading("UCAPAN TERIMA KASIH")
        p_utk = doc.add_paragraph(naskah.ucapan_terima_kasih)
        p_utk.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_paragraph(p_utk, space_before=0, space_after=6)

    # --- Daftar Pustaka ---
    if getattr(naskah, "daftar_pustaka", "") and naskah.daftar_pustaka.strip():
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


# Alias kompatibilitas
build_docx = bangun
