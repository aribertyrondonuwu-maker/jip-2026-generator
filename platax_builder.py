"""
platax_builder.py — Versi Lengkap & Diperbaiki
Generator dokumen Word untuk Jurnal Ilmiah PLATAX 2026
"""
import io
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree

try:
    from latex2mathml.converter import convert as latex_to_mathml
    HAS_LATEX = True
except:
    HAS_LATEX = False

# =============================================================================
# KONSTANTA
# =============================================================================
FONT_NAME = 'Cambria'
BLUE = RGBColor(0x00, 0x7B, 0xB8)
BLACK = RGBColor(0x00, 0x00, 0x00)
LINK_BLUE = RGBColor(0x21, 0x96, 0xD1)


# =============================================================================
# CLASS PENULIS & NASKAH
# =============================================================================
class Penulis:
    def __init__(self, nama="", afiliasi_ids=None, is_corresp=False, email="", orcid="", **kw):
        self.nama = nama
        self.afiliasi_ids = afiliasi_ids or []
        self.is_corresp = is_corresp
        self.email = email
        self.orcid = orcid


class Naskah:
    def __init__(self, **kwargs):
        for k, v in kwargs.items():
            setattr(self, k, v)
        self.penulis_list = kwargs.get('penulis_list', [])
        self.afiliasi_list = kwargs.get('afiliasi_list', [])
        self.tabel_list = kwargs.get('tabel_list', [])
        self.gambar_list = kwargs.get('gambar_list', [])
        self.equations = {k: v for k, v in kwargs.items() if k.startswith('eq_')}


# =============================================================================
# HEADING LOOKUP — sesuai bahasa template
# =============================================================================
HEADINGS = {
    "id": {
        "bab1":    "1. Pendahuluan",
        "met21":   "2.1. Waktu dan lokasi penelitian",
        "met22":   "2.2. Pengumpulan data",
        "met221":  "2.2.1. Analisis laboratorium",
        "met23":   "2.3. Analisis data",
        "bab3":    "3. Hasil",
        "bab4":    "4. Pembahasan",
        "bab5":    "5. Simpulan",
        # Pernyataan akhir
        "konflik":    "Konflik kepentingan (Competing interests)",
        "dana":       "Sumber dana (Funding sources)",
        "ucapan":     "Ucapan terima kasih (Acknowledgements)",
        "kontribusi": "Kontribusi penulis (Authors\u2019 contributions)",
        "data_avail": "Ketersediaan data (Availability of data and materials)",
        "etik":       "Persetujuan etik (Ethics approval and consent to participate)",
        # Judul & placeholder
        "judul_id":    "[Judul penelitian dalam bahasa Indonesia:",
        "judul_en":    "[Complete Research Title in English:",
        "running":     "[Judul singkat, maks. 60 karakter]",
        "penulis":     "[Nama Penulis 1]",
        "afiliasi":    "[Departemen, Fakultas",
        "abs_en":      "[Tulis abstract bahasa Inggris di sini.",
        "kw_en":       "Keywords: [keyword 1",
        "abs_id":      "[Tulis abstrak bahasa Indonesia di sini",
        "kw_id":       "Kata kunci: [kata kunci 1",
        "footnote":    "*Penulis korespondensi (Corresponding author):",
        "refs":        "[APA 7th Edition",
        "tabel_label": "Tabel",   # prefix label tabel
        "gambar_label": "Gambar", # prefix label gambar
    },
    "en": {
        "bab1":    "1. Introduction",
        "met21":   "2.1. Study period and location",
        "met22":   "2.2. Data collection",
        "met221":  "2.2.1. Laboratory analysis",
        "met23":   "2.3. Data analysis",
        "bab3":    "3. Results",
        "bab4":    "4. Discussion",
        "bab5":    "5. Conclusion",
        # Pernyataan akhir
        "konflik":    "Competing interests",
        "dana":       "Funding sources",
        "ucapan":     "Acknowledgements",
        "kontribusi": "Authors\u2019 contributions",
        "data_avail": "Availability of data and materials",
        "etik":       "Ethics approval and consent to participate",
        # Judul & placeholder
        "judul_id":    "[Research Title in Indonesian:",
        "judul_en":    "[Complete Research Title in English:",
        "running":     "[Short title, max. 60 characters]",
        "penulis":     "[Author Name 1]",
        "afiliasi":    "[Department, Faculty",
        "abs_en":      "[Write the English abstract here.",
        "kw_en":       "Keywords: [keyword 1",
        "abs_id":      "[Write the Indonesian abstract here",
        "kw_id":       "Kata kunci: [kata kunci 1",
        "footnote":    "*Corresponding author:",
        "refs":        "[APA 7th Edition",
        "tabel_label": "Table",   # prefix label tabel
        "gambar_label": "Figure", # prefix label gambar
    }
}


# =============================================================================
# FUNGSI XML UNTUK SET FONT & FORMATTING
# =============================================================================
def set_run_xml(run, size=None, bold=None, italic=None, color=None, name=FONT_NAME):
    rPr = run._r.get_or_add_rPr()
    if name:
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}/>')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), name)
        rFonts.set(qn('w:hAnsi'), name)
        rFonts.set(qn('w:eastAsia'), name)
        rFonts.set(qn('w:cs'), name)
    if size is not None:
        sz_val = str(int(size * 2))
        sz = rPr.find(qn('w:sz'))
        if sz is None:
            sz = parse_xml(f'<w:sz {nsdecls("w")}/>')
            rPr.append(sz)
        sz.set(qn('w:val'), sz_val)
        szCs = rPr.find(qn('w:szCs'))
        if szCs is None:
            szCs = parse_xml(f'<w:szCs {nsdecls("w")}/>')
            rPr.append(szCs)
        szCs.set(qn('w:val'), sz_val)
    if bold is not None:
        b_elem = rPr.find(qn('w:b'))
        if bold:
            if b_elem is None:
                b_elem = parse_xml(f'<w:b {nsdecls("w")}/>')
                rPr.append(b_elem)
        else:
            if b_elem is not None:
                rPr.remove(b_elem)
    if italic is not None:
        i_elem = rPr.find(qn('w:i'))
        if italic:
            if i_elem is None:
                i_elem = parse_xml(f'<w:i {nsdecls("w")}/>')
                rPr.append(i_elem)
        else:
            if i_elem is not None:
                rPr.remove(i_elem)
    if color is not None:
        color_elem = rPr.find(qn('w:color'))
        if color_elem is None:
            color_elem = parse_xml(f'<w:color {nsdecls("w")}/>')
            rPr.append(color_elem)
        color_elem.set(qn('w:val'), f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')


# =============================================================================
# FUNGSI BANTU PARAGRAF
# =============================================================================
def clear_paragraph_runs(para):
    runs = list(para.runs)
    if len(runs) > 1:
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    if runs:
        runs[0].text = ""
    return para


def add_formatted_text(para, text, size=9, bold=False, italic=False,
                       color=BLACK, alignment=None, name=FONT_NAME):
    if alignment is not None:
        para.alignment = alignment
    run = para.add_run(text)
    set_run_xml(run, size=size, bold=bold, italic=italic, color=color, name=name)
    return run


def fill_paragraph(para, text, size=9, bold=False, italic=False,
                   color=BLACK, alignment=None, name=FONT_NAME):
    clear_paragraph_runs(para)
    if alignment is not None:
        para.alignment = alignment
    if para.runs:
        run = para.runs[0]
    else:
        run = para.add_run("")
    run.text = text
    set_run_xml(run, size=size, bold=bold, italic=italic, color=color, name=name)
    return run


def find_paragraph(doc, keywords):
    for para in doc.paragraphs:
        for kw in keywords:
            if kw in para.text:
                return para
    return None


# =============================================================================
# TABEL DENGAN BORDER HORIZONTAL SAJA (OPEN TABLE)
# =============================================================================
def set_table_open_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders_xml = (
        f'<w:tblBorders {nsdecls("w")}>'
        '  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '</w:tblBorders>'
    )
    borders = parse_xml(borders_xml)
    old_borders = tblPr.find(qn('w:tblBorders'))
    if old_borders is not None:
        tblPr.remove(old_borders)
    tblPr.append(borders)


def create_table_in_doc(doc, headers, rows, nomor, caption, footnote="", bahasa="id"):
    """
    Buat tabel baru di dokumen.
    caption: satu teks caption sesuai bahasa naskah (tidak bilingual).
    """
    H = HEADINGS[bahasa]
    label_prefix = H["tabel_label"]   # "Tabel" atau "Table"

    # Caption (satu bahasa sesuai naskah)
    cap_para = doc.add_paragraph()
    cap_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_cap = cap_para.add_run(f"{label_prefix} {nomor}. {caption}")
    set_run_xml(run_cap, size=7.6, bold=True, name=FONT_NAME)

    # Buat tabel
    num_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(header)
        set_run_xml(run, size=9, bold=True, name=FONT_NAME)

    # Data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            if col_idx < num_cols:
                cell = table.rows[row_idx + 1].cells[col_idx]
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(cell_text))
                set_run_xml(run, size=9, name=FONT_NAME)

    set_table_open_borders(table)

    # Catatan kaki
    if footnote:
        keterangan_label = "Keterangan" if bahasa == "id" else "Note"
        fn_para = doc.add_paragraph()
        run_fn = fn_para.add_run(f"{keterangan_label}: {footnote}")
        set_run_xml(run_fn, size=8, italic=True, name=FONT_NAME)

    return table


# =============================================================================
# FUNGSI UTAMA: BANGUN DOKUMEN
# =============================================================================
def bangun(naskah=None, template_path=None, *args, **kwargs):
    if naskah is None and args:
        naskah = args[0]

    is_blind = getattr(naskah, 'blind', False)
    bahasa = getattr(naskah, 'bahasa', 'id')
    is_en = bahasa == 'en'
    H = HEADINGS[bahasa]

    # Load template — wajib dikirim dari app.py, tidak ada fallback
    if not template_path:
        raise ValueError("template_path harus diberikan.")
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")

    doc = Document(template_path)

    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_NAME
    font.size = Pt(9)

    if is_blind:
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = ""

    # =========================================================================
    # 1. JUDUL
    # =========================================================================
    p = find_paragraph(doc, [H["judul_id"]])
    if p:
        fill_paragraph(p, getattr(naskah, 'judul_id', '').upper(),
                       size=15, bold=True, color=BLUE,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT, name=FONT_NAME)

    p = find_paragraph(doc, [H["judul_en"]])
    if p:
        fill_paragraph(p, getattr(naskah, 'judul_en', ''),
                       size=10, bold=True, italic=True, color=BLACK,
                       alignment=WD_ALIGN_PARAGRAPH.LEFT, name=FONT_NAME)

    p = find_paragraph(doc, [H["running"]])
    if p:
        fill_paragraph(p, getattr(naskah, 'running_title', ''), size=9, name=FONT_NAME)

    # =========================================================================
    # 2. PENULIS & AFILIASI
    # =========================================================================
    if not is_blind and naskah.penulis_list:
        parts = []
        for p_auth in naskah.penulis_list:
            nama = p_auth.nama if hasattr(p_auth, 'nama') else str(p_auth.get("nama", ""))
            aff = p_auth.afiliasi_ids if hasattr(p_auth, 'afiliasi_ids') else p_auth.get("afiliasi_ids", [1])
            corr = p_auth.is_corresp if hasattr(p_auth, 'is_corresp') else p_auth.get("is_corresp", False)
            aff_str = "".join(str(a) for a in aff) + ("*" if corr else "")
            parts.append(f"{nama}{aff_str}")

        p = find_paragraph(doc, [H["penulis"]])
        if p:
            fill_paragraph(p, ", ".join(parts), size=10, bold=True, color=BLACK, name=FONT_NAME)

        aff_text = "\n".join(f"{i} {a}" for i, a in enumerate(naskah.afiliasi_list, 1))
        p = find_paragraph(doc, [H["afiliasi"]])
        if p:
            fill_paragraph(p, aff_text, size=8, italic=True, color=BLACK, name=FONT_NAME)

    elif is_blind:
        blind_nama = "[ANONIM]" if not is_en else "[ANONYMOUS]"
        blind_aff  = "[AFILIASI DISEMBUNYIKAN]" if not is_en else "[AFFILIATION REMOVED]"

        p = find_paragraph(doc, [H["penulis"]])
        if p:
            fill_paragraph(p, blind_nama, size=10, bold=True, name=FONT_NAME)
        p = find_paragraph(doc, [H["afiliasi"]])
        if p:
            fill_paragraph(p, blind_aff, size=8, italic=True, name=FONT_NAME)

    # =========================================================================
    # 3. ABSTRAK & KATA KUNCI
    # =========================================================================
    p = find_paragraph(doc, [H["abs_en"]])
    if p:
        fill_paragraph(p, getattr(naskah, 'abstrak_en', ''), size=8, name=FONT_NAME)

    p = find_paragraph(doc, [H["kw_en"]])
    if p:
        fill_paragraph(p, f"Keywords: {getattr(naskah, 'kata_kunci_en', '')}", size=8, name=FONT_NAME)

    p = find_paragraph(doc, [H["abs_id"]])
    if p:
        fill_paragraph(p, getattr(naskah, 'abstrak_id', ''), size=8, name=FONT_NAME)

    p = find_paragraph(doc, [H["kw_id"]])
    if p:
        fill_paragraph(p, f"Kata kunci: {getattr(naskah, 'kata_kunci_id', '')}", size=8, name=FONT_NAME)

    # =========================================================================
    # 4. FOOTNOTE / KORESPONDENSI
    # =========================================================================
    if not is_blind:
        corresp = next((p for p in naskah.penulis_list
                        if (hasattr(p, 'is_corresp') and p.is_corresp) or
                           (isinstance(p, dict) and p.get("is_corresp"))), None)
        cn = corresp.nama if corresp and hasattr(corresp, 'nama') else "N/A"
        email_val = getattr(naskah, 'email_korespondensi', '')
        telepon_val = getattr(naskah, 'telepon', '')

        if is_en:
            ft = (f"*Corresponding author: {cn}\n"
                  f"Tel: {telepon_val}, E-mail: {email_val}")
        else:
            ft = (f"*Penulis korespondensi (Corresponding author): {cn}\n"
                  f"Tel: {telepon_val}, E-mail: {email_val}")

        p = find_paragraph(doc, [H["footnote"]])
        if p:
            fill_paragraph(p, ft, size=7.6, name=FONT_NAME)
    else:
        blind_ft = ("*Information removed for blind review" if is_en
                    else "*Informasi disembunyikan untuk blind review")
        p = find_paragraph(doc, [H["footnote"]])
        if p:
            fill_paragraph(p, blind_ft, size=7.6, name=FONT_NAME)

    # =========================================================================
    # 5. BAB 1–5 (KONTEN UTAMA)
    # =========================================================================
    sections = [
        (H["bab1"],   getattr(naskah, 'bab1', '')),
        (H["met21"],  getattr(naskah, 'metode_21', '')),
        (H["met22"],  getattr(naskah, 'metode_22', '')),
        (H["met221"], getattr(naskah, 'metode_221', '')),
        (H["met23"],  getattr(naskah, 'metode_23', '')),
        (H["bab3"],   getattr(naskah, 'bab3', '')),
        (H["bab4"],   getattr(naskah, 'bab4', '')),
        (H["bab5"],   getattr(naskah, 'bab5', '')),
    ]

    for heading, content in sections:
        if content and content.strip():
            p = find_paragraph(doc, [heading, "[Paragraf", "[Paragraph"])
            if p:
                fill_paragraph(p, content, size=9, name=FONT_NAME)

    # =========================================================================
    # 6. PERNYATAAN AKHIR
    # =========================================================================
    closing = [
        (H["konflik"],    getattr(naskah, 'konflik', '')),
        (H["dana"],       getattr(naskah, 'dana', '')),
        (H["ucapan"],     "" if is_blind else getattr(naskah, 'ucapan', '')),
        (H["kontribusi"], getattr(naskah, 'kontribusi', '')),
        (H["data_avail"], getattr(naskah, 'data_avail', '')),
        (H["etik"],       getattr(naskah, 'etik', '')),
    ]

    for label, content in closing:
        if content and content.strip():
            p = find_paragraph(doc, [label])
            if p:
                fill_paragraph(p, content.strip(), size=9, name=FONT_NAME)

    # =========================================================================
    # 7. DAFTAR PUSTAKA
    # =========================================================================
    refs = getattr(naskah, 'daftar_pustaka', '')
    if refs and refs.strip():
        p = find_paragraph(doc, [H["refs"]])
        if p:
            fill_paragraph(p, refs.strip(), size=9, name=FONT_NAME)

    # =========================================================================
    # 8. TABEL
    # =========================================================================
    if naskah.tabel_list:
        for tbl_data in naskah.tabel_list:
            data_text = tbl_data.get('data', '')
            cap       = tbl_data.get('cap', '')        # satu caption sesuai bahasa
            footnote  = tbl_data.get('catatan', '')
            nomor     = tbl_data.get('nomor', 1)

            if data_text:
                lines = data_text.strip().split('\n')
                headers = [h.strip() for h in lines[0].split(';')]
                rows = []
                for line in lines[1:]:
                    if line.strip():
                        rows.append([c.strip() for c in line.split(';')])

                # Hapus placeholder tabel dari template
                placeholder_keys = ['[Parameter]', '[Station 1]', '[Stasiun 1]']
                for tbl in doc.tables:
                    if any(k in cell.text for row in tbl.rows for cell in row.cells for k in placeholder_keys):
                        tbl._tbl.getparent().remove(tbl._tbl)
                        break

                create_table_in_doc(doc, headers, rows, nomor, cap, footnote, bahasa=bahasa)

    # =========================================================================
    # 9. SIMPAN
    # =========================================================================
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


# Alias untuk kompatibilitas
build_docx = bangun
