"""
platx_builder.py — Versi Minimal & Aman
"""

import io
import re
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

try:
    from latex2mathml.converter import convert as latex_to_mathml
    HAS_LATEX = True
except:
    HAS_LATEX = False

FONT = 'Cambria'
BLUE = RGBColor(0x00, 0x7B, 0xB8)
BLACK = RGBColor(0x00, 0x00, 0x00)


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


def set_run(run, size=None, bold=None, italic=None, color=None, name=FONT):
    if name: run.font.name = name
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if italic is not None: run.font.italic = italic
    if color: run.font.color.rgb = color


def find_para(doc, keywords):
    """Cari paragraf yang mengandung salah satu keyword."""
    for para in doc.paragraphs:
        for kw in keywords:
            if kw in para.text:
                return para
    return None


def fill_para(para, text, size=9, bold=False, italic=False, color=BLACK):
    """Isi paragraf dengan text baru, hapus isi lama dengan aman."""
    # Hapus semua run kecuali yang pertama
    runs = list(para.runs)
    if len(runs) > 1:
        for run in runs[1:]:
            r = run._r
            r.getparent().remove(r)
    # Set text di run pertama
    if runs:
        runs[0].text = text
        set_run(runs[0], size=size, bold=bold, italic=italic, color=color)
    else:
        r = para.add_run(text)
        set_run(r, size=size, bold=bold, italic=italic, color=color)


def bangun(naskah=None, template_path=None, *args, **kwargs):
    if naskah is None and args:
        naskah = args[0]
    
    is_blind = getattr(naskah, 'blind', False)
    is_en = getattr(naskah, 'bahasa', 'id') == 'en'
    
    tpl = template_path or "Template_Artikel_PLATAX_2026_FINAL_OJS.docx"
    if not Path(tpl).exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {tpl}")
    
    doc = Document(tpl)
    
    # Blind review
    if is_blind:
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = ""
    
    # JUDUL
    p = find_para(doc, ["[Judul penelitian dalam bahasa Indonesia:"])
    if p: fill_para(p, getattr(naskah, 'judul_id', '').upper(), size=15, bold=True, color=BLUE)
    
    p = find_para(doc, ["[Complete Research Title in English:"])
    if p: fill_para(p, getattr(naskah, 'judul_en', ''), size=10, bold=True, italic=True)
    
    p = find_para(doc, ["[Judul singkat, maks. 60 karakter]"])
    if p: fill_para(p, getattr(naskah, 'running_title', ''), size=9)
    
    # PENULIS
    if not is_blind and naskah.penulis_list:
        parts = []
        for p_auth in naskah.penulis_list:
            nama = p_auth.nama if hasattr(p_auth, 'nama') else p_auth.get("nama", "")
            aff = p_auth.afiliasi_ids if hasattr(p_auth, 'afiliasi_ids') else p_auth.get("afiliasi_ids", [1])
            corr = p_auth.is_corresp if hasattr(p_auth, 'is_corresp') else p_auth.get("is_corresp", False)
            aff_str = "".join(str(a) for a in aff) + ("*" if corr else "")
            parts.append(f"{nama}{aff_str}")
        p = find_para(doc, ["[Nama Penulis 1]"])
        if p: fill_para(p, ", ".join(parts), size=9)
        
        aff_lines = "\n".join(f"{i+1} {a}" for i, a in enumerate(naskah.afiliasi_list))
        p = find_para(doc, ["[Departemen, Fakultas"])
        if p: fill_para(p, aff_lines, size=9, italic=True)
    elif is_blind:
        p = find_para(doc, ["[Nama Penulis 1]"])
        if p: fill_para(p, "[ANONIM]", size=9)
        p = find_para(doc, ["[Departemen, Fakultas"])
        if p: fill_para(p, "[AFILIASI DISEMBUNYIKAN]", size=9, italic=True)
    
    # ABSTRAK
    p = find_para(doc, ["[Tulis abstract bahasa Inggris di sini."])
    if p: fill_para(p, getattr(naskah, 'abstrak_en', ''), size=8)
    
    p = find_para(doc, ["Keywords: [keyword 1"])
    if p: fill_para(p, f"Keywords: {getattr(naskah, 'kata_kunci_en', '')}", size=8)
    
    p = find_para(doc, ["[Tulis abstrak bahasa Indonesia di sini"])
    if p: fill_para(p, getattr(naskah, 'abstrak_id', ''), size=8)
    
    p = find_para(doc, ["Kata kunci: [kata kunci 1"])
    if p: fill_para(p, f"Kata kunci: {getattr(naskah, 'kata_kunci_id', '')}", size=8)
    
    # FOOTNOTE
    if not is_blind:
        corresp = next((p for p in naskah.penulis_list 
                        if (hasattr(p, 'is_corresp') and p.is_corresp) or 
                           (isinstance(p, dict) and p.get("is_corresp"))), None)
        cn = corresp.nama if corresp and hasattr(corresp, 'nama') else "N/A"
        ft = f"*Penulis korespondensi: {cn}\nTel: {getattr(naskah, 'telepon', '')}, E-mail: {getattr(naskah, 'email_korespondensi', '')}"
        p = find_para(doc, ["*Penulis korespondensi (Corresponding author):"])
        if p: fill_para(p, ft, size=7.6)
    else:
        p = find_para(doc, ["*Penulis korespondensi (Corresponding author):"])
        if p: fill_para(p, "*Informasi disembunyikan untuk blind review", size=7.6)
    
    # BAB 1-5
    sections = [
        ("1. Pendahuluan", getattr(naskah, 'bab1', '')),
        ("2.1. Waktu dan lokasi penelitian", getattr(naskah, 'metode_21', '')),
        ("2.2. Pengumpulan data", getattr(naskah, 'metode_22', '')),
        ("2.2.1. Analisis laboratorium", getattr(naskah, 'metode_221', '')),
        ("2.3. Analisis data", getattr(naskah, 'metode_23', '')),
        ("3. Hasil", getattr(naskah, 'bab3', '')),
        ("4. Pembahasan", getattr(naskah, 'bab4', '')),
        ("5. Simpulan", getattr(naskah, 'bab5', '')),
    ]
    for heading, content in sections:
        if content and content.strip():
            p = find_para(doc, [heading, "[Paragraf"])
            if p: fill_para(p, content, size=9)
    
    # PERNYATAAN AKHIR
    closing = [
        ("Konflik kepentingan (Competing interests)", getattr(naskah, 'konflik', '')),
        ("Sumber dana (Funding sources)", getattr(naskah, 'dana', '')),
        ("Ucapan terima kasih (Acknowledgements)", "" if is_blind else getattr(naskah, 'ucapan_terima_kasih', '')),
        ("Kontribusi penulis (Authors' contributions)", getattr(naskah, 'kontribusi', '')),
        ("Ketersediaan data (Availability of data and materials)", getattr(naskah, 'data_avail', '')),
        ("Persetujuan etik (Ethics approval and consent to participate)", getattr(naskah, 'etik', '')),
    ]
    for label, content in closing:
        if content and content.strip():
            p = find_para(doc, [label])
            if p: fill_para(p, content.strip(), size=9)
    
    # DAFTAR PUSTAKA
    refs = getattr(naskah, 'daftar_pustaka', '')
    if refs and refs.strip():
        p = find_para(doc, ["[APA 7th Edition"])
        if p: fill_para(p, refs.strip(), size=7.6)
    
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

build_docx = bangun