"""
title_page_builder.py — Generator Halaman Judul (Title Page) JIP 2026
Mengisi Template_Title_Page_JIP_2026.docx dengan data dari form Streamlit.
"""

import io
import re
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml, OxmlElement
from lxml import etree

TEMPLATE_TITLE_PAGE = "Template_Title_Page_JIP_2026.docx"

CREDIT_ROLES = [
    "Conceptualization", "Data curation", "Formal analysis",
    "Funding acquisition", "Investigation", "Methodology",
    "Project administration", "Resources", "Software",
    "Supervision", "Validation", "Visualization",
    "Writing – original draft", "Writing – review & editing",
]


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI BANTU XML
# ─────────────────────────────────────────────────────────────────────────────

def _iter_paragraphs(doc):
    """Iterasi semua paragraf termasuk yang di dalam tabel."""
    for para in doc.paragraphs:
        yield para
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para


def _find_para(doc, *keywords):
    """Cari paragraf pertama yang mengandung salah satu keyword."""
    for para in _iter_paragraphs(doc):
        txt = para.text
        if any(kw in txt for kw in keywords):
            return para
    return None


def _find_para_exact(doc, *keywords):
    """Cari paragraf yang teks-nya dimulai dengan salah satu keyword."""
    for para in _iter_paragraphs(doc):
        txt = para.text.strip()
        if any(txt.startswith(kw) for kw in keywords):
            return para
    return None


def _copy_rpr(src_run, dst_run):
    """Salin properti run (font, size, bold, italic, color) dari src ke dst."""
    src_rpr = src_run._r.find(qn('w:rPr'))
    if src_rpr is None:
        return
    import copy
    dst_rpr = copy.deepcopy(src_rpr)
    old = dst_run._r.find(qn('w:rPr'))
    if old is not None:
        dst_run._r.remove(old)
    dst_run._r.insert(0, dst_rpr)


def _fill_para(para, label_text, value_text, label_run_idx=0, value_run_idx=1):
    """
    Isi paragraf dengan mempertahankan format run asli.
    Run pertama = label (bold), run kedua = value (italic/grey).
    Jika hanya ada 1 run, pakai seluruhnya untuk value.
    """
    runs = para.runs
    if not runs:
        return

    if len(runs) == 1:
        runs[0].text = value_text
        return

    # Bersihkan semua run lebih dari 2
    while len(para.runs) > 2:
        para.runs[-1]._r.getparent().remove(para.runs[-1]._r)

    para.runs[0].text = label_text
    para.runs[1].text = value_text


def _replace_text_preserve_format(para, new_text):
    """
    Ganti seluruh teks paragraf, pertahankan format run pertama.
    Hapus run sisanya.
    """
    runs = para.runs
    if not runs:
        r = para.add_run(new_text)
        return
    runs[0].text = new_text
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def bangun_title_page(data: dict, template_path: str = TEMPLATE_TITLE_PAGE) -> bytes:
    """
    Isi template Title Page dengan data dari form.

    Parameter `data` (dict):
        ojs_id          : str  — ID Naskah OJS, mis. "PLATAX-2026-0042"
        judul_id        : str  — Judul Bahasa Indonesia
        judul_en        : str  — Title in English
        running_title   : str  — Running title
        penulis         : list[dict] — list penulis, tiap dict:
                            {nama, aff, is_corresp, orcid}
        afiliasi        : dict[str, str] — {"a": "...", "b": "..."}
        koresp_nama     : str
        koresp_afiliasi : str
        koresp_alamat   : str
        koresp_telepon  : str
        koresp_email    : str
        koresp_orcid    : str
        kontribusi      : dict[str, str] — {nama_penulis: "roles..."}
        ucapan          : str
        pendanaan       : str  — teks pendanaan
        konflik         : str
    """

    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")

    doc = Document(template_path)

    # ── 0. ID NASKAH OJS ────────────────────────────────────────────────────
    p = _find_para(doc, "ID Naskah OJS:", "OJS Submission ID:")
    if p:
        _fill_para(p,
                   label_text="ID Naskah OJS: ",
                   value_text=data.get("ojs_id", ""))

    # ── 1. JUDUL ─────────────────────────────────────────────────────────────
    p = _find_para(doc, "Judul (Bahasa Indonesia):")
    if p:
        _fill_para(p,
                   label_text="Judul (Bahasa Indonesia): ",
                   value_text=data.get("judul_id", ""))

    p = _find_para(doc, "Title (English):")
    if p:
        _fill_para(p,
                   label_text="Title (English): ",
                   value_text=data.get("judul_en", ""))

    p = _find_para(doc, "Judul Singkat / Running Title:")
    if p:
        _fill_para(p,
                   label_text="Judul Singkat / Running Title: ",
                   value_text=data.get("running_title", ""))

    # ── 2. TABEL PENULIS ─────────────────────────────────────────────────────
    _isi_tabel_penulis(doc, data.get("penulis", []))

    # Keterangan afiliasi
    _isi_afiliasi(doc, data.get("afiliasi", {}))

    # ── 3. PENULIS KORESPONDENSI ─────────────────────────────────────────────
    kmap = {
        "Nama Lengkap / Full Name:": "koresp_nama",
        "Afiliasi / Affiliation:":   "koresp_afiliasi",
        "Alamat / Address:":         "koresp_alamat",
        "Telepon / Phone:":          "koresp_telepon",
        "E-mail:":                   "koresp_email",
        "ORCID iD:":                 "koresp_orcid",
    }
    for label, key in kmap.items():
        p = _find_para(doc, label)
        if p:
            _fill_para(p, label_text=label + " ", value_text=data.get(key, ""))

    # ── 4. KONTRIBUSI PENULIS (CRediT) ───────────────────────────────────────
    kontribusi = data.get("kontribusi", {})
    _isi_kontribusi(doc, kontribusi)

    # ── 5. UCAPAN TERIMA KASIH ───────────────────────────────────────────────
    p = _find_para(doc,
                   "Penulis mengucapkan terima kasih",
                   "The authors thank")
    # Cari paragraf di bawah heading ucapan terima kasih yang berisi teks contoh
    if p is None:
        # Cari paragraf setelah heading section 5
        p = _find_para(doc, "Penulis mengucapkan", "The authors thank")
    if p:
        _replace_text_preserve_format(p, data.get("ucapan", ""))

    # ── 6. PERNYATAAN PENDANAAN ───────────────────────────────────────────────
    # Cari baris yang berisi teks pendanaan contoh (bukan baris "contoh hibah")
    p = _find_para(doc,
                   "Penelitian ini didanai oleh Direktorat",
                   "This study was funded by the Directorate")
    if p:
        _replace_text_preserve_format(p, data.get("pendanaan", ""))

    # ── 7. KONFLIK KEPENTINGAN ────────────────────────────────────────────────
    p = _find_para(doc,
                   "Para penulis menyatakan tidak ada konflik",
                   "The authors declare no conflict")
    # Ada dua paragraf yang match; cari yang KEDUA (yang bukan contoh di kotak abu)
    matches = []
    for para in _iter_paragraphs(doc):
        if ("Para penulis menyatakan tidak ada konflik" in para.text or
                "The authors declare no conflict" in para.text):
            matches.append(para)
    # Paragraf terakhir = yang di luar kotak contoh
    if matches:
        _replace_text_preserve_format(matches[-1], data.get("konflik", ""))

    # ── Simpan ────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: ISI TABEL PENULIS
# ─────────────────────────────────────────────────────────────────────────────

def _isi_tabel_penulis(doc, penulis_list: list):
    """
    Isi tabel penulis di template.
    Template memiliki: baris header + 3 baris data + 1 baris "..."
    Kita hapus semua baris data lama, ganti dengan data baru.
    """
    if not penulis_list:
        return

    # Cari tabel yang punya header "No." / "Nama Penulis"
    target_tbl = None
    for tbl in doc.tables:
        header_cells = [c.text.strip() for c in tbl.rows[0].cells]
        if any("No" in c for c in header_cells) and any("Nama" in c or "Name" in c for c in header_cells):
            target_tbl = tbl
            break

    if target_tbl is None:
        return

    # Hapus semua baris data (baris 1 dst, pertahankan baris 0 = header)
    while len(target_tbl.rows) > 1:
        tr = target_tbl.rows[-1]._tr
        tr.getparent().remove(tr)

    # Tambah baris baru untuk setiap penulis
    for i, p in enumerate(penulis_list, 1):
        nama = p.get("nama", "")
        aff  = p.get("aff", "")
        orcid = p.get("orcid", "—")
        is_corresp = p.get("is_corresp", False)

        # Buat baris baru dengan menyalin baris header sebagai template dasar
        new_tr = _clone_row(target_tbl.rows[0]._tr)
        target_tbl._tbl.append(new_tr)

        row = target_tbl.rows[-1]
        cells = row.cells

        # Kolom 0: nomor
        _set_cell_text(cells[0], str(i) + ".", bold=True, color="007BB8")

        # Kolom 1: nama (tambah ★ jika korespondensi)
        nama_val = nama + ("  ★" if is_corresp else "")
        _set_cell_text(cells[1], nama_val, italic=True, color="555555")

        # Kolom 2: afiliasi
        _set_cell_text(cells[2], f"({aff})" if aff else "", bold=True, color="007BB8")

        # Kolom 3: ORCID
        _set_cell_text(cells[3], orcid, italic=True, color="555555")


def _clone_row(tr):
    """Clone sebuah <w:tr> untuk dijadikan baris baru."""
    import copy
    new_tr = copy.deepcopy(tr)
    # Hapus konten teks dari semua sel, pertahankan format
    for tc in new_tr.findall(qn('w:tc')):
        for p_elem in tc.findall(qn('w:p')):
            for r_elem in p_elem.findall(qn('w:r')):
                t_elem = r_elem.find(qn('w:t'))
                if t_elem is not None:
                    t_elem.text = ""
    return new_tr


def _set_cell_text(cell, text: str, bold=False, italic=False, color=None):
    """Set teks sebuah cell dengan formatting."""
    para = cell.paragraphs[0]
    # Bersihkan runs lama
    for run in list(para.runs):
        run._r.getparent().remove(run._r)
    run = para.add_run(text)
    # Terapkan format dasar
    from docx.shared import Pt, RGBColor
    run.font.name = "Cambria"
    run.font.size = Pt(8)
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if color:
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        run.font.color.rgb = RGBColor(r, g, b)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: ISI KETERANGAN AFILIASI
# ─────────────────────────────────────────────────────────────────────────────

def _isi_afiliasi(doc, afiliasi: dict):
    """
    Isi paragraf keterangan afiliasi.
    afiliasi = {"a": "Program Studi ...", "b": "Dept ..."}
    """
    # Cari paragraf yang berisi teks "(a)" atau "(b)" sebagai label afiliasi
    # (di luar tabel penulis)
    for key, teks in afiliasi.items():
        label = f"({key})"
        # Iterasi paragraf di luar tabel
        for para in doc.paragraphs:
            if para.text.strip().startswith(label):
                # Pertahankan label, ganti teks afiliasi
                runs = para.runs
                if len(runs) >= 2:
                    runs[1].text = teks
                elif runs:
                    runs[0].text = f"{label} {teks}"
                break


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: ISI KONTRIBUSI PENULIS
# ─────────────────────────────────────────────────────────────────────────────

def _isi_kontribusi(doc, kontribusi: dict):
    """
    kontribusi = {
        "Maria Sari Wulandari": "Conceptualization, Methodology, ...",
        "Johan Paulus Rompas": "Investigation, Data curation",
    }
    Cari paragraf dengan nama penulis tersebut dan ganti kontribusinya.
    """
    if not kontribusi:
        return

    for nama, roles in kontribusi.items():
        # Cari paragraf yang dimulai dengan nama tersebut
        for para in _iter_paragraphs(doc):
            if para.text.strip().startswith(nama + ":") or para.text.strip().startswith(nama + " :"):
                runs = para.runs
                if len(runs) >= 2:
                    runs[0].text = nama + ": "
                    runs[1].text = roles
                elif runs:
                    runs[0].text = f"{nama}: {roles}"
                break


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: DAFTAR PERAN CRediT (untuk UI Streamlit)
# ─────────────────────────────────────────────────────────────────────────────

def get_credit_roles():
    return CREDIT_ROLES
