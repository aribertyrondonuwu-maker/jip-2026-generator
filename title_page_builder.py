"""
title_page_builder.py — Generator Halaman Judul (Title Page) JIP 2026
Mengisi Template_Title_Page_JIP_2026.docx dengan data dari form Streamlit.
Template final: 5 kolom tabel (No, Nama, Aff, ORCID, Email)
"""

import io
import copy
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

TEMPLATE_TITLE_PAGE = "Template_Title_Page_JIP_2026.docx"

CREDIT_ROLES = [
    "Conceptualization", "Data curation", "Formal analysis",
    "Funding acquisition", "Investigation", "Methodology",
    "Project administration", "Resources", "Software",
    "Supervision", "Validation", "Visualization",
    "Writing – original draft", "Writing – review & editing",
]


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI BANTU
# ─────────────────────────────────────────────────────────────────────────────

def _all_paragraphs(doc):
    """Iterasi semua paragraf: body + dalam tabel."""
    for p in doc.paragraphs:
        yield p
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _find(doc, *keywords):
    """Paragraf pertama yang mengandung salah satu keyword."""
    for p in _all_paragraphs(doc):
        if any(kw in p.text for kw in keywords):
            return p
    return None


def _find_all(doc, *keywords):
    """Semua paragraf yang mengandung salah satu keyword."""
    result = []
    for p in _all_paragraphs(doc):
        if any(kw in p.text for kw in keywords):
            result.append(p)
    return result


def _set_run(run, text, bold=None, italic=None, color=None, size=None):
    """Set properti run tanpa mengubah run lain."""
    run.text = text
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    if size is not None:
        run.font.size = Pt(size)


def _fill2(para, label, value):
    """
    Isi paragraf dengan 2 run: run[0]=label (bold), run[1]=value (italic).
    Mempertahankan format asli dari template.
    """
    runs = para.runs
    if len(runs) >= 2:
        runs[0].text = label
        runs[1].text = value
        # Hapus run berlebih
        for r in runs[2:]:
            r._r.getparent().remove(r._r)
    elif len(runs) == 1:
        runs[0].text = value
    else:
        r = para.add_run(value)
        r.font.italic = True


def _fill1(para, value):
    """
    Isi paragraf dengan 1 run (italic), pertahankan format asli.
    """
    runs = para.runs
    if runs:
        runs[0].text = value
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        r = para.add_run(value)
        r.font.italic = True


def _set_cell(cell, text, bold=False, italic=False, color="555555"):
    """Isi cell tabel dengan teks + format."""
    para = cell.paragraphs[0]
    for run in list(para.runs):
        run._r.getparent().remove(run._r)
    run = para.add_run(text)
    run.font.name = "Cambria"
    run.font.size = Pt(8)
    run.font.bold = bold
    run.font.italic = italic
    r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    run.font.color.rgb = RGBColor(r, g, b)


def _clone_row(tr):
    """Duplikat baris tabel (struktur XML)."""
    new_tr = copy.deepcopy(tr)
    # Kosongkan semua teks
    for tc in new_tr.findall(qn('w:tc')):
        for p_el in tc.findall(qn('w:p')):
            for r_el in p_el.findall(qn('w:r')):
                t_el = r_el.find(qn('w:t'))
                if t_el is not None:
                    t_el.text = ""
    return new_tr


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI UTAMA
# ─────────────────────────────────────────────────────────────────────────────

def bangun_title_page(data: dict, template_path: str = TEMPLATE_TITLE_PAGE) -> bytes:
    """
    Isi template Title Page dengan data dari form.

    data = {
        ojs_id          : str
        judul_id        : str
        judul_en        : str
        running_title   : str
        penulis         : list[{nama, aff, is_corresp, orcid, email}]
        afiliasi        : dict {"a": "...", "b": "..."}
        koresp_nama     : str
        koresp_afiliasi : str
        koresp_alamat   : str
        koresp_telepon  : str
        koresp_email    : str
        koresp_orcid    : str
        kontribusi      : dict {nama: "roles..."}
        ucapan          : str
        pendanaan       : str
        konflik         : str
    }
    """
    if not Path(template_path).exists():
        raise FileNotFoundError(f"Template tidak ditemukan: {template_path}")

    doc = Document(template_path)

    # ── 0. ID NASKAH OJS ────────────────────────────────────────────────────
    p = _find(doc, "ID Naskah OJS:")
    if p:
        _fill2(p, "ID Naskah OJS: ", data.get("ojs_id", ""))

    # ── 1. JUDUL ─────────────────────────────────────────────────────────────
    p = _find(doc, "Judul (Bahasa Indonesia):")
    if p:
        _fill2(p, "Judul (Bahasa Indonesia): ", data.get("judul_id", ""))

    p = _find(doc, "Title (English):")
    if p:
        _fill2(p, "Title (English): ", data.get("judul_en", ""))

    p = _find(doc, "Judul Singkat / Running Title:")
    if p:
        _fill2(p, "Judul Singkat / Running Title: ", data.get("running_title", ""))

    # ── 2. TABEL PENULIS (5 kolom: No | Nama | Aff | ORCID | Email) ─────────
    penulis_list = data.get("penulis", [])
    if penulis_list:
        # Cari tabel dengan header "No." & "Email"
        target_tbl = None
        for tbl in doc.tables:
            header = [c.text.strip() for c in tbl.rows[0].cells]
            if "No." in header and "Email" in header:
                target_tbl = tbl
                break

        if target_tbl is not None:
            # Hapus semua baris data (baris 1 dst)
            while len(target_tbl.rows) > 1:
                tr = target_tbl.rows[-1]._tr
                tr.getparent().remove(tr)

            # Tambah baris baru tiap penulis
            template_tr = target_tbl.rows[0]._tr  # pakai header sebagai template klon
            for i, p_data in enumerate(penulis_list, 1):
                new_tr = _clone_row(template_tr)
                target_tbl._tbl.append(new_tr)
                row = target_tbl.rows[-1]
                cells = row.cells

                nama        = p_data.get("nama", "")
                aff         = p_data.get("aff", "")
                is_corresp  = p_data.get("is_corresp", False)
                orcid       = p_data.get("orcid", "—") or "—"
                email       = p_data.get("email", "—") or "—"

                nama_label  = nama + ("  ★" if is_corresp else "")
                aff_label   = f"({aff})" if aff else "—"

                _set_cell(cells[0], str(i) + ".",  bold=True,  color="007BB8")
                _set_cell(cells[1], nama_label,    italic=True, color="555555")
                _set_cell(cells[2], aff_label,     bold=True,  color="007BB8")
                _set_cell(cells[3], orcid,         italic=True, color="555555")
                _set_cell(cells[4], email,         italic=True, color="555555")

    # ── KETERANGAN AFILIASI ───────────────────────────────────────────────────
    afiliasi = data.get("afiliasi", {})
    for huruf, teks in afiliasi.items():
        label = f"({huruf})"
        for p in doc.paragraphs:
            if p.text.strip().startswith(label):
                runs = p.runs
                if len(runs) >= 2:
                    runs[1].text = teks
                elif runs:
                    runs[0].text = f"{label} {teks}"
                break

    # ── 3. PENULIS KORESPONDENSI ─────────────────────────────────────────────
    fields = {
        "Nama Lengkap / Full Name: ":   data.get("koresp_nama", ""),
        "Afiliasi / Affiliation: ":     data.get("koresp_afiliasi", ""),
        "Alamat / Address: ":           data.get("koresp_alamat", ""),
        "Telepon / Phone: ":            data.get("koresp_telepon", ""),
        "E-mail: ":                     data.get("koresp_email", ""),
        "ORCID iD: ":                   data.get("koresp_orcid", ""),
    }
    for label, value in fields.items():
        p = _find(doc, label.strip())
        if p:
            _fill2(p, label, value)

    # ── 4. KONTRIBUSI PENULIS (CRediT) ───────────────────────────────────────
    kontribusi = data.get("kontribusi", {})
    for nama, roles in kontribusi.items():
        if not roles:
            continue
        for p in _all_paragraphs(doc):
            txt = p.text.strip()
            if txt.startswith(nama + ":") or txt.startswith(nama + " :"):
                runs = p.runs
                if len(runs) >= 2:
                    runs[0].text = nama + ": "
                    runs[1].text = roles
                    for r in runs[2:]:
                        r._r.getparent().remove(r._r)
                elif runs:
                    runs[0].text = f"{nama}: {roles}"
                break

    # ── 5. UCAPAN TERIMA KASIH ───────────────────────────────────────────────
    p = _find(doc, "Penulis mengucapkan terima kasih",
              "The authors thank the Bunaken")
    if p:
        _fill1(p, data.get("ucapan", ""))

    # ── 6. PERNYATAAN PENDANAAN ───────────────────────────────────────────────
    # Baris pendanaan aktual = paragraf yang berisi "Direktorat Riset"
    p = _find(doc, "Penelitian ini didanai oleh Direktorat",
              "This study was funded by the Directorate")
    if p:
        _fill1(p, data.get("pendanaan", ""))

    # ── 7. KONFLIK KEPENTINGAN ────────────────────────────────────────────────
    # Ada dua paragraf yang mengandung teks konflik:
    # [45] = di dalam kotak contoh "Tanpa konflik / No conflict: ..."
    # [47] = paragraf aktual yang harus diisi
    # Ambil yang TERAKHIR
    matches = _find_all(doc, "Para penulis menyatakan tidak ada konflik",
                        "The authors declare no conflict")
    if matches:
        _fill1(matches[-1], data.get("konflik", ""))

    # ── Simpan ────────────────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def get_credit_roles():
    return CREDIT_ROLES
