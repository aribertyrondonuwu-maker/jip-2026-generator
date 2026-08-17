"""
platax_builder.py — pengisi Template Artikel PLATAX 2026.

Berbeda dari pendekatan sebelumnya, modul ini TIDAK membangun dokumen dari nol.
Modul ini membuka Template_Artikel_PLATAX_2026_FINAL_OJS.docx dan mengganti
teks placeholder di dalamnya, sehingga masthead, logo, sampul edisi, kotak
abstrak, style, margin, header, footer, dan seluruh tipografi terbawa apa adanya.

Tidak mengimpor streamlit — bisa diuji terpisah.
"""

from __future__ import annotations

import copy
import io
import re
from dataclasses import dataclass, field

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, RGBColor
from docx.text.paragraph import Paragraph

# Warna sitiran dalam naskah, sesuai isi template asli (bukan 2196D1).
BIRU_SITIRAN = RGBColor(0x00, 0x7B, 0xB8)
ABU_KETERANGAN = "808080"      # warna keterangan format dalam kurung siku
LEBAR_TABEL_DXA = 4333         # lebar tabel satu kolom pada template (≈7,64 cm)


# =========================================================
# STRUKTUR DATA
# =========================================================
@dataclass
class Penulis:
    nama: str
    afil: list                 # daftar nomor afiliasi, mis. [1] atau [1, 2]
    koresp: bool = False
    orcid: str = ""


@dataclass
class Naskah:
    judul_id: str = ""
    judul_en: str = ""
    running_title: str = ""
    penulis: list = field(default_factory=list)
    afiliasi: list = field(default_factory=list)
    telepon: str = ""
    email: str = ""

    abstract_en: str = ""
    keywords_en: str = ""
    abstrak_id: str = ""
    kata_kunci_id: str = ""

    bab1: str = ""
    metode_21: str = ""
    metode_22: str = ""
    metode_221: str = ""
    metode_23: str = ""
    bab3: str = ""
    bab4: str = ""
    bab5: str = ""

    cap_tabel_id: str = ""
    cap_tabel_en: str = ""
    tabel_data: str = ""
    cat_tabel: str = ""

    gambar_blob: bytes = b""
    lebar_gambar: float = 7.6
    cap_gambar_id: str = ""
    cap_gambar_en: str = ""

    konflik: str = ""
    dana: str = ""
    ucapan: str = ""
    kontribusi: str = ""
    data_avail: str = ""
    etik: str = ""
    daftar_pustaka: str = ""

    blind: bool = True


# =========================================================
# HELPER XML
# =========================================================
def hapus_paragraf(p: Paragraph) -> None:
    p._element.getparent().remove(p._element)


def rprot(p: Paragraph, idx: int = 0):
    """Salin rPr dari run ke-idx sebagai cetakan format."""
    if len(p.runs) > idx:
        return copy.deepcopy(p.runs[idx]._r.get_or_add_rPr())
    return None


def tulis(p: Paragraph, segmen, cetakan=None) -> Paragraph:
    """
    Tulis ulang isi paragraf dari daftar segmen [(teks, rPr_atau_None), ...],
    membuang seluruh run lama. rPr None memakai `cetakan`.
    """
    if cetakan is None:
        cetakan = rprot(p)
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    for teks, rpr in segmen:
        run = p.add_run(teks)
        sumber = rpr if rpr is not None else cetakan
        if sumber is not None:
            run._r.insert(0, copy.deepcopy(sumber))
    return p


def isi_run_pertama(p: Paragraph, teks: str, idx: int = 0) -> None:
    """Ganti teks satu run tertentu, biarkan run lain apa adanya."""
    if len(p.runs) > idx:
        p.runs[idx].text = teks


def buang_keterangan_abu(doc) -> int:
    """Hapus seluruh run keterangan format abu-abu (Petunjuk §1 & §10)."""
    n = 0
    for el in doc.element.body.iter(qn("w:r")):
        rpr = el.find(qn("w:rPr"))
        if rpr is None:
            continue
        col = rpr.find(qn("w:color"))
        if col is not None and (col.get(qn("w:val")) or "").upper() == ABU_KETERANGAN:
            el.getparent().remove(el)
            n += 1
    return n


def cari(doc, awalan: str, mulai: int = 0) -> int:
    """Indeks paragraf pertama yang teksnya diawali `awalan`."""
    for i in range(mulai, len(doc.paragraphs)):
        if doc.paragraphs[i].text.strip().startswith(awalan):
            return i
    raise KeyError(f"Placeholder tidak ditemukan: {awalan!r}")


def paragraf(doc, awalan: str) -> Paragraph:
    return doc.paragraphs[cari(doc, awalan)]


# =========================================================
# PEWARNAAN SITIRAN (Petunjuk §3 — nama, tahun, titik koma biru)
# =========================================================
POLA_KURUNG = re.compile(r"\(([^()]*?\b(?:1[89]|20)\d{2}[a-z]?[^()]*?)\)")
POLA_NARATIF = re.compile(
    r"\b([A-ZÀ-Þ][\w’'\-]+(?:\s+(?:&|dan|and)\s+[A-ZÀ-Þ][\w’'\-]+)?"
    r"(?:\s+et\s+al\.)?)\s+\(((?:1[89]|20)\d{2}[a-z]?)\)"
)


def segmen_sitiran(teks: str, cetakan):
    """Pecah teks menjadi segmen; bagian sitiran diberi rPr berwarna biru."""
    rentang = []
    for m in POLA_KURUNG.finditer(teks):
        rentang.append((m.start(1), m.end(1)))
    for m in POLA_NARATIF.finditer(teks):
        if any(a <= m.start(1) < b for a, b in rentang):
            continue
        rentang.append((m.start(1), m.end(1)))
        rentang.append((m.start(2), m.end(2)))
    rentang.sort()

    biru = copy.deepcopy(cetakan) if cetakan is not None else OxmlElement("w:rPr")
    lama = biru.find(qn("w:color"))
    if lama is not None:
        biru.remove(lama)
    col = OxmlElement("w:color")
    col.set(qn("w:val"), "007BB8")
    biru.append(col)

    keluar, kursor = [], 0
    for a, b in rentang:
        if a < kursor:
            continue
        if a > kursor:
            keluar.append((teks[kursor:a], None))
        keluar.append((teks[a:b], biru))
        kursor = b
    if kursor < len(teks):
        keluar.append((teks[kursor:], None))
    return keluar or [(teks, None)]


# =========================================================
# PENGISIAN BLOK PARAGRAF
# =========================================================
def isi_blok(doc, awalan: str, isi: str, jumlah_placeholder: int = 1,
             warnai: bool = False) -> None:
    """
    Ganti `jumlah_placeholder` paragraf berturut-turut mulai dari paragraf
    berawalan `awalan` dengan paragraf-paragraf dari `isi` (satu baris = satu
    paragraf). Format paragraf pertama dipakai sebagai cetakan.
    """
    baris = [b.strip() for b in isi.split("\n") if b.strip()] or [""]
    i = cari(doc, awalan)
    proto = doc.paragraphs[i]
    cetakan = rprot(proto)
    lama = [doc.paragraphs[i + k] for k in range(jumlah_placeholder)]

    target, terakhir = [proto], proto
    for _ in range(len(baris) - 1):
        baru = copy.deepcopy(proto._p)
        terakhir._p.addnext(baru)
        terakhir = Paragraph(baru, proto._parent)
        target.append(terakhir)

    for p in lama[1:]:
        hapus_paragraf(p)

    for p, teks in zip(target, baris):
        segmen = segmen_sitiran(teks, cetakan) if warnai else [(teks, None)]
        tulis(p, segmen, cetakan)


# =========================================================
# TABEL DATA
# =========================================================
def _set_jumlah_kolom(tabel, n: int) -> None:
    """Tambah/kurangi kolom dengan mengkloning kolom terakhir."""
    grid = tabel._tbl.find(qn("w:tblGrid"))
    kolom_sekarang = len(grid.findall(qn("w:gridCol")))

    while kolom_sekarang < n:
        grid.append(copy.deepcopy(grid.findall(qn("w:gridCol"))[-1]))
        for tr in tabel._tbl.findall(qn("w:tr")):
            tr.append(copy.deepcopy(tr.findall(qn("w:tc"))[-1]))
        kolom_sekarang += 1
    while kolom_sekarang > n:
        grid.remove(grid.findall(qn("w:gridCol"))[-1])
        for tr in tabel._tbl.findall(qn("w:tr")):
            tr.remove(tr.findall(qn("w:tc"))[-1])
        kolom_sekarang -= 1

    lebar = int(LEBAR_TABEL_DXA / n)
    for gc in grid.findall(qn("w:gridCol")):
        gc.set(qn("w:w"), str(lebar))
    for tr in tabel._tbl.findall(qn("w:tr")):
        for tc in tr.findall(qn("w:tc")):
            tcw = tc.find(qn("w:tcPr")).find(qn("w:tcW"))
            if tcw is not None:
                tcw.set(qn("w:w"), str(lebar))


def isi_tabel(tabel, data_teks: str) -> None:
    baris = [b for b in data_teks.split("\n") if b.strip()]
    if not baris:
        return
    matriks = [re.split(r"[\t;]", b.strip()) for b in baris]
    n_kolom = max(len(r) for r in matriks)
    _set_jumlah_kolom(tabel, n_kolom)

    trs = tabel._tbl.findall(qn("w:tr"))
    proto_tr = trs[-1]
    while len(tabel._tbl.findall(qn("w:tr"))) < len(matriks):
        tabel._tbl.append(copy.deepcopy(proto_tr))
    while len(tabel._tbl.findall(qn("w:tr"))) > len(matriks):
        tabel._tbl.remove(tabel._tbl.findall(qn("w:tr"))[-1])

    for i, isi_baris in enumerate(matriks):
        for j in range(n_kolom):
            sel = tabel.cell(i, j)
            p = sel.paragraphs[0]
            nilai = isi_baris[j] if j < len(isi_baris) else ""
            if p.runs:
                p.runs[0].text = nilai
                for r in list(p.runs[1:]):
                    r._r.getparent().remove(r._r)
            else:
                p.add_run(nilai)


# =========================================================
# HEADER, FOOTER, CATATAN KAKI
# =========================================================
def isi_header_footer(doc, n: Naskah) -> None:
    penulis_pertama = (n.penulis[0].nama if (n.penulis and not n.blind) else "[Penulis]")
    koresp = next((p for p in n.penulis if p.koresp), None)
    afil_koresp = (n.afiliasi[koresp.afil[0] - 1]
                   if (koresp and koresp.afil and not n.blind) else "")

    for s in doc.sections:
        # Header ganjil: penulis pertama (kiri) — nama jurnal (kanan)
        p = s.header.paragraphs[0]
        tulis(p, [(f"{penulis_pertama}, et al.", None), ("\t", None),
                  ("Jurnal Ilmiah PLATAX", None)])
        # Header genap: nama jurnal (kiri) — judul singkat (kanan)
        p = s.even_page_header.paragraphs[0]
        tulis(p, [("Jurnal Ilmiah PLATAX", None), ("\t", None),
                  (n.running_title, None)])

        # Catatan kaki halaman 1
        f1 = s.first_page_footer
        ps = f1.paragraphs
        idx_koresp = next((i for i, x in enumerate(ps)
                           if x.text.strip().startswith("*Penulis korespondensi")), None)
        if idx_koresp is not None:
            if n.blind:
                for k in (idx_koresp + 2, idx_koresp + 1, idx_koresp):
                    if k < len(ps):
                        hapus_paragraf(ps[k])
            else:
                isi_run_pertama(
                    ps[idx_koresp],
                    f"*Penulis korespondensi (Corresponding author): {koresp.nama if koresp else ''}")
                if idx_koresp + 1 < len(ps):
                    isi_run_pertama(ps[idx_koresp + 1], afil_koresp)
                if idx_koresp + 2 < len(ps):
                    isi_run_pertama(ps[idx_koresp + 2],
                                    f"Tel: {n.telepon}, E-mail: {n.email}")


# =========================================================
# FUNGSI UTAMA
# =========================================================
def bangun(n: Naskah, path_template: str) -> io.BytesIO:
    doc = Document(path_template)
    buang_keterangan_abu(doc)

    # ---- Judul dan judul singkat
    tulis(paragraf(doc, "[Judul penelitian dalam bahasa Indonesia"), [(n.judul_id, None)])
    tulis(paragraf(doc, "[Complete Research Title in English"), [(n.judul_en, None)])
    p = paragraf(doc, "Judul singkat (running title)")
    tulis(p, [("Judul singkat (running title): ", None), (n.running_title, None)])

    # ---- Penulis dan afiliasi
    p_pen = paragraf(doc, "[Nama Penulis 1]")
    fmt_nama = rprot(p_pen, 0)          # 10 pt bold, hitam
    fmt_sup = rprot(p_pen, 1)           # superskrip biru 007BB8
    p_afil = paragraf(doc, "1[Departemen")
    fmt_nomor = rprot(p_afil, 0)        # 8 pt italic biru superskrip
    fmt_teks = rprot(p_afil, 1)         # 8 pt italic

    if n.blind:
        tulis(p_pen, [("[Identitas penulis dan afiliasi dihapus untuk proses "
                       "double-blind review — lihat berkas Title Page terpisah]", None)],
              cetakan=fmt_nama)
        for _ in range(3):
            try:
                hapus_paragraf(paragraf(doc, "[Departemen"))
            except KeyError:
                break
    else:
        segmen = []
        for i, pen in enumerate(n.penulis):
            if i:
                segmen.append((", ", fmt_nama))
            segmen.append((pen.nama, fmt_nama))
            tanda = ",".join(str(x) for x in pen.afil)
            if pen.koresp:
                tanda += "*"
            if tanda:
                segmen.append((tanda, fmt_sup))
        tulis(p_pen, segmen, cetakan=fmt_nama)

        terpakai = sorted({x for pen in n.penulis for x in pen.afil})
        lama = [doc.paragraphs[cari(doc, "1[Departemen") + k] for k in range(3)]
        proto = lama[0]
        target, terakhir = [proto], proto
        for _ in range(max(len(terpakai) - 1, 0)):
            baru = copy.deepcopy(proto._p)
            terakhir._p.addnext(baru)
            terakhir = Paragraph(baru, proto._parent)
            target.append(terakhir)
        for x in lama[1:]:
            hapus_paragraf(x)
        for p_target, nomor in zip(target, terpakai):
            tulis(p_target, [(str(nomor), fmt_nomor),
                             (n.afiliasi[nomor - 1], fmt_teks)], cetakan=fmt_teks)

    # ---- Abstrak (kotak tabel 2 dan 3 pada template)
    for tabel, label, isi, label_kw, kw in (
            (doc.tables[2], "Abstract", n.abstract_en, "Keywords: ", n.keywords_en),
            (doc.tables[3], "Abstrak", n.abstrak_id, "Kata kunci: ", n.kata_kunci_id)):
        ps = tabel.cell(0, 0).paragraphs
        isi_run_pertama(ps[0], label)
        if len(ps) > 1:
            tulis(ps[1], [(isi.replace("\n", " ").strip(), None)])
        if len(ps) > 2:
            tulis(ps[2], [(label_kw, rprot(ps[2], 0)), (kw, rprot(ps[2], 1))])

    # ---- Badan naskah
    isi_blok(doc, "[Paragraf 1 —", n.bab1, jumlah_placeholder=3, warnai=True)
    isi_blok(doc, "[Periode penelitian", n.metode_21, warnai=True)
    isi_blok(doc, "[Alat, bahan,", n.metode_22, warnai=True)
    isi_blok(doc, "[Prosedur laboratorium", n.metode_221, warnai=True)
    isi_blok(doc, "[Formula/indeks", n.metode_23, warnai=True)
    isi_blok(doc, "[Sajikan temuan", n.bab3, warnai=True)
    isi_blok(doc, "[Interpretasi temuan", n.bab4, jumlah_placeholder=3, warnai=True)
    isi_blok(doc, "[Maksimal 150 kata", n.bab5)

    # ---- Tabel 1
    isi_run_pertama(paragraf(doc, "Tabel 1."), n.cap_tabel_id, idx=1)
    isi_run_pertama(paragraf(doc, "Table 1."), n.cap_tabel_en, idx=1)
    if n.tabel_data.strip():
        isi_tabel(doc.tables[4], n.tabel_data)
    isi_run_pertama(paragraf(doc, "Keterangan:"), n.cat_tabel)

    # ---- Gambar 1
    p_gbr = paragraf(doc, "[Sisipkan gambar di sini")
    if n.gambar_blob:
        tulis(p_gbr, [("", None)])
        p_gbr.runs[0].add_picture(io.BytesIO(n.gambar_blob), width=Cm(n.lebar_gambar))
    isi_run_pertama(paragraf(doc, "Gambar 1."), n.cap_gambar_id, idx=1)
    isi_run_pertama(paragraf(doc, "Figure 1."), n.cap_gambar_en, idx=1)

    # ---- Tujuh pernyataan akhir (urutan FAS sudah ada di template)
    isi_blok(doc, "Penulis menyatakan tidak ada konflik", n.konflik)
    isi_blok(doc, "[Tidak berlaku / Not applicable — atau: Penelitian ini didanai", n.dana)
    isi_blok(doc, "[Pihak yang memberi bantuan",
             "[Dihapus untuk proses double-blind review]" if n.blind else n.ucapan)
    isi_blok(doc, "[CRediT taxonomy", n.kontribusi)
    isi_blok(doc, "Dataset yang mendukung", n.data_avail)
    isi_blok(doc, "[Tidak berlaku / Not applicable — atau nama komisi etik", n.etik)

    # ---- ORCID
    baris_orcid = [f"{p.nama}   {p.orcid}" for p in n.penulis if p.orcid and not n.blind]
    isi_blok(doc, "[Nama Penulis 1]   https://orcid.org",
             "\n".join(baris_orcid) if baris_orcid
             else "[Tidak berlaku / Not applicable]",
             jumlah_placeholder=2)

    # ---- Daftar pustaka
    hapus_paragraf(paragraf(doc, "[APA 7th Edition"))
    isi_blok(doc, "Rondonuwu, A. B., Kepel", n.daftar_pustaka, jumlah_placeholder=2)

    isi_header_footer(doc, n)

    # ---- Metadata Word (Petunjuk §9: wajib dibersihkan untuk blind review)
    cp = doc.core_properties
    cp.author = "" if n.blind else (n.penulis[0].nama if n.penulis else "")
    cp.last_modified_by = ""
    cp.title = "" if n.blind else n.judul_id
    cp.subject = cp.comments = cp.category = cp.keywords = ""

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
