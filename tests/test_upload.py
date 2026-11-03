"""Uji dasar: baca naskah bebas-format, bangun 3 berkas, buka ulang dengan python-docx."""
import io, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import docx
import naskah_builder as B
import naskah_parser as P

SAMPLE = Path(__file__).with_name("contoh_naskah_bebas.docx").read_bytes()


def test_parse():
    M, rep = P.parse_docx(SAMPLE)
    assert len(M["authors"]) == 3 and M["authors"][0]["email"] == "budi@unsrat.ac.id"
    assert M["title_id"] and M["title_en"] and M["abs_id"] and M["abs_en"]
    assert len(M["refs"]) == 2 and rep["found"]["Sitiran dikenali"] == 2


def test_build_all():
    M, _ = P.parse_docx(SAMPLE)
    for kind in ("blind", "full", "title"):
        d = docx.Document(io.BytesIO(B.build(kind, M)))
        assert len(d.paragraphs) > 10


def test_blind_has_no_identity():
    M, _ = P.parse_docx(SAMPLE)
    d = docx.Document(io.BytesIO(B.build("blind", M)))
    txt = "\n".join(p.text for p in d.paragraphs)
    assert "Budi Santoso" not in txt and "budi@unsrat.ac.id" not in txt


def test_equations_and_title_page_placeholder():
    M, rep = P.parse_docx((Path(__file__).with_name("contoh_rumus.docx")).read_bytes())
    eqs = [b for b in M["sections"]["results"] if b["t"] == "eq"]
    assert len(eqs) == 2 and eqs[0]["num"] == "(1)"
    M["authors"] = []  # naskah blind tanpa penulis: Title Page tetap dibuat dengan penanda
    d = docx.Document(io.BytesIO(B.build("title", M)))
    assert "Penulis mengisi" in "\n".join(p.text for p in d.paragraphs)
