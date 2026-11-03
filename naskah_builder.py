"""
naskah_builder.py
Menerapkan template resmi JIP/PLATAX 2026 pada naskah yang sudah dibaca (lihat naskah_parser.py).

Berkas template asli dibuka sebagai cetakan, sehingga kop, pita biru, header ganjil-genap, footer,
dan dua kolom tetap persis seperti buatan redaksi. Hanya isi naskah yang diganti.

Isi paragraf dapat berupa:
  * str   : teks dengan penanda ringan  *miring*  **tebal**  ^atas^  ~bawah~
  * list  : daftar "run" hasil pembacaan .docx, tiap run berupa dict
            {"t": teks, "b": bool, "i": bool, "sup": bool, "sub": bool}  atau  {"math": xml_omml}
"""
from __future__ import annotations

import base64
import datetime as _dt
import io
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as _xml_escape

TEMPLATE_DIR = Path(__file__).resolve().parent / "templates"
TEMPLATE_FILES = {
    "id": "Template_Artikel_PLATAX_2026_ID.docx",
    "en": "Template_Article_PLATAX_2026_EN.docx",
    "bid": "Template_Naskah_BlindReview_JIP_2026.docx",
    "ben": "Template_Manuscript_BlindReview_JIP_2026.docx",
    "tp": "Template_Title_Page_JIP_2026.docx",
}

BLUE = "007BB8"
CREDIT = [
    "Conceptualization", "Data curation", "Formal analysis", "Funding acquisition", "Investigation",
    "Methodology", "Project administration", "Resources", "Software", "Supervision", "Validation",
    "Visualization", "Writing – original draft", "Writing – review & editing",
]
SECTION_KEYS = ["intro", "methods", "results", "disc", "concl"]

LBL = {
    "id": {
        "sec": ["Pendahuluan", "Bahan dan Metode", "Hasil", "Pembahasan", "Simpulan"], "refs": "Daftar Pustaka",
        "coi": "Konflik kepentingan (Competing interests)", "fund": "Sumber dana (Funding sources)",
        "ack": "Ucapan terima kasih (Acknowledgements)", "contrib": "Kontribusi penulis (Authors’ contributions)",
        "data": "Ketersediaan data (Availability of data and materials)",
        "ethics": "Persetujuan etik (Ethics approval and consent to participate)", "orcid": "ORCID",
        "run": "Judul singkat (running title): ", "table": "Tabel", "figure": "Gambar", "note": "Keterangan: ", "author": "Penulis",
        "defCoi": "Penulis menyatakan tidak ada konflik kepentingan yang relevan dengan artikel ini.",
        "defData": "Dataset yang mendukung hasil penelitian ini tersedia dari penulis korespondensi berdasarkan permintaan yang wajar.",
        "defFund": "Tidak berlaku / Not applicable", "defAck": "Tidak ada / None", "defEthics": "Tidak berlaku / Not applicable",
        "blindAck": "Dihapus untuk blind review; dicantumkan pada Title Page.", "noOrcid": "Tidak tersedia / Not available",
        "tel": "Tel (opsional): ",
    },
    "en": {
        "sec": ["Introduction", "Materials and Methods", "Results", "Discussion", "Conclusion"], "refs": "References",
        "coi": "Competing interests", "fund": "Funding sources", "ack": "Acknowledgements", "contrib": "Authors’ contributions",
        "data": "Availability of data and materials", "ethics": "Ethics approval and consent to participate", "orcid": "ORCID",
        "run": "Running title: ", "table": "Table", "figure": "Figure", "note": "Note: ", "author": "Author",
        "defCoi": "The authors declare no competing interests relevant to this article.",
        "defData": "The dataset supporting the findings of this study is available from the corresponding author upon reasonable request.",
        "defFund": "Not applicable", "defAck": "None", "defEthics": "Not applicable",
        "blindAck": "Removed for blind review; provided on the Title Page.", "noOrcid": "Not available",
        "tel": "Tel (optional): ",
    },
}

COLW = 5100  # lebar satu kolom (twips)
UP = "A-ZÀ-ÖØ-Þ"  # huruf kapital (termasuk beraksen)


# ----------------------------------------------------------------------------
# utilitas teks
# ----------------------------------------------------------------------------
_BAD = re.compile("[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]")


def clean(s) -> str:
    return _BAD.sub("", "" if s is None else str(s))


def esc(s) -> str:
    return _xml_escape(clean(s), {'"': "&quot;"})


def _strip_acc(s: str) -> str:
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", s) if not unicodedata.combining(c))


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]+", " ", _strip_acc(str(s)).lower())).strip()


# --- penanda ringan -> segmen berformat ------------------------------------
def parse_inline(s: str, fmt: dict | None = None) -> list[dict]:
    fmt = dict(fmt or {})
    out: list[dict] = []
    buf = ""
    i = 0

    def flush():
        nonlocal buf
        if buf:
            d = dict(fmt)
            d["t"] = buf
            out.append(d)
            buf = ""

    n = len(s)
    while i < n:
        c = s[i]
        if c == "\\" and i + 1 < n and s[i + 1] in "*^~\\":
            buf += s[i + 1]
            i += 2
            continue
        mk = None
        if s.startswith("**", i):
            mk = "**"
        elif c in "*^~":
            mk = c
        if mk and i + len(mk) < n and not s[i + len(mk)].isspace():
            j = i + len(mk)
            found = -1
            while j < n:
                if s[j] == "\\":
                    j += 2
                    continue
                if s.startswith(mk, j) and not s[j - 1].isspace() and not (mk == "*" and s[j + 1:j + 2] == "*"):
                    found = j
                    break
                j += 1
            if found > i + len(mk):
                flush()
                f = dict(fmt)
                if mk == "**":
                    f["b"] = True
                elif mk == "*":
                    f["i"] = not fmt.get("i")
                elif mk == "^":
                    f["sup"], f["sub"] = True, False
                else:
                    f["sub"], f["sup"] = True, False
                out.extend(parse_inline(s[i + len(mk):found], f))
                i = found + len(mk)
                continue
        buf += c
        i += 1
    flush()
    return out


def strip_markup(s) -> str:
    return "".join(x["t"] for x in parse_inline(clean(s))) if isinstance(s, str) else plain_text(s)


def plain_text(content) -> str:
    if isinstance(content, str):
        return strip_markup(content)
    return "".join(r.get("t", "") for r in (content or []) if "math" not in r)


def word_count(content) -> int:
    t = plain_text(content).strip()
    return len(t.split()) if t else 0


# ----------------------------------------------------------------------------
# run / paragraf
# ----------------------------------------------------------------------------
def rpr(f: dict) -> str:
    s = ""
    if f.get("b"):
        s += "<w:b/><w:bCs/>"
    if f.get("i") is True:
        s += "<w:i/><w:iCs/>"
    elif f.get("i") is False and f.get("_off"):
        s += '<w:i w:val="0"/><w:iCs w:val="0"/>'
    if f.get("color"):
        s += f'<w:color w:val="{f["color"]}"/>'
    if f.get("sz"):
        s += f'<w:sz w:val="{f["sz"]}"/><w:szCs w:val="{f["sz"]}"/>'
    if f.get("sup"):
        s += '<w:vertAlign w:val="superscript"/>'
    elif f.get("sub"):
        s += '<w:vertAlign w:val="subscript"/>'
    return f"<w:rPr>{s}</w:rPr>" if s else ""


def run(t: str, f: dict | None = None) -> str:
    return "" if t == "" else f'<w:r>{rpr(f or {})}<w:t xml:space="preserve">{esc(t)}</w:t></w:r>'


def para(inner: str = "", **p) -> str:
    s = f'<w:pStyle w:val="{p.get("style", "Normal")}"/>'
    if p.get("keepNext"):
        s += "<w:keepNext/>"
    if p.get("shd"):
        s += f'<w:shd w:val="clear" w:color="auto" w:fill="{p["shd"]}"/>'
    if p.get("tabs"):
        s += p["tabs"]
    if p.get("before") is not None or p.get("after") is not None:
        s += f'<w:spacing w:before="{p.get("before") or 0}" w:after="{p.get("after") or 0}"/>'
    if any(p.get(k) is not None for k in ("left", "right", "hanging", "firstLine")):
        s += "<w:ind"
        for k in ("left", "right", "hanging", "firstLine"):
            if p.get(k) is not None:
                s += f' w:{k}="{p[k]}"'
        s += "/>"
    if p.get("jc"):
        s += f'<w:jc w:val="{p["jc"]}"/>'
    return f"<w:p><w:pPr>{s}</w:pPr>{inner}</w:p>"


# ----------------------------------------------------------------------------
# sitiran dalam teks
# ----------------------------------------------------------------------------
_NAME_CORE = rf"[{UP}][^\s,;()]*(?:\s+(?:et\s+al\.|&|dan|and|[{UP}][^\s,;()]*))*"
_NAME_NAR = rf"[{UP}][^\s,;()]*(?:\s+(?:et\s+al\.|(?:&|dan|and)\s+[{UP}][^\s,;()]*))?"
_YEAR = r"(?:19|20)\d{2}[a-z]?"
_SEG_RE = re.compile(
    rf"^(?:(?:e\.g\.|mis\.|lihat|see|cf\.),?\s+)?({_NAME_CORE}),?\s+({_YEAR}(?:\s*,\s*{_YEAR})*)(?:,\s*(?:p|pp|hlm)\.?\s*[\d–-]+)?$")
_NAR_RE = re.compile(rf"(?<![^\W\d_])({_NAME_NAR})\s\(({_YEAR}(?:\s*[,;]\s*{_YEAR})*)\)")


def build_ref_index(refs: list) -> list[dict]:
    return [{"plain": norm(plain_text(r)), "raw": plain_text(r), "bm": f"ref_{i + 1}", "n": i} for i, r in enumerate(refs)]


def find_ref(ref_idx, name: str, years: str):
    if not ref_idx:
        return None
    nm = norm(re.sub(r"et\s+al\.", " ", name, flags=re.I))
    parts = [w for w in nm.split(" ") if w and w not in ("dan", "and")]
    if not parts:
        return None
    first = parts[0]
    second = parts[1] if len(parts) > 1 else None
    for y in re.findall(r"(?:19|20)\d{2}[a-z]?", years):
        pat = re.compile(r"\(" + re.escape(y) + r"\)|\b" + re.escape(y) + r"\b", re.I)
        hits = [r for r in ref_idx if r["plain"].startswith(first) and y[:4] in r["plain"] and pat.search(r["raw"])]
        if hits:
            if second:
                for r in hits:
                    if second in r["plain"][:160]:
                        return r
            return hits[0]
    return None


def cite_spans(plain: str, ref_idx) -> list[dict]:
    spans: list[dict] = []
    for m in re.finditer(r"\(([^()]*)\)", plain):
        content = m.group(1)
        if not re.search(r"(?:19|20)\d{2}", content):
            continue
        pos = m.start() + 1
        segs = content.split(";")
        flags = []
        for k, raw in enumerate(segs):
            trimmed = raw.strip()
            lead = len(raw) - len(raw.lstrip())
            mm = _SEG_RE.match(trimmed)
            f = {"ok": False}
            if mm:
                s = pos + lead + mm.start(1)
                e = pos + lead + mm.end(2)
                r = find_ref(ref_idx, mm.group(1), mm.group(2))
                spans.append({"s": s, "e": e, "anchor": r["bm"] if r else None, "label": f"{mm.group(1)}, {mm.group(2)}", "ref": r})
                f["ok"] = True
            if k < len(segs) - 1:
                f["semi"] = pos + len(raw)
            flags.append(f)
            pos += len(raw) + 1
        for k, f in enumerate(flags):
            if f["ok"] and "semi" in f and k + 1 < len(flags) and flags[k + 1]["ok"]:
                spans.append({"s": f["semi"], "e": f["semi"] + 1, "anchor": None, "semi": True})
    for m in _NAR_RE.finditer(plain):
        r = find_ref(ref_idx, m.group(1), m.group(2))
        a = r["bm"] if r else None
        spans.append({"s": m.start(1), "e": m.end(1), "anchor": a, "label": f"{m.group(1)} ({m.group(2)})", "ref": r})
        spans.append({"s": m.start(2), "e": m.end(2), "anchor": a, "no_count": True})
    return spans


def doi_spans(plain: str) -> list[dict]:
    out = []
    for m in re.finditer(r"https?://(?:dx\.)?doi\.org/[^\s<>]+", plain):
        t = re.sub(r"[.,;)\]]+$", "", m.group(0))
        out.append({"s": m.start(), "e": m.start() + len(t), "anchor": None})
    return out


def _resolve(spans):
    spans = sorted(spans, key=lambda x: (x["s"], -x["e"]))
    out, end = [], -1
    for sp in spans:
        if sp["s"] >= end:
            out.append(sp)
            end = sp["e"]
    return out


def _to_segs(content, base: dict) -> list[dict]:
    """Ubah isi (str atau daftar run) menjadi segmen berformat lengkap."""
    if isinstance(content, str):
        segs = parse_inline(clean(content), dict(base))
        for sg in segs:
            if sg.get("i") is False:
                sg["_off"] = True
        return segs
    segs = []
    for r in content or []:
        if "math" in r:
            segs.append({"math": r["math"], "t": ""})
            continue
        t = clean(r.get("t", ""))
        if not t:
            continue
        f = dict(base)
        if r.get("b") and not base.get("b"):
            f["b"] = True
        if r.get("i"):
            f["i"] = (not base.get("i")) if base.get("i") else True
            if base.get("i"):
                f["_off"] = True
        if r.get("sup"):
            f["sup"] = True
        if r.get("sub"):
            f["sub"] = True
        f["t"] = t
        segs.append(f)
    return segs


def inline(content, base: dict | None = None, cites: bool = False, ref_idx=None, doi: bool = False) -> str:
    base = base or {}
    segs = _to_segs(content, base)
    plain = "".join(sg["t"] for sg in segs)
    spans: list[dict] = []
    if cites:
        spans += cite_spans(plain, ref_idx or [])
    if doi:
        spans += doi_spans(plain)
    spans = _resolve(spans)
    xml = ""
    off = 0
    open_link = None
    link_buf = ""

    def close_link():
        nonlocal xml, open_link, link_buf
        if open_link:
            xml += f'<w:hyperlink w:anchor="{open_link}" w:history="1">{link_buf}</w:hyperlink>'
            open_link, link_buf = None, ""

    for sg in segs:
        if "math" in sg:
            close_link()
            xml += sg["math"]
            continue
        a, b = off, off + len(sg["t"])
        off = b
        cur = a
        pieces = []
        for sp in spans:
            if sp["e"] <= a or sp["s"] >= b:
                continue
            ps, pe = max(sp["s"], a), min(sp["e"], b)
            if ps > cur:
                pieces.append((cur, ps, None))
            pieces.append((ps, pe, sp))
            cur = pe
        if cur < b:
            pieces.append((cur, b, None))
        for ps, pe, sp in pieces:
            f = {k: v for k, v in sg.items() if k != "t"}
            if sp:
                f["color"] = BLUE
            rx = run(sg["t"][ps - a:pe - a], f)
            anc = sp["anchor"] if sp and sp.get("anchor") else None
            if anc:
                if open_link and open_link != anc:
                    close_link()
                open_link = anc
                link_buf += rx
            else:
                close_link()
                xml += rx
    close_link()
    return xml


# ----------------------------------------------------------------------------
# kata kunci, penulis, pustaka
# ----------------------------------------------------------------------------
def keywords(s, sort=True) -> list[str]:
    k = plain_text(s) if not isinstance(s, str) else clean(s)
    arr = re.split(r";|\n", k) if re.search(r";|\n", k) else k.split(",")
    arr = [re.sub(r"\.$", "", x.strip()) for x in arr]
    arr = [x for x in arr if x]
    if sort:
        arr.sort(key=lambda x: _strip_acc(x).lower())
    return arr


def surname(name: str) -> str:
    w = [x for x in re.sub(r"[,.]", " ", name or "").split() if x]
    return w[-1] if w else ""


def header_name(M: dict) -> str:
    A = [a for a in M["authors"] if (a.get("name") or "").strip()]
    if not A:
        return ""
    hn = (M.get("header_name") or "").strip()
    if len(A) == 1:
        return hn or surname(A[0]["name"])
    if len(A) == 2 and not hn:
        return f"{surname(A[0]['name'])} & {surname(A[1]['name'])}"
    return f"{hn or surname(A[0]['name'])} et al."


def orcid_url(v) -> str:
    v = (v or "").strip()
    if not v:
        return ""
    m = re.search(r"\d{4}-\d{4}-\d{4}-\d{3}[\dXx]", v)
    return "https://orcid.org/" + m.group(0).upper() if m else v


def get_refs(M: dict) -> list:
    refs = [r for r in (M.get("refs") or []) if plain_text(r).strip()]
    if M.get("sort_refs", True):
        refs = sorted(refs, key=lambda r: _strip_acc(plain_text(r)).lower())
    return refs


def paras_of(text) -> list[str]:
    return [x.strip() for x in re.split(r"\n\s*\n|\r?\n", clean(text)) if x.strip()]


def contrib_line(a: dict, i: int, blind: bool, L: dict):
    roles = [r for r in a.get("roles", []) if r in CREDIT]
    if not roles:
        return None
    nm = f"{L['author']} {i + 1}" if blind else a["name"].strip()
    return f"{nm}: {', '.join(roles)}"


# ----------------------------------------------------------------------------
# tabel dan gambar
# ----------------------------------------------------------------------------
H_TABS = '<w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="left" w:pos="284" w:leader="none"/></w:tabs>'


def _norm_rows(tbl: dict):
    rows = []
    for r in tbl.get("rows", []):
        cells = []
        for c in r:
            if isinstance(c, dict) and "c" in c:
                cells.append({"c": c["c"], "span": int(c.get("span") or 1), "vm": c.get("vm")})
            else:
                cells.append({"c": c, "span": 1, "vm": None})
        rows.append(cells)
    rows = [r for r in rows if any(plain_text(c["c"]).strip() or c["vm"] == "cont" for c in r)]
    return rows


def table_xml(tbl: dict, full: bool) -> str:
    rows = _norm_rows(tbl)
    if not rows:
        return ""
    ncols = max(sum(c["span"] for c in r) for r in rows)
    total = 10460 if full else 4309
    wts = [5] * ncols
    for r in rows:
        col = 0
        for c in r:
            if c["span"] == 1:
                wts[col] = max(wts[col], min(40, len(plain_text(c["c"]))))
            col += c["span"]
    sm = sum(wts)
    ws = [max(500, round(total * w / sm)) for w in wts]
    ws[-1] += total - sum(ws)

    def bd(t, b):
        return "<w:tcBorders>" + (f'<w:top w:val="single" w:sz="{t}" w:space="0" w:color="000000"/>' if t else "") + \
               (f'<w:bottom w:val="single" w:sz="{b}" w:space="0" w:color="000000"/>' if b else "") + "</w:tcBorders>"

    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
         '<w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar>'
         '<w:tblLook w:val="04a0" w:noHBand="0" w:noVBand="1" w:firstColumn="1" w:lastRow="0" w:lastColumn="0" w:firstRow="1"/></w:tblPr>'
         "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in ws) + "</w:tblGrid>")
    nr = len(rows)
    hdr = 1
    while hdr < nr and hdr < 4 and any(c["vm"] == "cont" for c in rows[hdr]) and any(c["vm"] == "restart" for c in rows[0]):
        hdr += 1
    if nr > 1 and hdr >= nr:
        hdr = 1
    for ri, r in enumerate(rows):
        top = 4 if ri == 0 else 0
        bot = 4 if ri == nr - 1 else (6 if ri == hdr - 1 else 0)
        x += "<w:tr>" + ("<w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>" if ri < hdr else "<w:trPr><w:cantSplit/></w:trPr>")
        col = 0
        for c in r:
            w = sum(ws[col:col + c["span"]])
            span = f'<w:gridSpan w:val="{c["span"]}"/>' if c["span"] > 1 else ""
            vm = ""
            if c["vm"] == "restart":
                vm = '<w:vMerge w:val="restart"/>'
            elif c["vm"] == "cont":
                vm = "<w:vMerge/>"
            x += f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{span}{vm}{bd(top, bot)}<w:vAlign w:val="center"/></w:tcPr>'
            base = {"sz": 15, "b": True} if ri < hdr else {"sz": 15}
            first_col = col == 0
            x += para(inline(c["c"], base), before=0, after=0, jc="left" if first_col and ri >= hdr else "center") + "</w:tc>"
            col += c["span"]
        x += "</w:tr>"
    return x + "</w:tbl>"


TARGET_DPI = 600
MAX_PIXELS = 60_000_000


def prep_image(im: dict, width_emu: int):
    """Naikkan resolusi ke TARGET_DPI pada ukuran cetak (Lanczos) dan tulis metadata dpi. Mengembalikan (b64, ext, w, h)."""
    try:
        from PIL import Image
        raw = base64.b64decode(im["b64"])
        with Image.open(io.BytesIO(raw)) as src:
            src.load()
            w, h = src.size
            width_in = width_emu / 914400
            want_w = round(width_in * TARGET_DPI)
            ext = (im.get("ext") or "png").lower()
            jpeg = ext in ("jpeg", "jpg")
            if want_w > w and want_w * round(h * want_w / w) <= MAX_PIXELS:
                nh = round(h * want_w / w)
                img = src.convert("RGBA" if (src.mode in ("RGBA", "LA", "P") and not jpeg) else "RGB").resize((want_w, nh), Image.LANCZOS)
                w, h = want_w, nh
            else:
                img = src.convert("RGBA" if (src.mode in ("RGBA", "LA", "P") and not jpeg) else "RGB") if src.mode not in ("RGB", "RGBA", "L") else src.copy()
            out = io.BytesIO()
            if jpeg:
                img.convert("RGB").save(out, "JPEG", quality=95, dpi=(TARGET_DPI, TARGET_DPI), subsampling=0)
                ext = "jpeg"
            else:
                img.save(out, "PNG", dpi=(TARGET_DPI, TARGET_DPI), optimize=True)
                ext = "png"
            return base64.b64encode(out.getvalue()).decode(), ext, w, h
    except Exception:
        return im["b64"], im.get("ext", "png"), im.get("w"), im.get("h")


def drawing_xml(ctx: dict, imgs: list, caption_alt: str, full: bool) -> str:
    out = ""
    maxw = 6480000 if full else 2736000
    maxh = 7200000
    n = max(1, len(imgs))
    for im in imgs:
        ctx["img_n"] += 1
        k = ctx["img_n"]
        w, h = im.get("w") or 800, im.get("h") or 600
        wmax = maxw // n if n > 1 else maxw
        cx, cy = wmax, round(wmax * h / w)
        if cy > maxh:
            cy = maxh
            cx = round(maxh * w / h)
        b64, ext, pw, ph = prep_image(im, cx)
        fname = f"media/fig{k}.{'jpg' if ext == 'jpeg' else ext}"
        rid = f"rIdFig{k}"
        ctx["media"].append({"rid": rid, "name": fname, "ext": ext, "b64": b64})
        did = 100 + k
        out += (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
                f'<wp:docPr id="{did}" name="Gambar {k}" descr="{esc(caption_alt)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
                '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:nvPicPr><pic:cNvPr id="{did}" name="Gambar {k}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
                f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                f'<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    return out


def sect_break(cols: int, parts: dict) -> str:
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:sectPr><w:type w:val="continuous"/>'
            f'{parts["pgSz"]}{parts["pgMar"]}<w:cols w:num="{cols}" w:space="244" w:equalWidth="true"/>{parts["grid"]}</w:sectPr></w:pPr></w:p>')


# ----------------------------------------------------------------------------
# isi naskah
# ----------------------------------------------------------------------------
def body_xml(M: dict, L: dict, blind: bool, ctx: dict) -> str:
    refs = get_refs(M)
    ref_idx = build_ref_index(refs)
    cite = M.get("color_cites", True)
    link_idx = ref_idx if M.get("link_cites", True) else []
    x = ""
    ctx["nT"] = ctx["nF"] = 0

    def wrap_full(inner):
        ctx["full_used"] = True
        return sect_break(2, ctx["parts"]) + inner + sect_break(1, ctx["parts"])

    merged = bool(M.get("merged_rd"))
    num = 0
    for si, k in enumerate(SECTION_KEYS):
        if merged and k == "disc":
            continue
        num += 1
        title = ("Hasil dan Pembahasan" if L is LBL["id"] else "Results and Discussion") if (merged and k == "results") else L["sec"][si]
        x += para(run(f"{num}. ") + "<w:r><w:tab/></w:r>" + inline(title),
                  style="Heading1", keepNext=True, tabs=H_TABS, before=120, after=100)
        h2 = h3 = 0
        blocks = list((M.get("sections") or {}).get(k, []))
        if k == "methods" and M.get("ai_use") == "yes" and (M.get("ai_text") or "").strip():
            blocks.append({"t": "p", "c": M["ai_text"]})
        for b in blocks:
            t = b.get("t")
            if t == "p":
                c = b.get("c", b.get("text", ""))
                items = [c] if not isinstance(c, str) else paras_of(c)
                for p in items:
                    if isinstance(p, list) and not any(plain_text([r]).strip() or "math" in r for r in p):
                        continue
                    if b.get("center"):
                        x += para(inline(p, cites=cite, ref_idx=link_idx), before=40, after=80, jc="center")
                    else:
                        x += para(inline(p, cites=cite, ref_idx=link_idx), before=0, after=60, firstLine=0 if b.get("noindent") else 283)
            elif t == "eq":
                ms = "".join(b.get("m") or [])
                if not ms:
                    continue
                if b.get("num"):
                    tabs = f'<w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="center" w:pos="{COLW // 2}"/><w:tab w:val="right" w:pos="{COLW}"/></w:tabs>'
                    x += para("<w:r><w:tab/></w:r>" + ms + "<w:r><w:tab/></w:r>" + run(b["num"]), before=60, after=100, tabs=tabs, jc="left")
                else:
                    x += para("<m:oMathPara>" + ms + "</m:oMathPara>", before=60, after=100, jc="center")
            elif t in ("h2", "h3"):
                txt = plain_text(b.get("c", b.get("text", ""))).strip()
                if not txt:
                    continue
                if t == "h2" or h2 == 0:
                    h2 += 1
                    h3 = 0
                    label, style = f"{num}.{h2}. ", "Heading2"
                else:
                    h3 += 1
                    label, style = f"{num}.{h2}.{h3}. ", "Heading3"
                x += para(run(label) + inline(txt, {"i": True}), style=style, keepNext=True, tabs=H_TABS)
            elif t == "table":
                tx = table_xml(b, bool(b.get("full")))
                if not tx:
                    continue
                ctx["nT"] += 1
                cap = para(run(f"{L['table']} {ctx['nT']}. ", {"b": True, "sz": 15}) + inline(b.get("cap", ""), {"sz": 15}, cites=cite, ref_idx=link_idx),
                           keepNext=True, before=140, after=50, left=567, hanging=567, jc="left")
                note_txt = b.get("note", "")
                note = (para(run(L["note"], {"i": True, "sz": 15}) + inline(note_txt, {"i": True, "sz": 15}), before=50, after=160)
                        if plain_text(note_txt).strip() else para("", before=0, after=120))
                inner = cap + tx + note
                x += wrap_full(inner) if b.get("full") else inner
            elif t == "fig":
                imgs = b.get("imgs") or []
                if not imgs:
                    continue
                ctx["nF"] += 1
                d = drawing_xml(ctx, imgs, plain_text(b.get("cap", "")), bool(b.get("full")))
                img = para(d, keepNext=True, before=120, after=40, jc="center")
                cap = para(run(f"{L['figure']} {ctx['nF']}. ", {"b": True, "sz": 15}) + inline(b.get("cap", ""), {"sz": 15}, cites=cite, ref_idx=link_idx),
                           before=50, after=160, left=709, hanging=709, jc="left")
                inner = img + cap
                x += wrap_full(inner) if b.get("full") else inner

    cl = M.get("closing") or {}

    def h2p(t):
        return para(inline(t), style="Heading2", keepNext=True)

    def txt(v, default):
        v = v if (isinstance(v, list) or (v or "").strip()) else default
        items = paras_of(v) if isinstance(v, str) else [v]
        return "".join(para(inline(p)) for p in items)

    x += h2p(L["coi"]) + txt(cl.get("coi"), L["defCoi"])
    x += h2p(L["fund"]) + txt(cl.get("fund"), L["defFund"])
    x += h2p(L["ack"]) + (para(run(L["blindAck"])) if blind else txt(cl.get("ack"), L["defAck"]))
    lines = []
    for i, a in enumerate(M["authors"]):
        if (a.get("name") or "").strip() or blind:
            ln = contrib_line(a, i, blind, L)
            if ln:
                lines.append(ln)
    x += h2p(L["contrib"]) + (para(inline("; ".join(lines) + ".")) if lines else para(run("—")))
    x += h2p(L["data"]) + txt(cl.get("data"), L["defData"])
    x += h2p(L["ethics"]) + txt(cl.get("ethics"), L["defEthics"])
    if not blind:
        x += h2p(L["orcid"])
        os_ = [a for a in M["authors"] if (a.get("name") or "").strip() and orcid_url(a.get("orcid"))]
        if os_:
            for a in os_:
                x += para(run(a["name"].strip() + " :   ") + run(orcid_url(a["orcid"]), {"i": True, "color": BLUE}), before=0, after=20)
        else:
            x += para(run(L["noOrcid"]))
    x += para(run(L["refs"]), style="Heading1", keepNext=True)
    if refs:
        for i, r in enumerate(refs):
            x += (f'<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="0" w:after="50"/><w:ind w:left="397" w:hanging="397"/><w:jc w:val="left"/></w:pPr>'
                  f'<w:bookmarkStart w:id="{1000 + i}" w:name="ref_{i + 1}"/>{inline(r, {"sz": 15}, doi=True)}<w:bookmarkEnd w:id="{1000 + i}"/></w:p>')
    else:
        x += para(run("—"))
    return x


# ----------------------------------------------------------------------------
# halaman depan
# ----------------------------------------------------------------------------
def front_xml(M: dict, L: dict, blind: bool) -> str:
    x = ""
    ti, te = M.get("title_id") or "", M.get("title_en") or ""
    if plain_text(ti).strip():
        x += para(inline(ti, {"b": True, "color": BLUE, "sz": 30}), before=0, after=60, jc="left")
    if plain_text(te).strip():
        x += para(inline(te, {"b": True, "i": True, "sz": 20}), jc="left")
    x += para(run(L["run"], {"b": True}) + inline((M.get("running") or "").strip()), before=0, after=140, jc="left")
    if not blind:
        A = [(i, a) for i, a in enumerate(M["authors"]) if (a.get("name") or "").strip()]
        if A:
            inner = ""
            for k, (i, a) in enumerate(A):
                if k:
                    inner += run(", ", {"b": True, "sz": 20})
                inner += run(a["name"].strip(), {"b": True, "sz": 20})
                sup = ",".join(str(n) for n in sorted(a.get("affs") or [])) + ("*" if M.get("corr", 0) == i else "")
                if sup:
                    inner += run(sup, {"b": True, "color": BLUE, "sup": True})
            x += para(inner, before=0, after=60, jc="left")
        affs = [(n + 1, (t or "").strip()) for n, t in enumerate(M.get("affs") or []) if (t or "").strip()]
        for k, (n, t) in enumerate(affs):
            x += para(run(str(n), {"i": True, "color": BLUE, "sz": 16, "sup": True}) + inline(t, {"i": True, "sz": 16}),
                      before=0, after=200 if k == len(affs) - 1 else 20, jc="left")
    return x


def abstract_table(M: dict) -> str:
    sort = M.get("sort_kw", True)
    kw_en = "; ".join(keywords(M.get("kw_en") or "", sort))
    kw_id = "; ".join(keywords(M.get("kw_id") or "", sort))

    def flat(t):
        if isinstance(t, str):
            return re.sub(r"\s*\n+\s*", " ", clean(t)).strip()
        return t

    def cell(top, bot, label, abs_, kwl, kw):
        return (f'<w:tc><w:tcPr><w:tcW w:w="10450" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="{top}" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="{bot}" w:space="0" w:color="000000"/></w:tcBorders></w:tcPr>'
                + para(run(label, {"b": True, "sz": 16}), before=0, after=60)
                + para(inline(flat(abs_), {"sz": 16}))
                + para(run(kwl, {"b": True, "sz": 16}) + inline(kw, {"sz": 16}), before=0, after=80, jc="both") + "</w:tc>")

    return ('<w:tbl><w:tblPr><w:tblW w:w="10450" w:type="dxa"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="140" w:type="dxa"/><w:left w:w="180" w:type="dxa"/><w:bottom w:w="140" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar>'
            '<w:tblLook w:val="04a0" w:noHBand="0" w:noVBand="1" w:firstColumn="1" w:lastRow="0" w:lastColumn="0" w:firstRow="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="10450"/></w:tblGrid>'
            f'<w:tr><w:trPr></w:trPr>{cell(8, 4, "Abstract", M.get("abs_en") or "", "Keywords: ", kw_en)}</w:tr>'
            f'<w:tr><w:trPr></w:trPr>{cell(4, 8, "Abstrak", M.get("abs_id") or "", "Kata kunci: ", kw_id)}</w:tr></w:tbl>')


# ----------------------------------------------------------------------------
# paket .docx
# ----------------------------------------------------------------------------
def _open_template(key: str):
    path = TEMPLATE_DIR / TEMPLATE_FILES[key]
    zin = zipfile.ZipFile(path)
    return {i.filename: zin.read(i.filename) for i in zin.infolist()}, [i.filename for i in zin.infolist()]


def _write_package(files: dict, order: list) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zo:
        for name in order:
            if name in files:
                zo.writestr(name, files[name])
        for name, data in files.items():
            if name not in order:
                zo.writestr(name, data)
    return bio.getvalue()


def _finish_meta(files: dict, title: str):
    try:
        c = files["docProps/core.xml"].decode("utf8")
        now = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        c = re.sub(r"<dc:title>.*?</dc:title>", lambda m: f"<dc:title>{esc(plain_text(title))}</dc:title>", c, flags=re.S)
        c = re.sub(r"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)", lambda m: m.group(1) + now + m.group(2), c, flags=re.S)
        c = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>", c, flags=re.S)
        files["docProps/core.xml"] = c.encode("utf8")
    except Exception:
        pass


def _ensure_math_ns(doc: str) -> str:
    if "xmlns:m=" in doc[:3000]:
        return doc
    return doc.replace("<w:document ", '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ', 1)


def build_manuscript(M: dict, kind: str) -> bytes:
    """kind: 'full' (dengan identitas) atau 'blind' (anonim)."""
    blind = kind == "blind"
    lang = "en" if M.get("lang") == "en" else "id"
    L = LBL[lang]
    key = ("ben" if lang == "en" else "bid") if blind else lang
    files, order = _open_template(key)
    d = files["word/document.xml"].decode("utf8")
    body_open = d.index("<w:body>") + 8
    vi = d.index("PLATAX Vol.")
    m_start = d.rindex("<w:p>", 0, vi)
    ei = d.index("e-ISSN", vi)
    m_end = d.index("</w:tbl>", ei) + 8
    sb = re.search(r"<w:p><w:pPr><w:sectPr>.*?</w:sectPr></w:pPr></w:p>", d[m_end:], re.S)
    final_idx = d.rindex("<w:sectPr")
    final_sect = d[final_idx:d.index("</w:body>")]
    parts = {
        "pgSz": (re.search(r"<w:pgSz[^>]*/>", final_sect) or [""])[0] if re.search(r"<w:pgSz[^>]*/>", final_sect) else "",
        "pgMar": re.search(r"<w:pgMar[^>]*/>", final_sect).group(0) if re.search(r"<w:pgMar[^>]*/>", final_sect) else "",
        "grid": re.search(r"<w:docGrid[^>]*/>", final_sect).group(0) if re.search(r"<w:docGrid[^>]*/>", final_sect) else "",
    }
    ctx = {"media": [], "img_n": 0, "parts": parts, "full_used": False}
    body = body_xml(M, L, blind, ctx)
    if ctx["full_used"]:
        final_sect = final_sect.replace("<w:titlePg/>", "")
    out = d[:body_open] + d[m_start:m_end] + front_xml(M, L, blind) + abstract_table(M) + sb.group(0) + body + final_sect + "</w:body></w:document>"
    files["word/document.xml"] = _ensure_math_ns(out).encode("utf8")
    if ctx["media"]:
        rels = files["word/_rels/document.xml.rels"].decode("utf8")
        ct = files["[Content_Types].xml"].decode("utf8")
        for m in ctx["media"]:
            files["word/" + m["name"]] = base64.b64decode(m["b64"])
            rels = rels.replace("</Relationships>", f'<Relationship Id="{m["rid"]}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{m["name"]}"/></Relationships>')
            e = m["name"].rsplit(".", 1)[1]
            mime = {"jpg": "image/jpeg", "png": "image/png", "gif": "image/gif", "bmp": "image/bmp", "tiff": "image/tiff"}.get(e, "image/" + e)
            if not re.search(rf'Extension="{e}"', ct, re.I):
                ct = ct.replace("<Default ", f'<Default Extension="{e}" ContentType="{mime}"/><Default ', 1)
        files["word/_rels/document.xml.rels"] = rels.encode("utf8")
        files["[Content_Types].xml"] = ct.encode("utf8")

    def rep(name, fn):
        p = "word/" + name
        if p in files:
            files[p] = fn(files[p].decode("utf8")).encode("utf8")

    running = (M.get("running") or "").strip()
    rep("header1.xml", lambda s: re.sub(r"\[Judul singkat, maks\. 60 karakter\]|\[Short title, max\. 60 characters\]", lambda m: esc(running), s))
    if not blind:
        hn = header_name(M)
        rep("header2.xml", lambda s: re.sub(r"\[[^\]]*\](?: et al\.)?(?=</w:t>)", lambda m: esc(hn), s))
        ca = M["authors"][M.get("corr", 0)] if M["authors"] else {}
        aff_idx = ((ca.get("affs") or [1])[0] or 1) - 1
        affs = M.get("affs") or []
        aff = (affs[aff_idx] if 0 <= aff_idx < len(affs) else "").strip()
        tel = (M.get("phone") or "").strip()
        email = (ca.get("email") or "").strip() or "—"

        def foot(s):
            s = re.sub(r"\[(?:Nama lengkap penulis korespondensi|Full name of corresponding author)\]", lambda m: esc((ca.get("name") or "").strip() or "—"), s)
            s = re.sub(r"\[(?:Departemen, Fakultas, Universitas, Kota, Kode Pos, Negara|Department, Faculty, University, City, ZIP Code, Country)\]", lambda m: esc(aff or "—"), s)
            s = re.sub(r"Tel \((?:opsional|optional)\): \[\+62-XXXXXXXXXXX\], E-mail: \[[^\]]*\]", lambda m: esc((L["tel"] + tel + ", " if tel else "") + "E-mail: " + email), s)
            return s
        rep("footer3.xml", foot)
    _finish_meta(files, M.get("title_en") or M.get("title_id") or "")
    return _write_package(files, order)


def ph(t: str = "Penulis mengisi") -> str:
    return run(f"[{t}]", {"i": True, "color": "7F7F7F"})


def build_title_page(M: dict) -> bytes:
    L = LBL["en" if M.get("lang") == "en" else "id"]
    files, order = _open_template("tp")
    d = files["word/document.xml"].decode("utf8")
    body_open = d.index("<w:body>") + 8
    zi = d.index("0.  ID NASKAH OJS")
    zs = d.rindex("<w:p>", 0, zi)
    ti = d.index("Ditetapkan oleh Dewan Redaksi")
    ts = d.rindex("<w:p>", 0, ti)
    tbl = re.search(r"<w:tbl>.*?</w:tbl>", d, re.S).group(0)
    tbl_open = tbl[:tbl.index("<w:tr>")]
    hdr_row = re.search(r"<w:tr>.*?</w:tr>", tbl, re.S).group(0)
    cell_w = [499, 2251, 599, 2951, 3200]
    CENTER = dict(before=20, after=20, left=120, right=120, jc="center")

    def bar(t):
        return para(run(t, {"b": True, "color": "FFFFFF", "sz": 19}), shd=BLUE, keepNext=True, before=140, after=40, left=120, right=120, jc="center")

    def kv(lab, val, phtext="Penulis mengisi"):
        v = (val if isinstance(val, list) else (val or "")).__class__
        txt = plain_text(val).strip() if val is not None else ""
        return para(run(lab, {"b": True}) + (inline(val) if txt else ph(phtext)), **CENTER)

    def box(val, default):
        v = val if (isinstance(val, list) or (val or "").strip()) else default
        items = paras_of(v) if isinstance(v, str) else [v]
        return "".join(para(inline(p), shd="F8F8F8", before=20, after=20, left=260, right=120, jc="center") for p in items)

    def spacer():
        return para(run(" "), before=60, after=0, jc="center")

    def letters(arr):
        return "(" + ",".join(chr(96 + n) for n in arr) + ")"

    A = [(i, a) for i, a in enumerate(M["authors"]) if (a.get("name") or "").strip()]
    if not A:
        A = [(0, {"name": "", "email": "", "orcid": "", "affs": [], "roles": []})]
    cl = M.get("closing") or {}
    x = bar("0.  ID NASKAH OJS / OJS SUBMISSION ID") + kv("ID Naskah OJS: ", M.get("ojs_id"), "Penulis mengisi setelah mengunggah naskah utama") + spacer()
    x += bar("1.  JUDUL ARTIKEL / ARTICLE TITLE") + kv("Judul (Bahasa Indonesia): ", M.get("title_id")) + kv("Title (English): ", M.get("title_en")) + kv("Judul Singkat / Running Title: ", M.get("running")) + spacer()
    x += bar("2.  PENULIS DAN AFILIASI / AUTHORS AND AFFILIATIONS")
    x += para(run("Cantumkan seluruh penulis sesuai urutan kontribusi beserta alamat e-mail aktif masing-masing (wajib untuk setiap penulis); tanda ★ = penulis korespondensi. / List all authors in order of contribution with an active e-mail address for each (mandatory for every author); ★ = corresponding author.", {"i": True, "color": "444444", "sz": 16}), before=40, after=60, left=120, right=120, jc="center")

    def tc(w, t, f, jc="center"):
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:tcBorders></w:tcBorders><w:shd w:color="auto" w:fill="FFFFFF" w:val="clear"/></w:tcPr>'
                + para(run(t, f) if t else ph(), before=20, after=20, jc=jc) + "</w:tc>")

    rows = hdr_row
    for k, (i, a) in enumerate(A):
        affs = letters(sorted(a.get("affs") or [])) if a.get("affs") else "[isi]"
        rows += ('<w:tr><w:trPr><w:trHeight w:val="270" w:hRule="atLeast"/></w:trPr>' + tc(cell_w[0], f"{k + 1}.", {"b": True, "color": BLUE, "sz": 16})
                 + tc(cell_w[1], (a["name"].strip() + ("  ★" if M.get("corr", 0) == i else "")) if a["name"].strip() else "", {"sz": 16})
                 + tc(cell_w[2], affs, {"b": True, "color": BLUE, "sz": 16}) + tc(cell_w[3], (a.get("email") or "").strip(), {"sz": 16})
                 + tc(cell_w[4], orcid_url(a.get("orcid")), {"sz": 16}) + "</w:tr>")
    x += tbl_open + rows + "</w:tbl>" + spacer()
    x += para(run("Keterangan Afiliasi / Affiliation Details:", {"b": True, "sz": 16}), **CENTER)
    for n, t in enumerate(M.get("affs") or []):
        if (t or "").strip():
            x += para(run(f"({chr(97 + n)}) ", {"b": True, "color": BLUE, "sz": 16}) + inline(t.strip(), {"sz": 16}), **CENTER)
    if not any((t or "").strip() for t in M.get("affs") or []):
        x += para(run("(a) ", {"b": True, "color": BLUE, "sz": 16}) + ph("Penulis mengisi afiliasi lengkap"), **CENTER)
    x += spacer()
    ca = M["authors"][M.get("corr", 0)] if M["authors"] else {}
    affs = M.get("affs") or []
    ai = ((ca.get("affs") or [1])[0] or 1) - 1
    c_aff = (affs[ai] if 0 <= ai < len(affs) else "").strip()
    x += bar("3.  PENULIS KORESPONDENSI / CORRESPONDING AUTHOR")
    x += (kv("Nama Lengkap / Full Name: ", ca.get("name")) + kv("Afiliasi / Affiliation: ", c_aff) + kv("Alamat / Address: ", M.get("corr_address"), "Penulis mengisi alamat lengkap lembaga")
          + kv("Telepon / Phone (opsional / optional): ", M.get("phone"), "Opsional; penulis mengisi") + kv("E-mail: ", ca.get("email")) + kv("ORCID iD: ", orcid_url(ca.get("orcid"))) + spacer())
    x += bar("4.  KONTRIBUSI PENULIS / AUTHOR CONTRIBUTIONS (CRediT)")
    for i, a in A:
        ln = contrib_line(a, i, False, L)
        nm = (a["name"] or "").strip()
        x += para((run(nm + ": ", {"b": True}) if nm else ph("Penulis mengisi nama") + run(": ", {"b": True}))
                  + (run(ln[ln.index(": ") + 2:]) if ln else ph("Penulis mengisi peran CRediT")), **CENTER)
    x += para(run("Peran CRediT / Available Roles:  ", {"b": True, "sz": 15}) + run(" · ".join(CREDIT), {"color": "444444", "sz": 15}), shd="F2F2F2", before=40, after=40, left=120, right=120, jc="center") + spacer()
    x += bar("5.  PERNYATAAN KONFLIK KEPENTINGAN / CONFLICT OF INTEREST STATEMENT") + box(cl.get("coi"), "Para penulis menyatakan tidak ada konflik kepentingan yang relevan dengan artikel ini. / The authors declare no competing interests relevant to this article.") + spacer()
    x += bar("6.  PERNYATAAN PENDANAAN / FUNDING STATEMENT") + box(cl.get("fund"), "Tidak berlaku / Not applicable") + spacer()
    x += bar("7.  UCAPAN TERIMA KASIH / ACKNOWLEDGEMENTS") + box(cl.get("ack"), "Tidak ada / None")
    ai_txt = M["ai_text"] if M.get("ai_use") == "yes" and (M.get("ai_text") or "").strip() else "Tidak menggunakan AI / No AI tools were used"
    x += bar("8.  DEKLARASI PENGGUNAAN AI / AI USE DECLARATION") + box(ai_txt, ai_txt)
    files["word/document.xml"] = (d[:body_open] + d[body_open:zs] + x + d[ts:]).encode("utf8")
    _finish_meta(files, M.get("title_en") or M.get("title_id") or "")
    return _write_package(files, order)


def build(kind: str, M: dict) -> bytes:
    return build_title_page(M) if kind == "title" else build_manuscript(M, kind)
