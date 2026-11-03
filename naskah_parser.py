"""
naskah_parser.py
Membaca naskah .docx bebas-format milik penulis dan mengenali bagian-bagiannya
(judul, penulis, afiliasi, abstrak, kata kunci, bab, tabel, gambar, rumus, pustaka, pernyataan akhir).

Hasilnya dipakai naskah_builder.py. Pengenalan bersifat heuristik, sehingga selalu disertai
laporan temuan dan peringatan agar penulis dapat memeriksa sebelum mengunduh.
"""
from __future__ import annotations

import base64
import io
import re

from docx import Document
from lxml import etree

from naskah_builder import CREDIT, SECTION_KEYS, norm, plain_text, word_count, keywords as split_keywords, _strip_acc

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W, "m": M_NS, "a": A_NS, "r": R_NS}


def q(tag):
    p, n = tag.split(":")
    return "{%s}%s" % (NS[p], n)


# ---------------------------------------------------------------------------
# pembacaan paragraf
# ---------------------------------------------------------------------------
def _on(el, tag):
    e = el.find(q(tag))
    if e is None:
        return False
    v = e.get(q("w:val"))
    return v not in ("0", "false", "off")


def _run_fmt(r):
    pr = r.find(q("w:rPr"))
    f = {"b": False, "i": False, "sup": False, "sub": False}
    if pr is not None:
        f["b"] = _on(pr, "w:b")
        f["i"] = _on(pr, "w:i")
        va = pr.find(q("w:vertAlign"))
        if va is not None:
            v = va.get(q("w:val"))
            f["sup"], f["sub"] = v == "superscript", v == "subscript"
        sz = pr.find(q("w:sz"))
        if sz is not None:
            try:
                f["sz"] = int(sz.get(q("w:val")))
            except Exception:
                pass
    return f


def _run_text(r):
    s = ""
    for c in r:
        t = c.tag
        if t == q("w:t"):
            s += c.text or ""
        elif t in (q("w:tab"), q("w:br"), q("w:cr")):
            s += " "
        elif t == q("w:noBreakHyphen"):
            s += "-"
    return s


def _collect(el, out, doc_part, imgs, warn):
    for c in el:
        t = c.tag
        if t == q("w:r"):
            f = _run_fmt(c)
            s = _run_text(c)
            if s:
                d = dict(f)
                d["t"] = s
                out.append(d)
            for bl in c.iter(q("a:blip")):
                im = _image(bl.get(q("r:embed")), doc_part, warn)
                if im:
                    imgs.append(im)
            for vi in c.iter("{urn:schemas-microsoft-com:vml}imagedata"):
                im = _image(vi.get(q("r:id")), doc_part, warn)
                if im:
                    imgs.append(im)
        elif t in (q("w:hyperlink"), q("w:ins"), q("w:smartTag"), q("w:fldSimple"), q("w:sdt"), q("w:sdtContent")):
            _collect(c, out, doc_part, imgs, warn)
        elif t == q("m:oMath"):
            out.append({"math": etree.tostring(c, encoding="unicode"), "t": ""})
        elif t == q("m:oMathPara"):
            out.append({"math": etree.tostring(c, encoding="unicode"), "t": "", "display": True})


def _image(rid, doc_part, warn):
    if not rid:
        return None
    try:
        part = doc_part.related_parts[rid]
    except Exception:
        return None
    ext = part.partname.ext.lower().lstrip(".")
    blob = part.blob
    if ext in ("emf", "wmf", "svg", "wdp"):
        warn.append(f"Gambar berformat .{ext} tidak dapat dipakai ulang; ganti dengan PNG atau JPG.")
        return None
    w = h = None
    try:
        from PIL import Image
        with Image.open(io.BytesIO(blob)) as im:
            w, h = im.size
            if im.format and im.format.lower() == "mpo":
                ext = "jpeg"
    except Exception:
        pass
    if ext == "jpg":
        ext = "jpeg"
    return {"b64": base64.b64encode(blob).decode(), "ext": ext, "w": w, "h": h}


def _merge(runs):
    out = []
    for r in runs:
        if "math" in r:
            out.append(r)
            continue
        if out and "math" not in out[-1] and all(out[-1].get(k) == r.get(k) for k in ("b", "i", "sup", "sub")):
            out[-1]["t"] += r["t"]
        else:
            out.append(dict(r))
    for r in out:
        if "t" in r and "math" not in r:
            r["t"] = re.sub(r"[    ]+", " ", r["t"])
    # pangkas ujung
    while out and "math" not in out[0] and not out[0]["t"].strip():
        out.pop(0)
    while out and "math" not in out[-1] and not out[-1]["t"].strip():
        out.pop()
    if out and "math" not in out[0]:
        out[0]["t"] = out[0]["t"].lstrip()
    if out and "math" not in out[-1]:
        out[-1]["t"] = out[-1]["t"].rstrip()
    return [r for r in out if "math" in r or r["t"]]


def _para_rec(p, doc_part, warn):
    runs, imgs = [], []
    _collect(p, runs, doc_part, imgs, warn)
    runs = _merge(runs)
    text = "".join(r["t"] for r in runs if "math" not in r)
    sty = ""
    ps = p.find(q("w:pPr") + "/" + q("w:pStyle"))
    if ps is not None:
        sty = (ps.get(q("w:val")) or "").lower()
    textrun = [r for r in runs if "math" not in r and r["t"].strip()]
    allbold = bool(textrun) and all(r.get("b") for r in textrun)
    return {"k": "p", "runs": runs, "text": text.strip(), "style": sty, "bold": allbold, "imgs": imgs,
            "math_only": bool(runs) and not text.strip() and any("math" in r for r in runs)}


def _cell_rec(tc, doc_part, warn):
    paras = [_para_rec(p, doc_part, warn) for p in tc.findall(q("w:p"))]
    runs = []
    for pr in paras:
        if pr["runs"]:
            if runs:
                runs.append({"t": " ", "b": False, "i": False, "sup": False, "sub": False})
            runs += pr["runs"]
    pr_ = tc.find(q("w:tcPr"))
    span, vm = 1, None
    if pr_ is not None:
        gs = pr_.find(q("w:gridSpan"))
        if gs is not None:
            span = int(gs.get(q("w:val")) or 1)
        v = pr_.find(q("w:vMerge"))
        if v is not None:
            vm = "restart" if v.get(q("w:val")) == "restart" else "cont"
    return {"c": _merge(runs), "span": span, "vm": vm}


def read_items(data: bytes, warn: list):
    doc = Document(io.BytesIO(data))
    items = []
    body = doc.element.body
    part = doc.part
    for el in body:
        if el.tag == q("w:p"):
            items.append(_para_rec(el, part, warn))
        elif el.tag == q("w:tbl"):
            trs = el.findall(q("w:tr"))
            if trs and all(len(tr.findall(q("w:tc"))) == 1 for tr in trs) and not any(tc.find(q("w:tcPr") + "/" + q("w:vMerge")) is not None for tc in el.iter(q("w:tc"))):
                for p in el.iter(q("w:p")):
                    items.append(_para_rec(p, part, warn))
                continue
            rows = []
            for tr in el.findall(q("w:tr")):
                rows.append([_cell_rec(tc, part, warn) for tc in tr.findall(q("w:tc"))])
            items.append({"k": "t", "rows": rows})
        elif el.tag == q("w:sdt"):
            for p in el.iter(q("w:p")):
                items.append(_para_rec(p, part, warn))
    return items


# ---------------------------------------------------------------------------
# alat bantu teks
# ---------------------------------------------------------------------------
ID_WORDS = set("yang dan di pada untuk dengan dari adalah ini dalam atau oleh sebagai juga dapat tidak terhadap".split())
EN_WORDS = set("the and of in to with for is was this that are by from as on were which be".split())


def detect_lang(s: str) -> str | None:
    w = re.findall(r"[a-zA-Z]+", s.lower())
    i = sum(x in ID_WORDS for x in w)
    e = sum(x in EN_WORDS for x in w)
    if i == e:
        return None
    return "id" if i > e else "en"


def esc_mark(t: str) -> str:
    return re.sub(r"([*^~\\])", r"\\\1", t)


def to_markup(runs) -> str:
    out = ""
    for r in runs:
        if "math" in r:
            continue
        t = esc_mark(r["t"])
        if not t.strip():
            out += t
            continue
        lead = len(t) - len(t.lstrip())
        trail = len(t) - len(t.rstrip())
        core = t.strip()
        if r.get("sup"):
            core = f"^{core}^"
        elif r.get("sub"):
            core = f"~{core}~"
        if r.get("i"):
            core = f"*{core}*"
        if r.get("b"):
            core = f"**{core}**"
        out += t[:lead] + core + (t[len(t) - trail:] if trail else "")
    return out.strip()


def strip_prefix(runs, rx):
    txt = "".join(r["t"] for r in runs if "math" not in r)
    m = re.match(rx, txt, re.I)
    if not m:
        return runs
    n = m.end()
    out = []
    for r in runs:
        if "math" in r:
            out.append(r)
            continue
        if n >= len(r["t"]):
            n -= len(r["t"])
            continue
        d = dict(r)
        d["t"] = r["t"][n:]
        n = 0
        out.append(d)
    return _merge(out)


def strip_text_prefix(runs, n):
    return strip_prefix(runs, r"^.{%d}" % n)


def _flat_marked(runs) -> str:
    """Teks datar; superskrip dibungkus \\x01...\\x02 dengan koma diganti \\x03."""
    s = ""
    for r in runs:
        if "math" in r:
            continue
        if r.get("sup") and r["t"].strip():
            s += "\x01" + r["t"].replace(",", "\x03").replace(" ", "") + "\x02"
        else:
            s += r["t"]
    return s


# ---------------------------------------------------------------------------
# pengenalan judul bagian
# ---------------------------------------------------------------------------
SEC_NAMES = {
    "intro": ["pendahuluan", "introduction", "latar belakang", "background"],
    "methods": ["bahan dan metode", "metode", "metode penelitian", "metodologi", "materials and methods", "methods",
                "materials and method", "method", "metode dan bahan", "bahan dan metode penelitian", "research methods",
                "methodology", "research method"],
    "results": ["hasil", "results", "hasil penelitian", "hasil dan pembahasan", "results and discussion"],
    "disc": ["pembahasan", "discussion"],
    "concl": ["simpulan", "kesimpulan", "conclusion", "conclusions", "kesimpulan dan saran", "simpulan dan saran",
              "conclusions and recommendations"],
}
SEC_LOOKUP = {n: k for k, v in SEC_NAMES.items() for n in v}

CLOSING = [
    ("coi", r"(konflik kepentingan|competing interests?|conflicts? of interests?|pernyataan konflik)"),
    ("fund", r"(sumber dana|pendanaan|funding|pernyataan pendanaan)"),
    ("ack", r"(ucapan terima kasih|acknowledg\w*|terima kasih)"),
    ("contrib", r"(kontribusi penulis|author.?s?.? contributions?|authors? contribution|credit)"),
    ("data", r"(ketersediaan data|availability of data|data availability|ketersediaan data dan bahan)"),
    ("ethics", r"(persetujuan etik|ethics?\b|etik\b|etika)"),
    ("orcid", r"(orcid)"),
    ("ai", r"(deklarasi penggunaan ai|penggunaan ai|ai use|use of ai|generative ai|kecerdasan buatan|artificial intelligence|ai disclosure|ai declaration|declaration of ai)"),
]
REF_RX = re.compile(r"^(daftar pustaka|references?|referensi|pustaka|bibliography|literature cited)$", re.I)
ABS_RX = re.compile(r"^(abstract|abstrak)\s*(?:[:.\-–—]\s*(.*))?$", re.I | re.S)
KW_RX = re.compile(r"^(key\s*words?|kata\s*kunci)\s*[:.\-–—]?\s*(.*)$", re.I | re.S)
TCAP = re.compile(r"^(tabel|table)\s*\d+\s*[.:\-–]?\s*", re.I)
FCAP = re.compile(r"^(gambar|figure|fig\.)\s*\d+\s*[.:\-–]?\s*", re.I)
NOTE_RX = re.compile(r"^(keterangan|note|notes|catatan|source|sumber)\s*[:.]\s*", re.I)


def _head_name(text):
    t = re.sub(r"^\s*(?:\d+(?:\.\d+)*[.)]?|[IVX]+\.)\s+", "", text.strip())
    t = re.sub(r"\s*[:.]\s*$", "", t)
    return norm(t)


def _is_headingish(rec):
    t = rec["text"]
    if not t or len(t) > 90 or t.endswith(",") or rec["math_only"]:
        return False
    return rec["style"].startswith(("heading", "judul", "title")) or rec["bold"] or bool(re.match(r"^\d+(\.\d+)*[.)]?\s+\S", t))


def match_closing(rec):
    t = rec["text"]
    if not t:
        return None, None
    n0 = norm(re.sub(r"^\s*\d+(\.\d+)*[.)]?\s+", "", t))
    # 1) judul berdiri sendiri
    if len(t) <= 90 and _is_headingish(rec) or len(t) <= 60:
        for k, pat in CLOSING:
            if re.match(pat, n0) and len(n0.split()) <= 9:
                if k == "contrib" and "credit" in n0 and "contribut" not in n0 and len(n0.split()) > 3:
                    continue
                return k, ""
    # 2) satu baris "Judul: isi"
    m = re.match(r"^\s*(?:\d+(?:\.\d+)*[.)]?\s+)?([^:]{3,70}?)\s*:\s*(.+)$", t, re.S)
    if m and rec["runs"] and rec["runs"][0].get("b") is not False:
        h = norm(m.group(1))
        for k, pat in CLOSING:
            if re.match(pat, h) and len(h.split()) <= 8 and rec["runs"][0].get("b"):
                return k, m.group(2)
    return None, None


# ---------------------------------------------------------------------------
# bagian depan
# ---------------------------------------------------------------------------
AFF_WORDS = re.compile(r"\b(universitas|university|institut|institute|fakultas|faculty|department|departemen|politeknik|"
                       r"jurusan|program studi|sekolah|school|college|badan|balai|lembaga|laboratorium|laboratory|"
                       r"pusat|centre|center|dinas|kementerian|rumah sakit|hospital|indonesia)\b", re.I)
EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
ORCID = re.compile(r"\d{4}-\d{4}-\d{4}-\d{3}[\dXx]")
SKIP_FRONT = re.compile(r"^(diterima|received|accepted|diterbitkan|published|revised|direvisi|doi\b|article info|info artikel|"
                        r"naskah untuk telaah|manuscript under review|jurnal ilmiah platax|e-?issn|vol\.)", re.I)


def _label_idx(lab: str):
    lab = lab.strip().strip("()")
    if lab.isdigit():
        return int(lab)
    if len(lab) == 1 and lab.isalpha():
        return ord(lab.lower()) - 96
    return None


def _parse_authors(auth_runs_list):
    """Kembalikan list penulis {name, labels:set, corr:bool}."""
    s = ""
    for runs in auth_runs_list:
        part = _flat_marked(runs)
        if s and part:
            s += ", "
        s += part
    s = re.sub(r"\s+", " ", s)
    s = s.replace("★", "\x01*\x02")  # ★
    parts = re.split(r"\s*(?:,|;|&|\band\b|\bdan\b)\s*(?![^\x01]*\x02)", s)
    # sup marker setelah pemisah: "Nama, 1" → tempelkan ke nama sebelumnya
    people = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        marks = re.findall(r"\x01([^\x02]*)\x02", p)
        name = re.sub(r"\x01[^\x02]*\x02", "", p)
        star = "*" in name or any("*" in m for m in marks)
        name = re.sub(r"[*†‡§¶\d]+", " ", name)
        name = re.sub(r"\s+", " ", name).strip(" ,.;")
        labels = set()
        for m in marks:
            for tok in re.split(r"[\x03,]", m):
                i = _label_idx(tok.replace("*", "").replace("†", "")) if tok.replace("*", "").strip() else None
                if i:
                    labels.add(i)
        if not name:
            if people:
                people[-1]["labels"] |= labels
                people[-1]["corr"] |= star
            continue
        if len(name) < 3 or len(name.split()) > 7:
            continue
        people.append({"name": name, "labels": labels, "corr": star})
    return people


def _parse_front(front, warns, infos):
    R = {"title_id": "", "title_en": "", "running": "", "authors": [], "affs": [], "emails": [], "orcids": [], "corr_hint": None}
    titles = []
    auth, affs_raw = [], []
    stage = "title"
    for rec in front:
        t = rec["text"]
        if not t:
            continue
        if SKIP_FRONT.match(t):
            continue
        rm = re.match(r"^(running title|judul singkat)[^:]*:\s*(.+)$", t, re.I)
        if rm:
            R["running"] = rm.group(2).strip()
            continue
        em = EMAIL.findall(t)
        if em:
            R["emails"] += em
            if len(t) - sum(len(e) for e in em) < 12 or re.match(r"^(corresponding|penulis korespondensi|email|e-mail|correspondence|\*)", t, re.I):
                R["orcids"] += ORCID.findall(t)
                continue
        if re.match(r"^(\*|corresponding|penulis korespondensi|correspondence|korespondensi)", t, re.I) and not AFF_WORDS.search(t):
            continue
        if ORCID.search(t) and not AFF_WORDS.search(t):
            R["orcids"] += ORCID.findall(t)
            continue
        has_sup = any(r.get("sup") and r["t"].strip() for r in rec["runs"] if "math" not in r)
        is_aff = bool(AFF_WORDS.search(t)) and (stage in ("author", "aff") or len(t) > 60) and not (stage == "title" and not titles)
        lead = re.match(r"^\(?(?:[1-9][.)]?|[a-d][.)])\s+\S", t)
        if stage == "title" and not titles:
            titles.append(rec)
            continue
        if stage == "title":
            l1 = detect_lang(titles[0]["text"])
            l2 = detect_lang(t)
            if len(titles) == 1 and l1 and l2 and l1 != l2 and not has_sup and not is_aff and len(t) > 15:
                titles.append(rec)
                continue
            stage = "author"
        if is_aff or (lead and stage in ("author", "aff")) or (stage == "aff" and AFF_WORDS.search(t)):
            stage = "aff"
            affs_raw.append(rec)
            continue
        if stage == "author":
            auth.append(rec)
        elif stage == "aff" and has_sup:
            affs_raw.append(rec)
    # judul
    for rec in titles:
        lg = detect_lang(rec["text"])
        key = "title_id" if lg == "id" else "title_en" if lg == "en" else None
        mk = to_markup(rec["runs"])
        mk = re.sub(r"^\*\*(.*)\*\*$", r"\1", mk) if mk.startswith("**") and mk.endswith("**") and mk.count("**") == 2 else mk
        if key is None:
            key = "title_id" if not R["title_id"] and (not R["title_en"] or len(titles) == 1 and False) else "title_en"
            if len(titles) == 1:
                key = "title_en" if R["title_en"] == "" and False else "title_id"
        if R[key]:
            key = "title_en" if key == "title_id" else "title_id"
        R[key] = mk
    R["_single_title"] = len(titles) == 1
    # penulis
    people = _parse_authors([r["runs"] for r in auth]) if auth else []
    # afiliasi
    aff_map = {}
    seq = 0
    for rec in affs_raw:
        s = _flat_marked(rec["runs"])
        pieces = re.split(r"\x01([^\x02]*)\x02", s)
        found = False
        if len(pieces) > 1:
            lead_txt = pieces[0].strip()
            for j in range(1, len(pieces), 2):
                lab = pieces[j].replace("\x03", ",")
                txt = pieces[j + 1].strip(" ;,\n") if j + 1 < len(pieces) else ""
                if txt:
                    idx = _label_idx(lab.split(",")[0])
                    seq += 1
                    aff_map[idx or seq] = txt
                    found = True
        if not found:
            t = rec["text"]
            m = re.match(r"^\(?(?:([1-9])[.)]?|([a-d])[.)])\s+(.+)$", t)
            seq += 1
            if m:
                aff_map[_label_idx(m.group(1) or m.group(2))] = m.group(3).strip()
            else:
                aff_map[seq] = t.strip()
    return R, people, aff_map


# ---------------------------------------------------------------------------
# pengurai utama
# ---------------------------------------------------------------------------
def _link_captions(items):
    """Gabungkan gambar berurutan, pasangkan keterangan tabel/gambar, dan tandai paragraf yang terpakai."""
    out = []
    i = 0
    n = len(items)
    while i < n:
        it = items[i]
        if it["k"] == "p" and it["imgs"] and not (it["text"] and len(it["text"]) > 150):
            imgs = list(it["imgs"])
            cap = it["text"] if FCAP.match(it["text"] or "") else None
            capr = it["runs"] if cap else None
            j = i + 1
            while j < n and items[j]["k"] == "p" and items[j]["imgs"] and not items[j]["text"]:
                imgs += items[j]["imgs"]
                j += 1
            if not cap and j < n and items[j]["k"] == "p" and FCAP.match(items[j]["text"] or ""):
                capr = items[j]["runs"]
                j += 1
            elif not cap and out and out[-1]["k"] == "p" and FCAP.match(out[-1]["text"] or "") and not out[-1].get("used"):
                capr = out.pop()["runs"]
            out.append({"k": "fig", "imgs": imgs, "cap": strip_prefix(capr, FCAP.pattern) if capr else []})
            i = j
            continue
        if it["k"] == "t":
            capr = None
            if out and out[-1]["k"] == "p" and TCAP.match(out[-1]["text"] or ""):
                capr = out.pop()["runs"]
            note = None
            j = i + 1
            if not capr and j < n and items[j]["k"] == "p" and TCAP.match(items[j]["text"] or ""):
                capr = items[j]["runs"]
                j += 1
            if j < n and items[j]["k"] == "p" and NOTE_RX.match(items[j]["text"] or ""):
                note = strip_prefix(items[j]["runs"], NOTE_RX.pattern)
                j += 1
            out.append({"k": "tbl", "rows": it["rows"], "cap": strip_prefix(capr, TCAP.pattern) if capr else [], "note": note or []})
            i = j
            continue
        out.append(it)
        i += 1
    return out


def parse_docx(data: bytes):
    warns: list[str] = []
    infos: list[str] = []
    items = read_items(data, warns)
    items = _link_captions(items)

    def is_p(x):
        return x["k"] == "p"

    # --- batas bagian depan
    front_end = len(items)
    for idx, it in enumerate(items):
        if not is_p(it):
            continue
        t = it["text"]
        if ABS_RX.match(t) or KW_RX.match(t) or _head_name(t) in SEC_LOOKUP and _is_headingish(it):
            front_end = idx
            break
    front = [x for x in items[:front_end] if is_p(x)]
    F, people, aff_map = _parse_front(front, warns, infos)

    M = {"lang": "id", "title_id": F["title_id"], "title_en": F["title_en"], "running": F["running"], "authors": [], "affs": [],
         "corr": 0, "phone": "", "abs_id": "", "abs_en": "", "kw_id": "", "kw_en": "", "sections": {k: [] for k in SECTION_KEYS},
         "closing": {}, "refs": [], "ai_use": "", "ai_text": "", "header_name": "", "color_cites": True, "link_cites": True,
         "sort_refs": True, "sort_kw": True, "ojs_id": ""}
    # afiliasi → daftar berurutan
    order = sorted(aff_map)
    remap = {old: i + 1 for i, old in enumerate(order)}
    M["affs"] = [aff_map[o] for o in order]
    emails, orcids = F["emails"], F["orcids"]
    for i, p in enumerate(people):
        labs = sorted(remap[l] for l in p["labels"] if l in remap)
        if not labs and len(M["affs"]) == 1:
            labs = [1]
        M["authors"].append({"name": p["name"], "email": "", "orcid": "", "affs": labs, "roles": []})
        if p["corr"]:
            M["corr"] = i
    if len(emails) == len(people) and people:
        for a, e in zip(M["authors"], emails):
            a["email"] = e
    elif emails and M["authors"]:
        M["authors"][M["corr"]]["email"] = emails[0]
    if len(orcids) == len(people) and people:
        for a, o in zip(M["authors"], orcids):
            a["orcid"] = o
    if not M["authors"]:
        infos.append("Nama penulis tidak ditemukan; dianggap naskah blind-review (Title Page tidak dibuat).")

    # --- badan naskah
    rest = items[front_end:]
    mode = "pre"
    cur = None
    ckey = None
    abs_lang = None
    kw_done = False
    h2 = 0
    cl_raw: dict[str, list] = {}
    refs: list = []
    sec_found = []

    def put_abs(lang, runs):
        key = "abs_id" if lang == "id" else "abs_en"
        mk = to_markup(runs)
        if mk:
            M[key] = (M[key] + " " + mk).strip()

    for it in rest:
        k = it["k"]
        if k in ("tbl", "fig"):
            if mode == "body" and cur:
                if k == "tbl":
                    ncol = max((sum(c["span"] for c in r) for r in it["rows"]), default=0)
                    M["sections"][cur].append({"t": "table", "cap": it["cap"], "rows": it["rows"], "note": it["note"], "full": ncol >= 6})
                else:
                    M["sections"][cur].append({"t": "fig", "cap": it["cap"], "imgs": it["imgs"], "full": False})
            elif mode == "pre":
                warns.append("Ada tabel/gambar sebelum bab pertama; dilewati.")
            continue
        if k == "t":
            continue
        t = it["text"]
        if not t and not it["math_only"]:
            continue
        # pustaka
        if REF_RX.match(re.sub(r"^\s*\d+[.)]?\s+", "", t).strip()) and len(t) < 40:
            mode = "refs"
            continue
        am = ABS_RX.match(t)
        if am and len(t) < 20 or (am and am.group(2) and mode in ("pre", "abs")):
            abs_lang = "id" if am.group(1).lower() == "abstrak" else "en"
            mode = "abs"
            kw_done = False
            if am.group(2):
                put_abs(abs_lang, strip_prefix(it["runs"], r"^(abstract|abstrak)\s*[:.\-–—]?\s*"))
            continue
        km = KW_RX.match(t)
        if km and mode in ("pre", "abs"):
            body = km.group(2)
            lang = abs_lang or (detect_lang(" ".join(plain_text(M["abs_id"]).split()[:30])) or "id")
            if t.lower().startswith("kata"):
                lang = "id"
            elif t.lower().startswith("key"):
                lang = "en"
            M["kw_id" if lang == "id" else "kw_en"] = re.sub(r"\s+", " ", body).strip()
            mode = "abs"
            continue
        # judul bab
        hn = _head_name(t)
        top = SEC_LOOKUP.get(hn) if _is_headingish(it) or len(t) < 40 else None
        ck, inline_txt = match_closing(it)
        if ck and mode != "refs":
            mode = "closing"
            ckey = ck
            cl_raw.setdefault(ck, [])
            if inline_txt:
                cl_raw[ck].append(strip_prefix(it["runs"], r"^[^:]*:\s*"))
            continue
        if top and mode in ("pre", "abs", "body", "closing") and (_is_headingish(it) or len(t) < 40) and not (mode == "closing" and not it["bold"] and not it["style"].startswith("heading") and not re.match(r"^\d", t)):
            mode = "body"
            cur = top
            h2 = 0
            sec_found.append(top)
            if hn in ("hasil dan pembahasan", "results and discussion"):
                infos.append("Bab “Hasil dan Pembahasan” digabung dibaca sebagai bab Hasil; bab Pembahasan dikosongkan.")
            continue
        if mode == "abs":
            if kw_done:
                continue
            put_abs(abs_lang or "id", it["runs"])
            continue
        if mode == "pre":
            if not M["authors"] and False:
                pass
            continue
        if mode == "body":
            sm = re.match(r"^(\d+)\.(\d+)(?:\.(\d+))?[.)]?\s+(.+)$", t)
            style_h = it["style"].startswith("heading") and len(t) < 150
            if sm and len(t) < 150 and not it["math_only"]:
                lvl = "h3" if sm.group(3) else "h2"
                M["sections"][cur].append({"t": lvl, "c": re.sub(r"^\d+(\.\d+)+[.)]?\s+", "", t)})
                continue
            if style_h and not re.match(r"^heading\s*1$|^judul\s*1$", it["style"]) or (style_h and top is None and it["style"] in ("heading1", "judul1") and False):
                lvl = "h3" if re.search(r"(heading|judul)\s*3", it["style"]) else "h2"
                M["sections"][cur].append({"t": lvl, "c": re.sub(r"^\s*\d+(\.\d+)*[.)]?\s+", "", t)})
                continue
            if it["style"] in ("heading1", "judul1") and len(t) < 90:
                M["sections"][cur].append({"t": "h2", "c": re.sub(r"^\s*\d+(\.\d+)*[.)]?\s+", "", t)})
                continue
            if it["bold"] and len(t) < 90 and not t.endswith((".", ",", ";", ":")) and not it["math_only"] and not t.endswith(")"):
                M["sections"][cur].append({"t": "h2", "c": re.sub(r"^\s*\d+(\.\d+)*[.)]?\s+", "", t)})
                continue
            M["sections"][cur].append({"t": "p", "c": it["runs"], "display": it["math_only"]})
            continue
        if mode == "closing":
            cl_raw[ckey].append(it["runs"])
            continue
        if mode == "refs":
            runs = strip_prefix(it["runs"], r"^\s*(?:\[\d+\]|\d+[.)])\s*")
            if runs:
                refs.append(runs)
            continue

    M["refs"] = refs
    # tandai paragraf rumus tunggal sebagai tampilan
    for k in SECTION_KEYS:
        for b in M["sections"][k]:
            if b.get("t") == "p" and b.pop("display", False):
                b["c"] = [r for r in b["c"] if "math" in r]
                if len(b["c"]) == 1:
                    b["c"][0]["display"] = True
            elif b.get("t") == "p":
                b.pop("display", None)

    # --- pernyataan akhir
    for key, paras in cl_raw.items():
        text = "\n".join(x for x in (to_markup(r) for r in paras) if x)
        if key in ("coi", "fund", "ack", "data", "ethics"):
            if text and not re.match(r"^\*?(tidak ada|none|n/?a|not applicable|tidak berlaku)\.?\*?$", text, re.I) or text:
                M["closing"][key] = text
        elif key == "contrib":
            _roles_from_text(M, text, warns)
        elif key == "orcid":
            for line in text.split("\n"):
                oc = ORCID.search(line)
                if oc:
                    nm = norm(line[:oc.start()])
                    for a in M["authors"]:
                        if norm(a["name"]) and (norm(a["name"]) in nm or nm and nm in norm(a["name"])) and not a["orcid"]:
                            a["orcid"] = oc.group(0)
        elif key == "ai":
            if re.search(r"\b(tidak|no|not|none|without)\b", text, re.I) and re.search(r"\b(ai|kecerdasan|artificial|generative)\b", text, re.I) \
                    and not re.search(r"\b(menggunakan|used|utili[sz]ed|employed|assist)", text.split(".")[0], re.I) or re.search(r"\b(tidak menggunakan|no ai|not use|did not use|none)\b", text, re.I):
                M["ai_use"] = "no"
            elif text:
                M["ai_use"], M["ai_text"] = "yes", text
    return _finalize(M, warns, infos, sec_found, F)


def _roles_from_text(M, text, warns):
    if not text:
        return
    got = 0
    for line in re.split(r"\n|(?<=\.)\s+(?=[A-Z])", text):
        parts = re.split(r"\s*[:–-]\s+", line, maxsplit=1)
        if len(parts) < 2:
            continue
        who, what = parts
        roles = [r for r in CREDIT if re.search(re.escape(r.lower().replace("–", "-")), what.lower().replace("–", "-").replace("&", "&"))]
        # nama bilangan: Penulis 1 / Author 2
        target = None
        m = re.match(r"^\W*(?:penulis|author)\s*(\d+)", who, re.I)
        if m and int(m.group(1)) <= len(M["authors"]):
            target = int(m.group(1)) - 1
        else:
            nw = norm(who)
            for i, a in enumerate(M["authors"]):
                na = norm(a["name"])
                sn = na.split()[-1] if na else ""
                if na and (na in nw or nw in na or (sn and sn in nw.split())):
                    target = i
                    break
        if target is not None and roles:
            M["authors"][target]["roles"] = sorted(set(M["authors"][target]["roles"] + roles), key=CREDIT.index)
            got += 1
    if not got:
        warns.append("Bagian kontribusi penulis ditemukan tetapi peran CRediT tidak dapat dicocokkan dengan nama penulis.")


def _finalize(M, warns, infos, sec_found, F):
    # bahasa naskah
    sample = " ".join(plain_text(b.get("c", "")) for k in SECTION_KEYS for b in M["sections"][k] if b.get("t") == "p")
    lg = detect_lang(sample) or detect_lang(plain_text(M["title_id"]) + " " + plain_text(M["title_en"])) or "id"
    M["lang"] = lg
    # judul tunggal
    if F.get("_single_title"):
        only = M["title_id"] or M["title_en"]
        lang_t = detect_lang(plain_text(only))
        if lang_t == "id" or (lang_t is None and lg == "id"):
            M["title_id"], M["title_en"] = only, ""
        else:
            M["title_en"], M["title_id"] = only, ""
    # judul singkat
    if not M["running"]:
        base = plain_text(M["title_en"] or M["title_id"])
        if len(base) > 60:
            base = base[:60].rsplit(" ", 1)[0]
        M["running"] = base.strip(" ,:;-")
        if base:
            infos.append("Judul singkat tidak tertulis; dibuat otomatis dari judul (maks. 60 karakter).")
    # akhir: ringkasan temuan
    rep = {"warnings": list(dict.fromkeys(warns)), "infos": infos, "found": {}, "problems": []}
    P = rep["problems"]
    f = rep["found"]
    f["Judul (Indonesia)"] = plain_text(M["title_id"])
    f["Judul (Inggris)"] = plain_text(M["title_en"])
    f["Jumlah penulis"] = len(M["authors"])
    f["Afiliasi"] = len(M["affs"])
    f["Abstrak (kata)"] = f"ID {word_count(M['abs_id'])} / EN {word_count(M['abs_en'])}"
    f["Kata kunci"] = f"ID {len(split_keywords(M['kw_id']))} / EN {len(split_keywords(M['kw_en']))}"
    f["Bab terdeteksi"] = ", ".join(sec_found) or "—"
    f["Tabel"] = sum(1 for k in SECTION_KEYS for b in M["sections"][k] if b.get("t") == "table")
    f["Gambar"] = sum(1 for k in SECTION_KEYS for b in M["sections"][k] if b.get("t") == "fig")
    f["Rumus"] = sum(1 for k in SECTION_KEYS for b in M["sections"][k] if b.get("t") == "p" and any("math" in r for r in (b["c"] if isinstance(b["c"], list) else [])))
    f["Pustaka"] = len(M["refs"])
    if not (M["title_id"] or M["title_en"]):
        P.append("Judul tidak ditemukan. Pastikan judul berada di baris paling atas naskah.")
    if not M["title_en"] or not M["title_id"]:
        P.append("Hanya satu bahasa judul yang ditemukan; templat meminta judul Indonesia dan Inggris.")
    if not M["abs_id"] or not M["abs_en"]:
        P.append("Abstrak dwibahasa belum lengkap (butuh “Abstrak” dan “Abstract”).")
    for lab, ab in (("Indonesia", M["abs_id"]), ("Inggris", M["abs_en"])):
        wc = word_count(ab)
        if ab and not 100 <= wc <= 300:
            P.append(f"Abstrak {lab} {wc} kata (disarankan sekitar 150–250 kata).")
    for lab, kw in (("Indonesia", M["kw_id"]), ("Inggris", M["kw_en"])):
        n = len(split_keywords(kw))
        if kw and not 3 <= n <= 6:
            P.append(f"Kata kunci {lab} berjumlah {n} (disarankan 3–5).")
        if not kw:
            P.append(f"Kata kunci {lab} tidak ditemukan.")
    for k, lbl in zip(SECTION_KEYS, ("Pendahuluan", "Metode", "Hasil", "Pembahasan", "Simpulan")):
        if not M["sections"][k]:
            P.append(f"Bab {lbl} tidak ditemukan atau kosong.")
    if not M["refs"]:
        P.append("Daftar pustaka tidak ditemukan.")
    if M["authors"]:
        if not any(a["email"] for a in M["authors"]):
            P.append("Alamat e-mail penulis tidak ditemukan (wajib untuk Title Page).")
        if not M["affs"]:
            P.append("Afiliasi penulis tidak ditemukan.")
        miss = [a["name"] for a in M["authors"] if not a["roles"]]
        if miss:
            P.append("Peran CRediT belum ada untuk: " + ", ".join(miss) + ".")
    # sitiran tanpa pustaka
    from naskah_builder import build_ref_index, cite_spans
    ri = build_ref_index(M["refs"])
    bad = set()
    total = 0
    for k in SECTION_KEYS:
        for b in M["sections"][k]:
            if b.get("t") == "p":
                pt = plain_text(b["c"]) if isinstance(b["c"], list) else b["c"]
                for sp in cite_spans(pt, ri):
                    if sp.get("label") and not sp.get("no_count"):
                        total += 1
                        if not sp.get("anchor"):
                            bad.add(sp["label"])
    f["Sitiran dikenali"] = total
    if bad and M["refs"]:
        rep["warnings"].append("Sitiran yang tidak ditemukan padanannya di daftar pustaka: " + "; ".join(sorted(bad)[:12]) + (" …" if len(bad) > 12 else ""))
    return M, rep
