"""
Penata Naskah JIP / PLATAX 2026 – versi unggah (Indonesia / English).
Penulis cukup mengunggah naskah .docx; aplikasi menerapkan templat resmi JIP 2026.
Jalankan / Run:  streamlit run app.py
"""
import hashlib
import re
from pathlib import Path

import streamlit as st

import naskah_builder as B
import naskah_parser as P

HERE = Path(__file__).resolve().parent
TPL = HERE / "templates"
MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
GUIDE = "Petunjuk_Penggunaan_Template_JIP_2026_Bilingual_v4.docx"

st.set_page_config(page_title="JIP 2026 – Manuscript Formatter", page_icon="📄", layout="centered")
st.markdown(
    """
<style>
.block-container{max-width:780px;padding-top:2rem}
h1{font-size:1.6rem!important;line-height:1.25!important}
.note{border-left:4px solid #007BB8;background:rgba(0,123,184,.08);padding:.7rem .9rem;border-radius:0 6px 6px 0;font-size:.93rem;margin:.4rem 0 1rem}
.must{border-left:4px solid #C0392B;background:rgba(192,57,43,.08);padding:.7rem .9rem;border-radius:0 6px 6px 0;font-size:.95rem;margin:.4rem 0 .6rem}
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Teks antarmuka dwibahasa
# ---------------------------------------------------------------------------
T = {
    "id": {
        "title": "Penata Naskah Jurnal Ilmiah PLATAX",
        "sub": "Unggah naskah .docx Anda. Aplikasi hanya menerapkan tata letak templat JIP 2026 pada isi naskah Anda.",
        "optional": "<b>Alat ini bersifat opsional.</b> Anda tetap boleh menyerahkan naskah memakai templat Word manual (lihat bagian bawah). "
                    "Pada tahap awal, kerapian format <b>tidak</b> menjadi alasan penolakan naskah. Redaksi menilai isi lebih dulu.",
        "read_first": "<b>Baca dulu sebelum mengunggah:</b><br>PETUNJUK PENGGUNAAN TEMPLATE ARTIKEL<br>GUIDE TO USING THE ARTICLE TEMPLATE",
        "guide_btn": "⬇️  Unduh petunjuk (dwibahasa)",
        "uploader": "Unggah naskah (.docx)",
        "uploader_help": "Satu berkas .docx berisi judul, penulis, abstrak, isi naskah, dan daftar pustaka. Format lama (.doc) harap disimpan ulang sebagai .docx.",
        "reading": "Membaca naskah…",
        "bad_file": "Berkas tidak dapat dibaca. Pastikan berkas adalah .docx yang valid (bukan .doc yang diganti namanya). Rincian teknis: ",
        "s1": "1. Periksa hasil pembacaan",
        "attention": "Perlu diperhatikan (tidak menghalangi pembuatan berkas):",
        "all_ok": "Semua bagian utama terbaca.",
        "outline": "Lihat kerangka yang terbaca",
        "refs3": "**Daftar pustaka (3 pertama)**",
        "authors": "**Penulis**",
        "para": "paragraf", "tbl": "tabel", "fig": "gambar",
        "aff_lbl": "afiliasi", "email_lbl": "e-mail",
        "f_title_id": "Judul (Indonesia)", "f_title_en": "Judul (Inggris)", "f_lang": "Bahasa naskah / templat",
        "f_authors": "Jumlah penulis", "f_affs": "Afiliasi", "f_abstract": "Abstrak (kata)", "f_keywords": "Kata kunci",
        "f_sections": "Bab terdeteksi", "f_tables": "Tabel", "f_figures": "Gambar", "f_equations": "Rumus",
        "f_refs": "Pustaka", "f_cites": "Sitiran dikenali",
        "lang_en": "English → templat Article (EN)", "lang_id": "Indonesia → templat Artikel (ID)",
        "s_titles": "2. Judul dua bahasa dan judul singkat",
        "titles_note": "Templat meminta judul dalam bahasa Indonesia dan Inggris, serta judul singkat (maks. 60 karakter).",
        "title_id": "Judul bahasa Indonesia", "title_en": "Judul bahasa Inggris", "running": "Judul singkat (running title)",
        "missing_title": "Judul bahasa {x} belum ada. Isi di bawah (wajib).",
        "x_id": "Indonesia", "x_en": "Inggris",
        "gen_btn": "Buat terjemahan otomatis", "gen_wait": "Menerjemahkan…", "gen_fail": "Terjemahan otomatis gagal; isi manual.",
        "gen_hint": "Terjemahan otomatis tersedia jika administrator mengisi ANTHROPIC_API_KEY pada Streamlit secrets. Periksa hasilnya.",
        "title_ok": "Judul dua bahasa dan judul singkat lengkap.",
        "s_auth": "3. Data penulis untuk Title Page",
        "auth_found": "Data penulis terbaca dari naskah Anda: {n} penulis, {a} afiliasi. Tidak perlu mengunggah lagi.",
        "auth_corr": "Koreksi data penulis (opsional)",
        "auth_none": "Naskah ini tidak memuat nama penulis (blind review). Title Page memerlukan data penulis. Lengkapi dengan salah satu cara di bawah.",
        "tab_up": "Unggah naskah lengkap (ada nama penulis)", "tab_form": "Isi singkat di sini",
        "full_up": "Naskah versi lengkap (.docx)",
        "full_help": "Versi naskah yang memuat nama, afiliasi, dan e-mail penulis. Hanya data penulis yang diambil.",
        "read_n": "Terbaca {n} penulis dan {a} afiliasi.", "no_auth_file": "Nama penulis tidak ditemukan di berkas ini.",
        "col_name": "Nama", "col_email": "E-mail", "col_orcid": "ORCID", "col_aff": "No. afiliasi",
        "aff_col": "Afiliasi (nomor 1, 2, … sesuai urutan)", "roles": "**Peran kontribusi (CRediT)**",
        "extra": "Pengaturan tambahan (opsional)", "ojs": "ID Naskah OJS (untuk Title Page)", "corr": "Penulis korespondensi",
        "phone": "Telepon penulis korespondensi (opsional)", "mslang": "Bahasa naskah",
        "tpl_auto": "Templat dipilih otomatis dari bahasa isi naskah: naskah berbahasa Inggris memakai templat EN, berbahasa Indonesia memakai templat ID. Jika keliru, ubah di “Pengaturan tambahan”.",
        "s_dl": "4. Unduh berkas",
        "dl_note": "Dua berkas untuk OJS dibuat dari templat resmi. Seluruh isi tetap milik Anda; hanya tata letak yang diterapkan.",
        "dl_fail": "Berkas gagal dibuat: ",
        "dl_blind": "⬇️  Naskah blind review (.docx) – untuk berkas utama OJS",
        "dl_tp": "⬇️  Title Page (.docx) – berkas tambahan terpisah",
        "tp_ph": "Sebagian data belum diisi, sehingga Title Page berisi penanda abu-abu [Penulis mengisi …]. Isi langsung di Word, atau lengkapi data di atas lalu unduh ulang.",
        "check_word": "Buka hasilnya di Microsoft Word dan periksa sekali sebelum mengunggah ke OJS. Jika ada bagian yang keliru terbaca, perbaiki langsung di Word atau di naskah asli lalu unggah ulang.",
        "backup": "Cadangan: templat Word manual",
        "backup_sub": "Jika alat ini tidak cocok untuk naskah Anda, pakai templat berikut secara langsung.",
        "privacy": "Naskah yang Anda unggah hanya diproses di memori selama sesi dan tidak disimpan oleh aplikasi ini.",
        "tpl_names": ["Artikel (Indonesia)", "Article (English)", "Naskah blind review (Indonesia)", "Manuscript blind review (English)", "Title Page", "Petunjuk penggunaan (dwibahasa)"],
    },
    "en": {
        "title": "Manuscript Formatter – Jurnal Ilmiah PLATAX",
        "sub": "Upload your .docx manuscript. The app only applies the JIP 2026 template layout to your content.",
        "optional": "<b>This tool is optional.</b> You may still submit using the manual Word templates (see the bottom of this page). "
                    "At the initial stage, formatting neatness is <b>not</b> a reason for rejection. The editors assess content first.",
        "read_first": "<b>Please read before uploading:</b><br>PETUNJUK PENGGUNAAN TEMPLATE ARTIKEL<br>GUIDE TO USING THE ARTICLE TEMPLATE",
        "guide_btn": "⬇️  Download the guide (bilingual)",
        "uploader": "Upload your manuscript (.docx)",
        "uploader_help": "One .docx file containing the title, authors, abstract, main text and references. Please re-save old .doc files as .docx.",
        "reading": "Reading the manuscript…",
        "bad_file": "The file could not be read. Make sure it is a valid .docx (not a renamed .doc). Technical details: ",
        "s1": "1. Check what was read",
        "attention": "Please note (these do not block file creation):",
        "all_ok": "All main parts were read.",
        "outline": "Show the detected outline",
        "refs3": "**References (first 3)**",
        "authors": "**Authors**",
        "para": "paragraphs", "tbl": "tables", "fig": "figures",
        "aff_lbl": "affiliations", "email_lbl": "e-mail",
        "f_title_id": "Title (Indonesian)", "f_title_en": "Title (English)", "f_lang": "Manuscript language / template",
        "f_authors": "Number of authors", "f_affs": "Affiliations", "f_abstract": "Abstract (words)", "f_keywords": "Keywords",
        "f_sections": "Sections detected", "f_tables": "Tables", "f_figures": "Figures", "f_equations": "Equations",
        "f_refs": "References", "f_cites": "Citations recognised",
        "lang_en": "English → Article template (EN)", "lang_id": "Indonesian → Artikel template (ID)",
        "s_titles": "2. Bilingual title and running title",
        "titles_note": "The template requires the title in both Indonesian and English, plus a running title (max. 60 characters).",
        "title_id": "Title in Indonesian", "title_en": "Title in English", "running": "Running title",
        "missing_title": "The {x} title is missing. Please fill it in below (required).",
        "x_id": "Indonesian", "x_en": "English",
        "gen_btn": "Generate translation", "gen_wait": "Translating…", "gen_fail": "Automatic translation failed; please type it in.",
        "gen_hint": "Automatic translation is available when the administrator sets ANTHROPIC_API_KEY in Streamlit secrets. Please check the result.",
        "title_ok": "Bilingual title and running title are complete.",
        "s_auth": "3. Author data for the Title Page",
        "auth_found": "Author data were read from your manuscript: {n} authors, {a} affiliations. No need to upload again.",
        "auth_corr": "Correct author data (optional)",
        "auth_none": "This manuscript has no author names (blind review). The Title Page needs author data. Complete it in one of the ways below.",
        "tab_up": "Upload the full manuscript (with author names)", "tab_form": "Fill in briefly here",
        "full_up": "Full manuscript version (.docx)",
        "full_help": "The version that contains author names, affiliations and e-mails. Only author data are taken from it.",
        "read_n": "Read {n} authors and {a} affiliations.", "no_auth_file": "No author names were found in this file.",
        "col_name": "Name", "col_email": "E-mail", "col_orcid": "ORCID", "col_aff": "Affiliation no.",
        "aff_col": "Affiliation (numbered 1, 2, … in order)", "roles": "**Author contributions (CRediT)**",
        "extra": "Additional settings (optional)", "ojs": "OJS submission ID (for the Title Page)", "corr": "Corresponding author",
        "phone": "Corresponding author phone (optional)", "mslang": "Manuscript language",
        "tpl_auto": "The template is chosen automatically from the language of your text: English manuscripts use the EN template, Indonesian ones the ID template. If wrong, change it under “Additional settings”.",
        "s_dl": "4. Download",
        "dl_note": "The two OJS files are built from the official template. Your content stays yours; only the layout is applied.",
        "dl_fail": "File creation failed: ",
        "dl_blind": "⬇️  Blind-review manuscript (.docx) – main file for OJS",
        "dl_tp": "⬇️  Title Page (.docx) – separate supplementary file",
        "tp_ph": "Some data are not filled in, so the Title Page contains grey [Penulis mengisi …] placeholders. Fill them in Word, or complete the data above and download again.",
        "check_word": "Open the files in Microsoft Word and check them once before uploading to OJS. If anything was misread, fix it in Word or in your original manuscript and upload again.",
        "backup": "Backup: manual Word templates",
        "backup_sub": "If this tool does not suit your manuscript, use these templates directly.",
        "privacy": "Your uploaded manuscript is processed in memory during the session only and is not stored by this app.",
        "tpl_names": ["Artikel (Indonesian)", "Article (English)", "Blind-review manuscript (Indonesian)", "Blind-review manuscript (English)", "Title Page", "Usage guide (bilingual)"],
    },
}

# ---------------------------------------------------------------------------
# Pemilih bahasa
# ---------------------------------------------------------------------------
if "ui_lang" not in st.session_state:
    st.session_state["ui_lang"] = "id"
_c1, _c2 = st.columns([3, 1])
with _c2:
    _choice = st.radio("🌐", ["Indonesia", "English"], index=0 if st.session_state["ui_lang"] == "id" else 1,
                       horizontal=True, label_visibility="collapsed", key="ui_lang_radio")
    st.session_state["ui_lang"] = "id" if _choice == "Indonesia" else "en"
LG = st.session_state["ui_lang"]
t = T[LG]

st.title(t["title"])
st.caption(t["sub"])
st.markdown(f'<div class="note">{t["optional"]}</div>', unsafe_allow_html=True)

# --- Petunjuk wajib dibaca -------------------------------------------------
st.markdown(f'<div class="must">{t["read_first"]}</div>', unsafe_allow_html=True)
_g = TPL / GUIDE
if _g.exists():
    st.download_button(t["guide_btn"], _g.read_bytes(), GUIDE, MIME, key="guide_top")

up = st.file_uploader(t["uploader"], type=["docx"], accept_multiple_files=False, help=t["uploader_help"], key="manuscript_up")


@st.cache_data(show_spinner=False, max_entries=8)
def _parse(h: str, data: bytes):
    return P.parse_docx(data)


def _fname(M, suffix):
    sn = B.surname(M["authors"][0]["name"]) if M["authors"] else "Manuscript"
    sn = re.sub(r"[^A-Za-z0-9]+", "", B._strip_acc(sn)) or "Manuscript"
    return f"{sn}_{suffix}.docx"


def _translate(text: str, target: str):
    """Terjemahan judul lewat Claude API bila ANTHROPIC_API_KEY tersedia; None bila tidak."""
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        key = None
    if not key:
        return None
    try:
        import anthropic
        cl = anthropic.Anthropic(api_key=key)
        lang = "English" if target == "en" else "Indonesian"
        model = st.secrets.get("ANTHROPIC_MODEL", "claude-haiku-5-5")
        r = cl.messages.create(model=model, max_tokens=300, messages=[{
            "role": "user",
            "content": f"Translate this scientific article title into {lang}. Keep scientific names, acronyms and proper nouns unchanged. "
                       f"Return only the translated title, in Title Case, without quotes.\n\n{text}"}])
        out = "".join(b.text for b in r.content if getattr(b, "type", "") == "text").strip().strip('"“”')
        return out or None
    except Exception:
        return None


if up is not None:
    data = up.getvalue()
    h = hashlib.sha1(data).hexdigest()
    try:
        with st.spinner(t["reading"]):
            M0, rep = _parse(h, data)
    except Exception as e:
        st.error(t["bad_file"] + str(e))
        st.stop()

    M = {**M0, "authors": [dict(a) for a in M0["authors"]], "closing": dict(M0["closing"])}

    # ---------------- 1. hasil pembacaan
    st.subheader(t["s1"])
    f = dict(rep["found"])
    f["lang"] = t["lang_en"] if f["lang"] == "en" else t["lang_id"]
    keys = ["title_id", "title_en", "lang", "authors", "affs", "abstract", "keywords", "sections", "tables", "figures", "equations", "refs", "cites"]
    c1, c2 = st.columns(2)
    half = (len(keys) + 1) // 2
    for col, part in ((c1, keys[:half]), (c2, keys[half:])):
        with col:
            for k in part:
                v = str(f[k]) if str(f[k]) else "—"
                st.markdown(f"**{t['f_' + k]}**  \n{v if len(v) < 110 else v[:107] + '…'}")

    probs = [P.pick(m, LG) for m in rep["problems"]]
    if probs:
        st.warning(t["attention"] + "\n\n" + "\n".join(f"- {p}" for p in probs))
    else:
        st.success(t["all_ok"])
    for w in rep["warnings"]:
        st.info(P.pick(w, LG))
    for w in rep["infos"]:
        st.caption("ℹ️ " + P.pick(w, LG))

    with st.expander(t["outline"]):
        for k, lab in zip(B.SECTION_KEYS, B.LBL[M["lang"]]["sec"]):
            blocks = M["sections"][k]
            st.markdown(f"**{lab}** – {sum(1 for b in blocks if b['t'] in ('p', 'eq'))} {t['para']}, "
                        f"{sum(1 for b in blocks if b['t']=='table')} {t['tbl']}, {sum(1 for b in blocks if b['t']=='fig')} {t['fig']}")
            for s in [b for b in blocks if b["t"] in ("h2", "h3")]:
                st.caption(("    " if s["t"] == "h3" else "") + "• " + B.plain_text(s["c"]))
        st.markdown(t["refs3"])
        for r in M["refs"][:3]:
            st.caption(B.plain_text(r))
        if M["authors"]:
            st.markdown(t["authors"])
            for i, a in enumerate(M["authors"]):
                st.caption(f"{i+1}. {a['name']}{' ★' if i == M['corr'] else ''} — {t['aff_lbl']} {a['affs'] or '—'}; {t['email_lbl']} {a['email'] or '—'}")

    # ---------------- 2. judul dua bahasa + judul singkat
    st.subheader(t["s_titles"])
    st.caption(t["titles_note"])
    ver = h[:8]
    for key, tgt, other in (("title_id", "id", "title_en"), ("title_en", "en", "title_id")):
        if not B.plain_text(M[key]).strip():
            st.warning(t["missing_title"].format(x=t["x_" + tgt]))
            src = B.plain_text(M[other]).strip()
            sk = f"gen_{key}_{ver}"
            if src and st.button(t["gen_btn"] + f" ({t['x_' + tgt]})", key=f"btn_{key}_{ver}"):
                with st.spinner(t["gen_wait"]):
                    g = _translate(src, tgt)
                if g:
                    st.session_state[sk] = g
                    st.session_state[f"in_{key}_{ver}"] = g
                else:
                    st.session_state[sk + "_fail"] = True
            if st.session_state.get(sk + "_fail"):
                st.caption(t["gen_fail"] + " " + t["gen_hint"])
            M[key] = st.text_input(t["title_" + tgt], value=st.session_state.get(sk, ""),
                                   key=f"in_{key}_{ver}")
    M["running"] = st.text_input(t["running"], value=M["running"], max_chars=60, key=f"run_{ver}")
    if B.plain_text(M["title_id"]).strip() and B.plain_text(M["title_en"]).strip() and M["running"].strip():
        st.success(t["title_ok"])

    # ---------------- 3. data penulis
    st.subheader(t["s_auth"])
    if M["authors"]:
        st.success(t["auth_found"].format(n=len(M["authors"]), a=len(M["affs"])))
        container = st.expander(t["auth_corr"])
    else:
        st.info(t["auth_none"])
        container = st.container()
    with container:
        tab_up, tab_form = st.tabs([t["tab_up"], t["tab_form"]])
    with tab_up:
        up2 = st.file_uploader(t["full_up"], type=["docx"], key="full_up", help=t["full_help"])
        if up2 is not None:
            try:
                M2, _ = _parse(hashlib.sha1(up2.getvalue()).hexdigest(), up2.getvalue())
                if M2["authors"]:
                    for k in ("authors", "affs", "corr", "phone"):
                        M[k] = M2[k]
                    for k in ("coi", "fund", "ack", "data", "ethics"):
                        if M2["closing"].get(k) and not M["closing"].get(k):
                            M["closing"][k] = M2["closing"][k]
                    if M2["ai_use"] and not M["ai_use"]:
                        M["ai_use"], M["ai_text"] = M2["ai_use"], M2["ai_text"]
                    st.success(t["read_n"].format(n=len(M2["authors"]), a=len(M2["affs"])))
                else:
                    st.warning(t["no_auth_file"])
            except Exception as e:
                st.error(t["bad_file"] + str(e))
    with tab_form:
        import pandas as pd
        cn, ce, co, ca = t["col_name"], t["col_email"], t["col_orcid"], t["col_aff"]
        base = pd.DataFrame([{cn: a["name"], ce: a["email"], co: a["orcid"], ca: ",".join(map(str, a["affs"]))} for a in M["authors"]]
                            or [{cn: "", ce: "", co: "", ca: "1"}])
        ed = st.data_editor(base, num_rows="dynamic", use_container_width=True, key="ed_auth")
        affd = st.data_editor(pd.DataFrame({t["aff_col"]: M["affs"] or [""]}), num_rows="dynamic", use_container_width=True, key="ed_aff")
        rows = [r for r in ed.to_dict("records") if str(r.get(cn) or "").strip()]
        stt = st.session_state.get("ed_auth", {})
        if rows and (not M["authors"] or stt.get("edited_rows") or stt.get("added_rows")):
            old = {a["name"]: a for a in M["authors"]}
            M["authors"] = []
            for r in rows:
                nm = str(r[cn]).strip()
                nums = [int(x) for x in re.findall(r"\d+", str(r.get(ca) or ""))]
                M["authors"].append({"name": nm, "email": str(r.get(ce) or "").strip(), "orcid": str(r.get(co) or "").strip(),
                                     "affs": nums, "roles": old.get(nm, {}).get("roles", [])})
            M["affs"] = [str(x).strip() for x in affd.iloc[:, 0].tolist() if str(x).strip() and str(x) != "nan"]
        if M["authors"]:
            st.markdown(t["roles"])
            for i, a in enumerate(M["authors"]):
                a["roles"] = st.multiselect(a["name"] or f"Author {i+1}", B.CREDIT, default=a["roles"], key=f"roles{i}_{a['name']}")

    st.caption(t["tpl_auto"])
    with st.expander(t["extra"]):
        M["ojs_id"] = st.text_input(t["ojs"], value="", placeholder="PLATAX-2026-XXXX", key="ojs_id")
        if len(M["authors"]) > 1:
            M["corr"] = st.selectbox(t["corr"], range(len(M["authors"])), index=min(M["corr"], len(M["authors"]) - 1),
                                     format_func=lambda i: M["authors"][i]["name"], key="corr_sel")
        M["phone"] = st.text_input(t["phone"], value=M.get("phone", ""), key="phone")
        M["lang"] = st.radio(t["mslang"], ["id", "en"], index=0 if M["lang"] == "id" else 1, horizontal=True,
                             format_func=lambda x: "Indonesia" if x == "id" else "English", key="ms_lang")

    # ---------------- 4. unduh
    st.subheader(t["s_dl"])
    st.markdown(t["dl_note"])
    try:
        d_blind = B.build("blind", M)
        d_tp = B.build("title", M)
    except Exception as e:
        st.error(t["dl_fail"] + str(e))
        st.stop()
    st.download_button(t["dl_blind"], d_blind, _fname(M, "Manuscript_BlindReview"), MIME, type="primary", use_container_width=True)
    st.download_button(t["dl_tp"], d_tp, _fname(M, "TitlePage"), MIME, use_container_width=True)
    if not M["authors"] or not all(a["email"] for a in M["authors"]) or not all(a["roles"] for a in M["authors"]) or not M["affs"]:
        st.caption(t["tp_ph"])
    st.caption(t["check_word"])

st.divider()
st.subheader(t["backup"])
st.caption(t["backup_sub"])
LIST = ["Template_Artikel_PLATAX_2026_ID.docx", "Template_Article_PLATAX_2026_EN.docx", "Template_Naskah_BlindReview_JIP_2026.docx",
        "Template_Manuscript_BlindReview_JIP_2026.docx", "Template_Title_Page_JIP_2026.docx", GUIDE]
cols = st.columns(2)
for i, (lab, fn) in enumerate(zip(t["tpl_names"], LIST)):
    p = TPL / fn
    if p.exists():
        cols[i % 2].download_button(lab, p.read_bytes(), fn, MIME, key=f"tpl{i}", use_container_width=True)
st.caption(t["privacy"])
