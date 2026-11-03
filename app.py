"""
Penata Naskah JIP / PLATAX 2026 – versi unggah.
Penulis cukup mengunggah naskah .docx; aplikasi menerapkan templat resmi. Tanpa mengetik atau menyalin.
Jalankan:  streamlit run app.py
"""
import hashlib
import re
from pathlib import Path

import streamlit as st

import naskah_builder as B
import naskah_parser as P

HERE = Path(__file__).resolve().parent
TPL = HERE / "templates"

st.set_page_config(page_title="Penata Naskah JIP 2026", page_icon="📄", layout="centered")

st.markdown(
    """
<style>
.block-container{max-width:760px;padding-top:2rem}
h1{font-size:1.6rem!important;line-height:1.25!important}
.note{border-left:4px solid #007BB8;background:rgba(0,123,184,.08);padding:.7rem .9rem;border-radius:0 6px 6px 0;font-size:.93rem;margin:.4rem 0 1rem}
.ok{color:#1a7f37}.bad{color:#b35900}
</style>
""",
    unsafe_allow_html=True,
)

st.title("Penata Naskah Jurnal Ilmiah PLATAX")
st.caption("Unggah naskah .docx Anda. Aplikasi hanya menerapkan tata letak templat JIP 2026 pada isi naskah Anda.")

st.markdown(
    '<div class="note"><b>Alat ini bersifat opsional.</b> Anda tetap boleh menyerahkan naskah memakai templat Word manual (lihat bagian bawah). '
    "Pada tahap awal, kerapian format <b>tidak</b> menjadi alasan penolakan naskah. Redaksi menilai isi lebih dulu.</div>",
    unsafe_allow_html=True,
)

up = st.file_uploader("Unggah naskah (.docx)", type=["docx"], accept_multiple_files=False,
                      help="Satu berkas .docx berisi judul, penulis, abstrak, isi naskah, dan daftar pustaka. Format lama (.doc) harap disimpan ulang sebagai .docx.")


@st.cache_data(show_spinner=False, max_entries=8)
def _parse(h: str, data: bytes):
    return P.parse_docx(data)


def _fname(M, suffix):
    sn = B.surname(M["authors"][0]["name"]) if M["authors"] else "Naskah"
    sn = re.sub(r"[^A-Za-z0-9]+", "", B._strip_acc(sn)) or "Naskah"
    return f"{sn}_{suffix}.docx"


if up is not None:
    data = up.getvalue()
    h = hashlib.sha1(data).hexdigest()
    try:
        with st.spinner("Membaca naskah…"):
            M0, rep = _parse(h, data)
    except Exception as e:  # berkas rusak / bukan docx
        st.error("Berkas tidak dapat dibaca. Pastikan berkas adalah .docx yang valid (bukan .doc yang diganti namanya). " f"Rincian teknis: {e}")
        st.stop()

    M = {**M0, "authors": [dict(a) for a in M0["authors"]], "closing": dict(M0["closing"])}

    st.subheader("1. Periksa hasil pembacaan")
    f = rep["found"]
    c1, c2 = st.columns(2)
    items = list(f.items())
    half = (len(items) + 1) // 2
    for col, part in ((c1, items[:half]), (c2, items[half:])):
        with col:
            for k, v in part:
                v = str(v) if str(v) else "—"
                st.markdown(f"**{k}**  \n{v if len(v) < 110 else v[:107] + '…'}")

    if rep["problems"]:
        st.warning("Perlu diperhatikan (tidak menghalangi pembuatan berkas):\n\n" + "\n".join(f"- {p}" for p in rep["problems"]))
    else:
        st.success("Semua bagian utama terbaca.")
    for w in rep["warnings"]:
        st.info(w)
    for w in rep["infos"]:
        st.caption("ℹ️ " + w)

    with st.expander("Lihat kerangka yang terbaca"):
        for k, lab in zip(B.SECTION_KEYS, B.LBL[M["lang"]]["sec"]):
            blocks = M["sections"][k]
            subs = [b for b in blocks if b["t"] in ("h2", "h3")]
            st.markdown(f"**{lab}** – {sum(1 for b in blocks if b['t']=='p')} paragraf, "
                        f"{sum(1 for b in blocks if b['t']=='table')} tabel, {sum(1 for b in blocks if b['t']=='fig')} gambar")
            for s in subs:
                st.caption(("    " if s["t"] == "h3" else "") + "• " + B.plain_text(s["c"]))
        st.markdown("**Daftar pustaka (3 pertama)**")
        for r in M["refs"][:3]:
            st.caption(B.plain_text(r))
        if M["authors"]:
            st.markdown("**Penulis**")
            for i, a in enumerate(M["authors"]):
                st.caption(f"{i+1}. {a['name']}{' ★' if i == M['corr'] else ''} — afiliasi {a['affs'] or '—'}; e-mail {a['email'] or '—'}")

    st.subheader("2. Data penulis untuk Title Page")
    if not M["authors"]:
        st.info("Naskah ini tidak memuat nama penulis (blind review). Title Page memerlukan data penulis. "
                "Lengkapi dengan salah satu cara di bawah.")
    else:
        st.caption("Data penulis sudah terbaca dari naskah. Anda boleh menimpanya dengan cara di bawah.")
    tab_up, tab_form = st.tabs(["Unggah naskah lengkap (ada nama penulis)", "Isi singkat di sini"])
    with tab_up:
        up2 = st.file_uploader("Naskah versi lengkap (.docx)", type=["docx"], key="full_up",
                               help="Versi naskah yang memuat nama, afiliasi, dan e-mail penulis. Hanya data penulis yang diambil.")
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
                    st.success(f"Terbaca {len(M2['authors'])} penulis dan {len(M2['affs'])} afiliasi.")
                else:
                    st.warning("Nama penulis tidak ditemukan di berkas ini.")
            except Exception as e:
                st.error(f"Berkas tidak dapat dibaca: {e}")
    with tab_form:
        import pandas as pd
        n0 = max(len(M["authors"]), 1)
        base = pd.DataFrame([{"Nama": a["name"], "E-mail": a["email"], "ORCID": a["orcid"],
                              "No. afiliasi": ",".join(map(str, a["affs"]))} for a in M["authors"]]
                            or [{"Nama": "", "E-mail": "", "ORCID": "", "No. afiliasi": "1"}])
        ed = st.data_editor(base, num_rows="dynamic", use_container_width=True, key="ed_auth")
        affd = st.data_editor(pd.DataFrame({"Afiliasi (nomor 1, 2, … sesuai urutan)": M["affs"] or [""]}),
                              num_rows="dynamic", use_container_width=True, key="ed_aff")
        rows = [r for r in ed.to_dict("records") if str(r.get("Nama") or "").strip()]
        if rows and (not M["authors"] or st.session_state.get("ed_auth", {}).get("edited_rows") or st.session_state.get("ed_auth", {}).get("added_rows")):
            old = {a["name"]: a for a in M["authors"]}
            M["authors"] = []
            for r in rows:
                nm = str(r["Nama"]).strip()
                nums = [int(x) for x in re.findall(r"\d+", str(r.get("No. afiliasi") or ""))]
                M["authors"].append({"name": nm, "email": str(r.get("E-mail") or "").strip(),
                                     "orcid": str(r.get("ORCID") or "").strip(), "affs": nums,
                                     "roles": old.get(nm, {}).get("roles", [])})
            M["affs"] = [str(x).strip() for x in affd.iloc[:, 0].tolist() if str(x).strip() and str(x) != "nan"]
        if M["authors"]:
            st.markdown("**Peran kontribusi (CRediT)**")
            for i, a in enumerate(M["authors"]):
                a["roles"] = st.multiselect(a["name"] or f"Penulis {i+1}", B.CREDIT, default=a["roles"], key=f"roles{i}_{a['name']}")

    with st.expander("Pengaturan tambahan (opsional)"):
        M["ojs_id"] = st.text_input("ID Naskah OJS (untuk Title Page)", value="", placeholder="PLATAX-2026-XXXX")
        if len(M["authors"]) > 1:
            M["corr"] = st.selectbox("Penulis korespondensi", range(len(M["authors"])), index=min(M["corr"], len(M["authors"]) - 1),
                                     format_func=lambda i: M["authors"][i]["name"])
        M["phone"] = st.text_input("Telepon penulis korespondensi (opsional)", value="")
        M["running"] = st.text_input("Judul singkat (maks. 60 karakter)", value=M["running"], max_chars=60)
        M["lang"] = st.radio("Bahasa naskah", ["id", "en"], index=0 if M["lang"] == "id" else 1, horizontal=True,
                             format_func=lambda x: "Indonesia" if x == "id" else "English")

    st.subheader("3. Unduh berkas")
    st.markdown("Dua berkas untuk OJS dibuat dari templat resmi. Seluruh isi tetap milik Anda; hanya tata letak yang diterapkan.")
    try:
        d_blind = B.build("blind", M)
        d_tp = B.build("title", M)
    except Exception as e:
        st.error(f"Berkas gagal dibuat: {e}")
        st.stop()
    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    st.download_button("⬇️  Naskah blind review (.docx) – untuk berkas utama OJS", d_blind, _fname(M, "Naskah_BlindReview"), MIME,
                       type="primary", use_container_width=True)
    st.download_button("⬇️  Title Page (.docx) – berkas tambahan terpisah", d_tp, _fname(M, "TitlePage"), MIME, use_container_width=True)
    if not M["authors"]:
        st.caption("Data penulis belum diisi, sehingga Title Page berisi penanda [Penulis mengisi …] berwarna abu-abu. "
                   "Isi langsung di Word, atau lengkapi data pada bagian 2 lalu unduh ulang.")
    st.caption("Buka hasilnya di Microsoft Word dan periksa sekali sebelum mengunggah ke OJS. Jika ada bagian yang keliru terbaca, "
               "perbaiki langsung di Word atau di naskah asli lalu unggah ulang.")

st.divider()
st.subheader("Cadangan: templat Word manual")
st.caption("Jika alat ini tidak cocok untuk naskah Anda, pakai templat berikut secara langsung.")
MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
LIST = [
    ("Artikel (Indonesia)", "Template_Artikel_PLATAX_2026_ID.docx"),
    ("Article (English)", "Template_Article_PLATAX_2026_EN.docx"),
    ("Naskah blind review (Indonesia)", "Template_Naskah_BlindReview_JIP_2026.docx"),
    ("Manuscript blind review (English)", "Template_Manuscript_BlindReview_JIP_2026.docx"),
    ("Title Page", "Template_Title_Page_JIP_2026.docx"),
    ("Petunjuk penggunaan (dwibahasa)", "Petunjuk_Penggunaan_Template_JIP_2026_Bilingual_v4.docx"),
]
cols = st.columns(2)
for i, (lab, fn) in enumerate(LIST):
    p = TPL / fn
    if p.exists():
        cols[i % 2].download_button(lab, p.read_bytes(), fn, MIME, key=f"tpl{i}", use_container_width=True)

st.caption("Naskah yang Anda unggah hanya diproses di memori selama sesi dan tidak disimpan oleh aplikasi ini.")
