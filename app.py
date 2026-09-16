import streamlit as st

st.set_page_config(page_title="Manuscript Submission", layout="wide")

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .section-header {
        font-size: 1.2rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
    }
    .table-header {
        font-weight: 600;
        font-size: 0.85rem;
        color: #aaa;
        padding: 4px 0;
        border-bottom: 1px solid #444;
        margin-bottom: 4px;
    }
    .author-row {
        border-bottom: 1px solid #2a2a2a;
        padding: 4px 0;
    }
    div[data-testid="stCheckbox"] label { font-size: 0.85rem; }
    div[data-testid="stTextInput"] input { font-size: 0.85rem; }
    .affil-note { font-size: 0.8rem; color: #888; margin-top: 0.3rem; }
    .credit-role-grid { display: grid; gap: 6px; }
</style>
""", unsafe_allow_html=True)

# ── CRediT Roles ──────────────────────────────────────────────────────────────
CREDIT_ROLES = [
    "Conceptualization",
    "Data curation",
    "Formal analysis",
    "Funding acquisition",
    "Investigation",
    "Methodology",
    "Project administration",
    "Resources",
    "Software",
    "Supervision",
    "Validation",
    "Visualization",
    "Writing – original draft",
    "Writing – review & editing",
]

# ── Session State Init ────────────────────────────────────────────────────────
if "authors" not in st.session_state:
    st.session_state.authors = [
        {
            "nama": "Febry S. I. Menajang",
            "singkatan": "F.S.I. Menajang",
            "afiliasi": "a",
            "email": "",
            "korespondensi": True,
            "orcid": "0000-0000-0000-0000",
            "contrib_roles": [],
            "contrib_tambahan": "",
        }
    ]

if "affil_details" not in st.session_state:
    st.session_state.affil_details = [{"label": "a", "detail": ""}]


def add_author():
    letters = "abcdefghijklmnopqrstuvwxyz"
    n = len(st.session_state.authors)
    label = letters[n] if n < 26 else f"a{n}"
    st.session_state.authors.append(
        {
            "nama": "",
            "singkatan": "",
            "afiliasi": label,
            "email": "",
            "korespondensi": False,
            "orcid": "",
            "contrib_roles": [],
            "contrib_tambahan": "",
        }
    )
    st.session_state.affil_details.append({"label": label, "detail": ""})


def remove_author(idx):
    if len(st.session_state.authors) > 1:
        st.session_state.authors.pop(idx)


def auto_singkatan(nama: str) -> str:
    """'Febry S. I. Menajang' → 'F.S.I. Menajang'"""
    parts = [p for p in nama.strip().split() if p]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0]
    surname = parts[-1]
    initials = "".join(p[0].upper() + "." for p in parts[:-1])
    return f"{initials} {surname}"


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Authors and Affiliations
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("### 2. Penulis dan Afiliasi / Authors and Affiliations")

# Table header
h = st.columns([3, 1.5, 3, 1, 2.5, 1])
h[0].markdown('<div class="table-header">Nama</div>', unsafe_allow_html=True)
h[1].markdown('<div class="table-header">Afiliasi (huruf: a, b, …)</div>', unsafe_allow_html=True)
h[2].markdown('<div class="table-header">Email</div>', unsafe_allow_html=True)          # ← FIX 1
h[3].markdown('<div class="table-header">Korespondensi ★</div>', unsafe_allow_html=True)
h[4].markdown('<div class="table-header">ORCID iD</div>', unsafe_allow_html=True)
h[5].markdown('<div class="table-header"></div>', unsafe_allow_html=True)

# Author rows
for i, author in enumerate(st.session_state.authors):
    cols = st.columns([3, 1.5, 3, 1, 2.5, 1])

    author["nama"] = cols[0].text_input(
        "Nama", value=author["nama"], key=f"nama_{i}", label_visibility="collapsed",
        placeholder="Nama lengkap penulis"
    )
    author["afiliasi"] = cols[1].text_input(
        "Afiliasi", value=author["afiliasi"], key=f"afil_{i}", label_visibility="collapsed",
        placeholder="a"
    )
    author["email"] = cols[2].text_input(                                               # ← FIX 1
        "Email", value=author["email"], key=f"email_{i}", label_visibility="collapsed",
        placeholder="email@institusi.ac.id"
    )
    author["korespondensi"] = cols[3].checkbox(
        "★", value=author["korespondensi"], key=f"koresp_{i}"
    )
    author["orcid"] = cols[4].text_input(
        "ORCID", value=author["orcid"], key=f"orcid_{i}", label_visibility="collapsed",
        placeholder="0000-0000-0000-0000"
    )
    if cols[5].button("🗑️", key=f"del_{i}", help="Hapus penulis ini"):
        remove_author(i)
        st.rerun()

# Add author button
st.button("＋ Tambah Penulis", on_click=add_author)

# Affiliation Details
st.markdown("**Keterangan Afiliasi / Affiliation Details:**")
for i, affil in enumerate(st.session_state.affil_details):
    c1, c2 = st.columns([0.5, 9.5])
    c1.markdown(f"**{affil['label']}**")
    affil["detail"] = c2.text_input(
        f"Afiliasi {affil['label']}",
        value=affil["detail"],
        key=f"affil_detail_{i}",
        label_visibility="collapsed",
        placeholder="Nama institusi, kota, negara"
    )

st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Author Contributions (CRediT)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("### 4. Kontribusi Penulis / Author Contributions (CRediT)")
st.caption("Pilih peran untuk setiap penulis.")

for i, author in enumerate(st.session_state.authors):
    nama_full    = author["nama"].strip() if author["nama"].strip() else f"Penulis {i + 1}"
    # Auto-generate singkatan if empty
    if not author.get("singkatan"):
        author["singkatan"] = auto_singkatan(author["nama"])
    singkatan_val = author["singkatan"]

    # Expander label: nomor · nama lengkap · (singkatan)
    label = f"📄  {nama_full}"
    if singkatan_val:
        label += f"  —  {singkatan_val}"

    with st.expander(label, expanded=True):

        # ── Singkatan nama (editable) ───────────────────────────────────────
        c_singkat, c_hapus = st.columns([10, 1])
        author["singkatan"] = c_singkat.text_input(
            "Singkatan Nama / Abbreviated Name",
            value=singkatan_val,
            key=f"singkatan_{i}",
            placeholder="misal: F.S.I. Menajang",
            help="Nama singkat yang akan muncul pada pernyataan kontribusi, "
                 "misal: F.S.I. Menajang · Menajang FSI · Menajang et al.",
        )
        # Hapus penulis dari dalam expander
        if c_hapus.button("🗑️", key=f"del_credit_{i}", help="Hapus penulis ini"):
            remove_author(i)
            st.rerun()

        st.markdown("---")

        # ── CRediT checkboxes (3 kolom) ────────────────────────────────────
        role_cols   = st.columns(3)
        checked_roles = []
        for j, role in enumerate(CREDIT_ROLES):
            col     = role_cols[j % 3]
            current = role in author.get("contrib_roles", [])
            if col.checkbox(role, value=current, key=f"role_{i}_{j}"):
                checked_roles.append(role)
        author["contrib_roles"] = checked_roles

        # ── Tambahan ───────────────────────────────────────────────────────
        author["contrib_tambahan"] = st.text_input(
            "Tambahan (pisah koma) / Additional roles (comma-separated)",
            value=author.get("contrib_tambahan", ""),
            key=f"tambahan_{i}",
            placeholder="misal: Project administration",
        )

        # ── Pratinjau baris kontribusi ─────────────────────────────────────
        all_roles = list(author["contrib_roles"])
        if author.get("contrib_tambahan"):
            all_roles += [r.strip() for r in author["contrib_tambahan"].split(",") if r.strip()]
        if all_roles:
            st.caption(
                f"**Output:** {author['singkatan'] or nama_full}: "
                + ", ".join(all_roles) + "."
            )

# ── Tombol tambah penulis (di dalam section CRediT) ──────────────────────────
st.button(
    "＋  Tambah Penulis / Add Author",
    on_click=add_author,
    key="add_author_credit",
    help="Menambahkan baris penulis baru di sini dan di Bagian 2",
)

st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════════
# Summary / Preview
# ═══════════════════════════════════════════════════════════════════════════════
with st.expander("📋 Pratinjau / Preview"):
    st.markdown("**Penulis & Afiliasi:**")
    for a in st.session_state.authors:
        koresp = " ★" if a["korespondensi"] else ""
        st.markdown(
            f"- **{a['nama']}** ({a['afiliasi']}){koresp}  \n"
            f"  Email: {a['email'] or '—'} | ORCID: {a['orcid'] or '—'}"
        )
    st.markdown("**Kontribusi:**")
    for a in st.session_state.authors:
        nama = a["nama"].strip() or "—"
        roles = a.get("contrib_roles", [])
        if a.get("contrib_tambahan"):
            roles += [r.strip() for r in a["contrib_tambahan"].split(",") if r.strip()]
        st.markdown(f"- **{nama}**: {', '.join(roles) if roles else '—'}")
