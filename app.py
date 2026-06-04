import io
import datetime
import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from trial_plan     import PLAN_LABELS,    resolve_plan_columns,       build_trial_plan_zip
from schedule_plan  import ACTIVITY_DEFS,  resolve_activity_columns,   build_schedule_zip
from treatment_plan import TREATMENT_HEADERS, resolve_treatment_columns, build_treatment_zip

# ═════════════════════════════════════════════
# PAGE CONFIG
# ═════════════════════════════════════════════
st.set_page_config(
    page_title="Acacia's Trial Tools Automation",
    page_icon="🌳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═════════════════════════════════════════════
# CUSTOM CSS — Elegant light theme + forced light mode
# ═════════════════════════════════════════════
st.markdown("""
<style>
    /* ════ Force light theme regardless of system ════ */
    :root {
        color-scheme: light only;
    }
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background: #fbfcfb !important;
        color: #1a1a1a !important;
    }
    [data-testid="stHeader"] {
        background: transparent !important;
    }
    [data-testid="stSidebar"] {
        background: #ffffff !important;
        border-right: 1px solid #ebefeb !important;
    }

    /* ════ Import elegant font ════ */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, sans-serif !important;
    }

    /* ════ Layout ════ */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 4rem;
        max-width: 1380px;
    }
    #MainMenu, footer, [data-testid="stDecoration"] { visibility: hidden; }

    /* ════ Hero header ════ */
    .hero {
        position: relative;
        background: linear-gradient(120deg, #14532d 0%, #166534 45%, #22c55e 130%);
        padding: 2.6rem 2.8rem;
        border-radius: 24px;
        color: white;
        margin-bottom: 2rem;
        overflow: hidden;
        box-shadow: 0 20px 50px -15px rgba(20, 83, 45, 0.45);
    }
    .hero::before {
        content: "";
        position: absolute;
        top: -60%; right: -8%;
        width: 380px; height: 380px;
        background: radial-gradient(circle, rgba(255,255,255,0.14) 0%, transparent 70%);
        border-radius: 50%;
    }
    .hero::after {
        content: "🌿";
        position: absolute;
        right: 2.5rem; bottom: -1rem;
        font-size: 7rem;
        opacity: 0.12;
        transform: rotate(-12deg);
    }
    .hero h1 {
        font-size: 2.3rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -1px;
        position: relative;
        z-index: 1;
    }
    .hero p {
        font-size: 1.05rem;
        opacity: 0.9;
        margin: 0.6rem 0 0 0;
        font-weight: 300;
        max-width: 640px;
        position: relative;
        z-index: 1;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        backdrop-filter: blur(8px);
        border: 1px solid rgba(255,255,255,0.25);
        padding: 0.3rem 0.9rem;
        border-radius: 30px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 0.5px;
        margin-bottom: 0.9rem;
        position: relative;
        z-index: 1;
    }

    /* ════ Section title ════ */
    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #14532d;
        margin: 1.2rem 0 0.9rem 0;
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    .section-title .num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px; height: 28px;
        background: #dcfce7;
        color: #166534;
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 800;
    }

    /* ════ Metric cards ════ */
    .metric-card {
        background: white;
        border: 1px solid #edf1ed;
        border-radius: 16px;
        padding: 1.25rem 1.4rem;
        box-shadow: 0 2px 12px rgba(20,83,45,0.04);
        transition: all 0.2s cubic-bezier(.4,0,.2,1);
        position: relative;
        overflow: hidden;
    }
    .metric-card::before {
        content: "";
        position: absolute;
        left: 0; top: 0; bottom: 0;
        width: 4px;
        background: linear-gradient(180deg, #22c55e, #15803d);
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 28px rgba(20,83,45,0.12);
        border-color: #c8e6c9;
    }
    .metric-card .icon {
        font-size: 1.3rem;
        margin-bottom: 0.4rem;
    }
    .metric-card .label {
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #8a948a;
        font-weight: 700;
    }
    .metric-card .value {
        font-size: 1.9rem;
        font-weight: 800;
        color: #14532d;
        margin-top: 0.15rem;
        letter-spacing: -0.5px;
        line-height: 1.1;
    }

    /* ════ Buttons ════ */
    .stButton > button {
        border-radius: 12px;
        font-weight: 600;
        padding: 0.7rem 1.2rem;
        border: none;
        transition: all 0.2s ease;
        font-size: 0.95rem;
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(120deg, #166534, #22c55e);
        box-shadow: 0 6px 18px -6px rgba(34,197,94,0.6);
    }
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px -6px rgba(34,197,94,0.7);
    }
    div[data-testid="stDownloadButton"] > button {
        border-radius: 12px;
        font-weight: 600;
        background: linear-gradient(120deg, #15803d, #16a34a);
        border: none;
        box-shadow: 0 6px 18px -6px rgba(22,163,74,0.6);
    }
    div[data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px);
    }

    /* ════ Tabs ════ */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #f1f6f1;
        padding: 7px;
        border-radius: 16px;
        border: 1px solid #e7eee7;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 11px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        background: transparent;
        color: #5a6b5a;
        transition: all 0.18s ease;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #166534;
    }
    .stTabs [aria-selected="true"] {
        background: white !important;
        color: #14532d !important;
        box-shadow: 0 3px 10px rgba(20,83,45,0.1);
    }
    .stTabs [data-baseweb="tab-highlight"],
    .stTabs [data-baseweb="tab-border"] { display: none; }

    /* ════ Tool description card ════ */
    .tool-desc {
        background: linear-gradient(135deg, #ffffff 0%, #f7fbf7 100%);
        border: 1px solid #e4ede4;
        border-radius: 16px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1.4rem;
        transition: all 0.2s ease;
        height: 100%;
    }
    .tool-desc:hover {
        border-color: #b6e0b6;
        box-shadow: 0 8px 24px rgba(20,83,45,0.07);
    }
    .tool-desc .tool-icon {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 46px; height: 46px;
        background: #dcfce7;
        border-radius: 12px;
        font-size: 1.4rem;
        margin-bottom: 0.8rem;
    }
    .tool-desc h4 {
        margin: 0 0 0.45rem 0;
        color: #14532d;
        font-size: 1.1rem;
        font-weight: 700;
    }
    .tool-desc p {
        margin: 0;
        color: #5d685d;
        font-size: 0.92rem;
        line-height: 1.55;
    }

    /* ════ Columns chips ════ */
    .col-chips {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin: 0.4rem 0 1rem 0;
    }
    .col-chip {
        background: #f1f6f1;
        border: 1px solid #e0eae0;
        color: #2d4a2d;
        padding: 0.28rem 0.7rem;
        border-radius: 8px;
        font-size: 0.78rem;
        font-family: 'SF Mono', 'Consolas', monospace;
        font-weight: 500;
    }

    /* ════ Alerts polish ════ */
    [data-testid="stAlert"] {
        border-radius: 12px;
        border: none;
    }

    /* ════ Dataframe ════ */
    [data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid #ebefeb;
    }

    /* ════ Expander ════ */
    [data-testid="stExpander"] {
        border: 1px solid #ebefeb !important;
        border-radius: 14px !important;
        background: white;
    }

    /* ════ Sidebar polish ════ */
    [data-testid="stSidebar"] .sidebar-brand {
        font-size: 1.3rem;
        font-weight: 800;
        color: #14532d;
        margin-bottom: 0.1rem;
    }

    /* ════ Footer ════ */
    .footer {
        text-align: center;
        color: #a3aaa3;
        font-size: 0.82rem;
        padding: 2rem 0 0.5rem 0;
        border-top: 1px solid #eef2ee;
        margin-top: 2.5rem;
    }
    .footer b { color: #166534; }

    /* ════ Generate panel ════ */
    .gen-panel {
        background: #f7fbf7;
        border: 1px dashed #c8e0c8;
        border-radius: 14px;
        padding: 1.2rem 1.4rem;
        margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ═════════════════════════════════════════════
# CONSTANTS
# ═════════════════════════════════════════════
NOT_YET    = {"not yet register", "not yet registered"}
HEADER_ROW = 3

REQUIRED_COLS = [
    "Code Trial", "Sector", "Comp", "Year", "Register", "Established by",
]


# ═════════════════════════════════════════════
# DATABASE LOADER — cached per uploaded file
# ═════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_database(file_bytes: bytes):
    wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)

    if "Trial Register" not in wb.sheetnames:
        return None, None, None, None, None, None, "Sheet 'Trial Register' not found."

    ws = wb["Trial Register"]

    sheet_headers: dict[str, list[int]] = {}
    for col in range(1, ws.max_column + 1):
        raw = ws.cell(HEADER_ROW, col).value
        if raw is not None:
            name = str(raw).strip()
            sheet_headers.setdefault(name, []).append(col)

    missing_req = [c for c in REQUIRED_COLS if c not in sheet_headers]
    if missing_req:
        return None, None, None, None, None, None, (
            f"Required column(s) not found in row {HEADER_ROW}: "
            + ", ".join(f"'{c}'" for c in missing_req)
        )

    plan_cols,  _  = resolve_plan_columns(sheet_headers)
    act_cols       = resolve_activity_columns(sheet_headers)
    treat_cols     = resolve_treatment_columns(sheet_headers)

    special = {name: sheet_headers[name][0] for name in REQUIRED_COLS}

    col_code = special["Code Trial"]
    col_sect = special["Sector"]
    col_comp = special["Comp"]
    col_year = special["Year"]
    col_reg  = special["Register"]

    needed = (
        set(c for c in plan_cols if c is not None)
        | set(act_cols.values())
        | set(treat_cols.values())
        | set(special.values())
    )

    rows_meta:  list[dict]       = []
    rows_cache: dict[int, dict]  = {}

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        code = ws.cell(row, col_code).value
        if not code:
            continue

        reg    = str(ws.cell(row, col_reg).value  or "").strip()
        year   = ws.cell(row, col_year).value
        sector = str(ws.cell(row, col_sect).value or "").strip()
        comp   = str(ws.cell(row, col_comp).value or "").strip()

        rows_meta.append({
            "row": row, "code": str(code).strip(), "register": reg,
            "year": year, "sector": sector, "comp": comp,
            "filename": f"{str(code).strip()} ({sector}{comp}).xlsx",
        })

        if reg.lower() in NOT_YET:
            rows_cache[row] = {col: ws.cell(row, col).value for col in needed}

    return rows_meta, rows_cache, plan_cols, act_cols, treat_cols, special, None


def filter_rows(rows_meta: list, year_filter: str) -> list:
    out = [r for r in rows_meta if r["register"].lower() in NOT_YET]
    if year_filter != "All":
        out = [r for r in out if str(r["year"]) == year_filter]
    return out


def metric_card(col, icon: str, label: str, value):
    col.markdown(
        f"""<div class="metric-card">
                <div class="icon">{icon}</div>
                <div class="label">{label}</div>
                <div class="value">{value}</div>
            </div>""",
        unsafe_allow_html=True,
    )


def col_chips(cols: list[str]) -> str:
    chips = "".join(f'<span class="col-chip">{c}</span>' for c in cols)
    return f'<div class="col-chips">{chips}</div>'


# ═════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div class="sidebar-brand">🌳 Trial Tools</div>', unsafe_allow_html=True)
    st.caption("Acacia Silviculture R&D")
    st.divider()

    st.markdown("#####  Upload Database")
    uploaded = st.file_uploader(
        "Excel file — sheet **Trial Register**, headers on **row 3**",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    st.divider()
    st.markdown("##### ℹ️ How it works")
    st.caption(
        "Upload your **Trial Register** database. The app reads records "
        "marked **‘Not yet Register’**, lets you filter by year, then "
        "generates **Trial, Schedule & Treatment Plan Automatically** — one Excel per trial, "
        "bundled into a ZIP."
    )
    # st.divider()
    # st.caption("v2.0 · Light Edition")


# ═════════════════════════════════════════════
# HERO HEADER
# ═════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <span class="hero-badge">🌱 ACACIA SILVICULTURE R&D TOOLKIT</span>
    <h1>Trial Tools Automation</h1>
    <p>Transform your Trial Register database into ready-to-use Trial, Schedule &amp; Treatment Plan files — instantly, and beautifully.</p>
</div>
""", unsafe_allow_html=True)


# ═════════════════════════════════════════════
# NO FILE STATE
# ═════════════════════════════════════════════
if not uploaded:
    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        st.markdown("""
        <div class="tool-desc">
            <div class="tool-icon">📋</div>
            <h4>Trial Plan</h4>
            <p>One Excel per trial with all plan fields mapped directly from your database columns.</p>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="tool-desc">
            <div class="tool-icon">📅</div>
            <h4>Schedule Plan</h4>
            <p>One Excel per trial listing every activity that has a scheduled date, neatly structured.</p>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="tool-desc">
            <div class="tool-icon">🧪</div>
            <h4>Treatment Plan</h4>
            <p>One Excel per trial with a parsed row per treatment, extracted from the design notes.</p>
        </div>""", unsafe_allow_html=True)

    st.info("👈  **Upload your database file** in the sidebar to begin.")
    st.markdown('<div class="footer">Trial Tools Automation &nbsp;•&nbsp; <b>Acacia Silviculture R&D 🌳</b></div>', unsafe_allow_html=True)
    st.stop()


file_bytes = uploaded.read()

with st.spinner("Reading database…"):
    rows_meta, rows_cache, plan_cols, act_cols, treat_cols, special, err = \
        load_database(file_bytes)

if err:
    st.error(f"❌  {err}")
    st.stop()

col_estb = special["Established by"]
st.success(f"✅  Database loaded successfully — **{len(rows_meta)}** total trial record(s) detected.")


# ═════════════════════════════════════════════
# FILTERS + METRICS
# ═════════════════════════════════════════════
all_not_yet = filter_rows(rows_meta, "All")
available_years = sorted(
    {str(r["year"]) for r in all_not_yet if r["year"]},
    reverse=True,
)

st.markdown('<div class="section-title"><span class="num">1</span> Configure &amp; Preview</div>', unsafe_allow_html=True)

fcol, _ = st.columns([1, 3])
with fcol:
    year_filter = st.selectbox("📅 Filter by Year", options=["All"] + available_years)

filtered = filter_rows(rows_meta, year_filter)

m1, m2, m3, m4 = st.columns(4, gap="medium")
metric_card(m1, "📊", "Total Records",    len(rows_meta))
metric_card(m2, "⏳", "Not Yet Register", len(all_not_yet))
metric_card(m3, "✅", "Matched (Filter)", len(filtered))
metric_card(m4, "📅", "Active Year",      year_filter)

st.write("")

if not filtered:
    st.warning("No records found with **‘Not yet Register’** status for the selected year.")
    st.markdown('<div class="footer">Trial Tools Generator &nbsp;•&nbsp; <b>Acacia Silviculture R&D</b></div>', unsafe_allow_html=True)
    st.stop()

with st.expander(f"📑  Preview matched trials  ·  {len(filtered)} records", expanded=True):
    st.dataframe(
        pd.DataFrame([{
            "File Name":  r["filename"],
            "Code Trial": r["code"],
            "Sector":     r["sector"],
            "Comp":       r["comp"],
            "Year":       r["year"],
        } for r in filtered]),
        use_container_width=True,
        hide_index=True,
        height=min(440, 60 + 36 * len(filtered)),
    )


# ═════════════════════════════════════════════
# GENERATION TABS
# ═════════════════════════════════════════════
st.markdown('<div class="section-title"><span class="num">2</span> Generate Files</div>', unsafe_allow_html=True)

tab_trial, tab_sched, tab_treat = st.tabs([
    "📋  Trial Plan", "📅  Schedule Plan", "🧪  Treatment Plan",
])

ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
yr = year_filter if year_filter != "All" else "ALL"


# ── TAB 1 : TRIAL PLAN ───────────────────────
with tab_trial:
    st.markdown("""
    <div class="tool-desc">
        <div class="tool-icon">📋</div>
        <h4>Trial Plan</h4>
        <p>Generates one Excel file per trial, with all plan fields mapped automatically from your database.</p>
    </div>""", unsafe_allow_html=True)

    with st.expander("🔍  View column mapping"):
        st.dataframe(
            pd.DataFrame([
                {"Trial Plan Label": label,
                 "Resolved Column": f"Col {col}" if col else "⚠️ Not found"}
                for label, col in zip(PLAN_LABELS, plan_cols)
            ]),
            use_container_width=True, hide_index=True,
        )

    if st.button(f"⚡  Generate {len(filtered)} Trial Plan(s)",
                 type="primary", key="gen_trial", use_container_width=True):
        with st.spinner(f"Building {len(filtered)} Trial Plan files…"):
            zip_bytes = build_trial_plan_zip(filtered, rows_cache, plan_cols)
        st.success(f"✅  {len(filtered)} Trial Plan(s) generated successfully!")
        st.download_button(
            "📥  Download Trial Plans (ZIP)", data=zip_bytes,
            file_name=f"Trial_Plans_{yr}_{ts}.zip", mime="application/zip",
            type="primary", use_container_width=True, key="dl_trial",
        )
        st.caption("📦  ZIP → `Trial Plan/` folder · one Excel per trial.")


# ── TAB 2 : SCHEDULE PLAN ────────────────────
with tab_sched:
    st.markdown("""
    <div class="tool-desc">
        <div class="tool-icon">📅</div>
        <h4>Schedule Plan</h4>
        <p>Generates one Excel per trial listing only activities that have a scheduled date.</p>
    </div>""", unsafe_allow_html=True)

    st.markdown("**Output columns**", unsafe_allow_html=True)
    st.markdown(col_chips([
        "activity_group", "activity_code", "plan_date", "do_date",
        "update_date", "required", "note", "ass_program_by", "ass_staff_by",
    ]), unsafe_allow_html=True)

    with st.expander("🔍  View activity column mapping"):
        st.dataframe(
            pd.DataFrame([
                {"Activity Code": code, "Activity Group": group,
                 "DB Column": f"Col {act_cols[code]}" if code in act_cols else "⚠️ Not found"}
                for code, group in ACTIVITY_DEFS
            ]),
            use_container_width=True, hide_index=True,
        )
        st.caption(
            f"`required` = 1 always · `ass_staff_by` ← **Established by** (Col {col_estb}) · "
            f"Dates use Afrikaans locale `[$-0436]YYYY-MM-DD`."
        )

    if st.button(f"⚡  Generate {len(filtered)} Schedule Plan(s)",
                 type="primary", key="gen_sched", use_container_width=True):
        with st.spinner(f"Building {len(filtered)} Schedule Plan files…"):
            zip_bytes = build_schedule_zip(filtered, rows_cache, act_cols, col_estb)
        st.success(f"✅  {len(filtered)} Schedule Plan(s) generated successfully!")
        st.download_button(
            "📥  Download Schedule Plans (ZIP)", data=zip_bytes,
            file_name=f"Schedule_Plans_{yr}_{ts}.zip", mime="application/zip",
            type="primary", use_container_width=True, key="dl_sched",
        )
        st.caption("📦  ZIP → `Schedule Plan/` folder · one Excel per trial.")


# ── TAB 3 : TREATMENT PLAN ───────────────────
with tab_treat:
    st.markdown("""
    <div class="tool-desc">
        <div class="tool-icon">🧪</div>
        <h4>Treatment Plan</h4>
        <p>Generates one Excel per trial with a row per treatment, parsed from
        <b>Design And Treatment</b> (falls back to <b>Title</b> if empty).</p>
    </div>""", unsafe_allow_html=True)

    st.markdown("**Output columns**", unsafe_allow_html=True)
    st.markdown(col_chips([
        "treat_no", "treat_lo", "treat_name", "treat_type", "germplasmid",
        "deployid", "femaleid", "maleid", "provenance", "species", "bri", "iri",
    ]), unsafe_allow_html=True)

    with st.expander("🔍  View parsing rules"):
        st.markdown("""
| Pattern | Action |
|---|---|
| Lines starting with `T1.` `T2.` … | Extracted as individual treatments (prefix stripped) |
| Lines starting with `-` | Extracted as individual treatments (dash stripped) |
| Neither pattern found | Entire cell text used as one treatment |
| Design And Treatment is empty | Falls back to **Title** column |

**Fixed values:** `treat_type` = `Treatment` · `required` = always filled
**Species mapping:** `AC` → `ACRA` · `AH` → `AHYB`
**Spacing:** `3 x 2.3 m` → `bri = 3`, `iri = 2.3`
        """)

    if st.button(f"⚡  Generate {len(filtered)} Treatment Plan(s)",
                 type="primary", key="gen_treat", use_container_width=True):
        with st.spinner(f"Building {len(filtered)} Treatment Plan files…"):
            zip_bytes = build_treatment_zip(filtered, rows_cache, treat_cols)
        st.success(f"✅  {len(filtered)} Treatment Plan(s) generated successfully!")
        st.download_button(
            "📥  Download Treatment Plans (ZIP)", data=zip_bytes,
            file_name=f"Treatment_Plans_{yr}_{ts}.zip", mime="application/zip",
            type="primary", use_container_width=True, key="dl_treat",
        )
        st.caption("📦  ZIP → `Treatment Plan/` folder · one Excel per trial.")


# ═════════════════════════════════════════════
# FOOTER
# ═════════════════════════════════════════════
st.markdown("""
<div class="footer">
    Trial Tools Automation &nbsp;•&nbsp; <b>Acacia Silviculture R&amp;D 🌳</b> &nbsp;
</div>
""", unsafe_allow_html=True)