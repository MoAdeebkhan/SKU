"""
app.py
------
Stock Balancing Agent - Streamlit Frontend
Two modes: AI  |  ETL

Run with:
    streamlit run app.py
"""

import time
import streamlit as st

from core import (
    ExcelLoadError,
    load_excel_bytes,
    find_duplicate_skus,
    build_sku_groups,
    apply_balancing_to_workbook,
    generate_analysis_text,
    generate_summary_text,
    build_result,
    COL_SKU,
    COL_STOCK,
    COL_LOCATION,
)
from ai_engine import (
    check_connection,
    ai_analyze,
    ai_summary,
    AI_HOST,
    KNOWN_MODELS,
)

# ---------------------------------------------------------------------------
# Page config — must be first Streamlit call
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Stock Balancer",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Force light theme via inline style injection
# Targets both the app container and sidebar regardless of system theme
# Includes styling to completely hide the top menu, toolbar, and deploy footer
# ---------------------------------------------------------------------------

st.markdown("""
<style>

/* ── Hide Streamlit Elements (Toolbar, Footer, Deploy Button) ── */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="stToolbar"] {visibility: hidden;}
[data-testid="stDecoration"] {visibility: hidden;}
[data-testid="stStatusWidget"] {visibility: hidden;}
.stAppDeployButton {display: none;}

/* ── Force light background everywhere ── */
html, body, [data-testid="stApp"],
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
.main, .block-container {
    background-color: #f4f6fb !important;
    color: #111827 !important;
}

[data-testid="stSidebar"],
[data-testid="stSidebar"] > div,
[data-testid="stSidebar"] section {
    background-color: #ffffff !important;
    border-right: 1px solid #e2e8f0 !important;
}

/* ── Force all text to dark ── */
body, p, span, label, div,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div,
.stMarkdown, .stMarkdown p {
    color: #111827 !important;
}

/* ── Headings ── */
h1 { font-size: 1.7rem !important; font-weight: 700 !important;
     letter-spacing: -0.3px !important; color: #111827 !important; }
h2 { font-size: 1.2rem !important; font-weight: 600 !important;
     color: #1f2937 !important; }
h3 { font-size: 1rem !important;  font-weight: 600 !important;
     color: #1f2937 !important; }

/* ── Streamlit selectbox, toggle, expander ── */
[data-testid="stSelectbox"] > div > div {
    background: #ffffff !important;
    border: 1px solid #d1d5db !important;
    color: #111827 !important;
    border-radius: 7px !important;
}
[data-testid="stSelectbox"] label { color: #374151 !important; font-size: 0.85rem !important; }

.stToggle label, [data-testid="stToggle"] label {
    color: #374151 !important;
}

details summary { color: #374151 !important; }

/* ── Main button ── */
div.stButton > button {
    background: #2563eb !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 10px 22px !important;
    font-size: 0.9rem !important;
    width: 100% !important;
}
div.stButton > button:hover { background: #1d4ed8 !important; }

/* ── Download button ── */
div[data-testid="stDownloadButton"] > button {
    background: #f0fdf4 !important;
    color: #166534 !important;
    border: 1.5px solid #86efac !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    width: 100% !important;
    font-size: 0.9rem !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: #dcfce7 !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border: 2px dashed #cbd5e1 !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: #ffffff !important;
}

/* ── Progress bar ── */
.stProgress > div > div > div { background: #2563eb !important; }

/* ── Alerts ── */
.stAlert { border-radius: 8px !important; }

/* ── Divider ── */
hr { border-color: #e2e8f0 !important; margin: 16px 0 !important; }

/* ── Cards (custom HTML) ── */
.card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 18px 22px;
    margin-bottom: 12px;
}
.card-blue   { border-top: 3px solid #3b82f6; }
.card-orange { border-top: 3px solid #f97316; }
.card-green  { border-top: 3px solid #22c55e; }

.metric-label {
    font-size: 0.72rem;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    font-weight: 600;
}
.metric-value {
    font-size: 2rem;
    font-weight: 700;
    color: #111827;
    margin-top: 4px;
    line-height: 1.1;
}

/* ── Mode indicator ── */
.mode-banner {
    border-radius: 8px;
    padding: 10px 16px;
    font-size: 0.86rem;
    font-weight: 500;
    margin-bottom: 16px;
}
.mode-ai    { background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af; }
.mode-local { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }

/* ── Pills ── */
.pill-row { display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0 4px; }
.pill {
    background: #f1f5f9;
    border: 1px solid #e2e8f0;
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.79rem;
    color: #475569;
}
.pill b { color: #1e293b; }

/* ── Diff table ── */
.diff-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.84rem;
    margin-top: 12px;
}
.diff-table th {
    background: #f8fafc;
    color: #94a3b8;
    font-weight: 600;
    text-align: left;
    padding: 8px 14px;
    border-bottom: 1px solid #e2e8f0;
    font-size: 0.73rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.diff-table td {
    padding: 9px 14px;
    border-bottom: 1px solid #f1f5f9;
    color: #374151;
    vertical-align: middle;
}
.diff-table tr:hover td { background: #f8fafc; }
.val-old  { color: #dc2626; font-weight: 600; }
.val-new  { color: #16a34a; font-weight: 600; }
.val-same { color: #94a3b8; }
.arrow    { color: #d1d5db; margin: 0 6px; }
.dot-changed {
    display: inline-block; width: 8px; height: 8px;
    border-radius: 50%; background: #f97316;
    margin-right: 5px; vertical-align: middle;
}
.dot-same {
    display: inline-block; width: 8px; height: 8px;
    border-radius: 50%; background: #22c55e;
    margin-right: 5px; vertical-align: middle;
}

/* ── Badges ── */
.badge {
    display: inline-block; border-radius: 5px;
    padding: 2px 9px; font-size: 0.72rem; font-weight: 700;
}
.badge-warn { background: #fff7ed; color: #c2410c; border: 1px solid #fed7aa; }
.badge-ok   { background: #f0fdf4; color: #166534; border: 1px solid #bbf7d0; }

/* ── Analysis box ── */
.analysis-box {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-left: 4px solid #2563eb;
    border-radius: 8px;
    padding: 16px 20px;
    font-size: 0.87rem;
    color: #374151;
    line-height: 1.75;
    white-space: pre-wrap;
    margin-bottom: 14px;
}
.analysis-label {
    font-size: 0.69rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    color: #2563eb;
    text-transform: uppercase;
    margin-bottom: 10px;
}

/* ── Status indicator ── */
.status-row {
    display: flex; align-items: center; gap: 8px;
    padding: 8px 12px; border-radius: 7px;
    font-size: 0.83rem; font-weight: 500; margin-bottom: 10px;
}
.status-online  { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
.status-offline { background: #fef2f2; border: 1px solid #fecaca; color: #b91c1c; }
.sdot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.sdot-green { background: #22c55e; }
.sdot-red   { background: #ef4444; }

/* ── Sidebar labels ── */
.sb-label {
    font-size: 0.69rem; font-weight: 700;
    letter-spacing: 0.1em; text-transform: uppercase;
    color: #94a3b8 !important; margin: 14px 0 6px 0;
    display: block;
}
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

with st.sidebar:
    st.markdown("## Stock Balancer")
    st.markdown("---")

    st.markdown('<span class="sb-label">Processing Mode</span>', unsafe_allow_html=True)

    mode_choice = st.selectbox(
        "Select Mode",
        options=["AI", "ETL ( )"],
        index=0,
        label_visibility="collapsed"
    )
