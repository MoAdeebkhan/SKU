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
# ---------------------------------------------------------------------------

st.markdown("""
<style>

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
        label_visibility="collapsed",
    )
    mode_is_ai = mode_choice == "AI AI"

    st.markdown("---")

    if mode_is_ai:
        st.markdown('<span class="sb-label">AI Settings</span>', unsafe_allow_html=True)

        reachable, live_models, conn_err = check_connection()

        if reachable:
            st.markdown(
                '<div class="status-row status-online">'
                '<div class="sdot sdot-green"></div>Connected</div>',
                unsafe_allow_html=True,
            )
            available_models = live_models if live_models else KNOWN_MODELS
        else:
            st.markdown(
                '<div class="status-row status-offline">'
                '<div class="sdot sdot-red"></div>Offline</div>',
                unsafe_allow_html=True,
            )
            st.caption(conn_err)
            available_models = KNOWN_MODELS

        selected_model = st.selectbox(
            "Model",
            options=available_models,
            label_visibility="visible",
        )
        st.caption(AI_HOST)

    else:
        reachable      = False
        selected_model = None
        st.markdown('<span class="sb-label">Mode Info</span>', unsafe_allow_html=True)
        st.caption(
            "Analysis and summary are generated directly from "
            "the data.   model or internet connection required."
        )

    st.markdown("---")
    st.markdown('<span class="sb-label">Options</span>', unsafe_allow_html=True)
    dry_run = st.toggle("Dry Run (preview only, no file saved)", value=False)

    st.markdown("---")
    st.markdown('<span class="sb-label">Expected Column Names</span>', unsafe_allow_html=True)
    st.caption("Material Number")
    st.caption("SKU Name")
    st.caption("Receiving Location")
    st.caption("Number of Stocks")


# ---------------------------------------------------------------------------
# Main header
# ---------------------------------------------------------------------------

st.markdown("# Stock Balancing Agent")
st.markdown(
    '<p style="color:#6b7280; margin-top:-8px; margin-bottom:16px; font-size:0.9rem;">'
    "Upload an Excel file, review stock imbalances, run the agent, download the updated file."
    "</p>",
    unsafe_allow_html=True,
)

if mode_is_ai:
    st.markdown(
        '<div class="mode-banner mode-ai">'
        '<b>Mode: AI AI</b> — Analysis and summary will be generated by the language model.'
        '</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="mode-banner mode-local">'
        '<b>Mode: ETL</b> — Analysis and summary are generated from the data.   model needed.'
        '</div>',
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------------------------
# File upload
# ---------------------------------------------------------------------------

uploaded = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xlsm"],
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown("""
<div class="card" style="text-align:center; padding:36px 20px; margin-top:8px;">
  <div style="font-size:0.95rem; color:#6b7280; margin-bottom:4px;">
    Drop your <b style="color:#111827">.xlsx</b> or <b style="color:#111827">.xlsm</b> file above to get started
  </div>
  <div style="font-size:0.82rem; color:#9ca3af; margin-top:4px;">
    Detects SKUs at multiple locations and redistributes stock evenly
  </div>
</div>
""", unsafe_allow_html=True)
    st.stop()


# ---------------------------------------------------------------------------
# Load file
# ---------------------------------------------------------------------------

file_bytes = uploaded.read()

try:
    df, wb = load_excel_bytes(file_bytes)
except ExcelLoadError as e:
    st.error(f"Could not load file: {e}")
    st.stop()
except Exception as e:
    st.error(f"Unexpected error loading file: {e}")
    st.stop()


# ---------------------------------------------------------------------------
# Overview metrics
# ---------------------------------------------------------------------------

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(f"""
<div class="card card-blue">
  <div class="metric-label">Total Rows</div>
  <div class="metric-value">{len(df):,}</div>
</div>""", unsafe_allow_html=True)
with c2:
    st.markdown(f"""
<div class="card card-blue">
  <div class="metric-label">Unique SKUs</div>
  <div class="metric-value">{df[COL_SKU].nunique():,}</div>
</div>""", unsafe_allow_html=True)
with c3:
    st.markdown(f"""
<div class="card card-blue">
  <div class="metric-label">Total Stock Units</div>
  <div class="metric-value">{int(df[COL_STOCK].sum()):,}</div>
</div>""", unsafe_allow_html=True)

with st.expander("Raw Data Preview", expanded=False):
    st.dataframe(df, use_container_width=True, height=280)

st.markdown("---")

# ---------------------------------------------------------------------------
# Duplicate detection
# ---------------------------------------------------------------------------

duplicates = find_duplicate_skus(df)
sku_groups = build_sku_groups(duplicates)
result     = build_result(df, sku_groups)

if not sku_groups:
    st.markdown("""
<div class="card card-green" style="padding:20px 22px;">
  <b style="color:#166534;">All SKUs are already balanced</b>
  <p style="color:#6b7280; margin:6px 0 0; font-size:0.85rem;">
    No SKU appears in more than one location. Nothing to do.
  </p>
</div>""", unsafe_allow_html=True)
    st.stop()

st.markdown(f"### Found {len(sku_groups)} SKU(s) Across Multiple Locations")

for g in sku_groups:
    already = g.is_already_balanced()
    badge   = (
        '<span class="badge badge-ok">Already Balanced</span>'
        if already else
        '<span class="badge badge-warn">Unbalanced</span>'
    )

    rows_html = ""
    for loc, old, new in zip(g.locations, g.old_stocks, g.new_stocks):
        if old != new:
            dot      = '<span class="dot-changed"></span>'
            old_cell = f'<span class="val-old">{old:,}</span>'
            new_cell = f'<span class="val-new">{new:,}</span>'
        else:
            dot      = '<span class="dot-same"></span>'
            old_cell = f'<span class="val-same">{old:,}</span>'
            new_cell = f'<span class="val-same">{new:,}</span>'

        rows_html += f"""
<tr>
  <td style="width:22px;">{dot}</td>
  <td>{loc}</td>
  <td>{old_cell}</td>
  <td style="width:30px; text-align:center;"><span class="arrow">&#8594;</span></td>
  <td>{new_cell}</td>
</tr>"""

    st.markdown(f"""
<div class="card card-orange" style="margin-bottom:14px;">
  <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
    <span style="font-weight:700; color:#111827; font-size:0.97rem;">{g.sku_name}</span>
    {badge}
  </div>
  <div class="pill-row">
    <span class="pill">Locations: <b>{len(g.locations)}</b></span>
    <span class="pill">Total Stock: <b>{g.total:,}</b></span>
    <span class="pill">Variance: <b>{g.variance():,}</b></span>
    <span class="pill">Approx. per location: <b>{g.total // len(g.locations):,}</b></span>
  </div>
  <table class="diff-table">
    <thead>
      <tr>
        <th></th><th>Location</th><th>Current Stock</th>
        <th></th><th>New Stock</th>
      </tr>
    </thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>""", unsafe_allow_html=True)

st.markdown("---")

# ---------------------------------------------------------------------------
# Run section
# ---------------------------------------------------------------------------

st.markdown("### Run Balancing Agent")

if mode_is_ai and not reachable:
    st.warning(
        f"AI is not reachable at {AI_HOST}. "
        "Switch to ETL mode in the sidebar, or start the AI server."
    )

btn_col, _ = st.columns([1, 2])
with btn_col:
    run_clicked = st.button("Balance Stocks Now", use_container_width=True)

if not run_clicked:
    st.stop()

# ---------------------------------------------------------------------------
# Execution pipeline
# ---------------------------------------------------------------------------

progress = st.progress(0)
status   = st.empty()

status.markdown(
    '<p style="color:#2563eb; font-size:0.86rem; margin:4px 0;">Generating analysis...</p>',
    unsafe_allow_html=True,
)
progress.progress(20)

if mode_is_ai and reachable:
    analysis_text  = ai_analyze(selected_model, sku_groups)
    analysis_label = f"AI Analysis  |  {selected_model}"
else:
    analysis_text  = generate_analysis_text(sku_groups)
    analysis_label = "Analysis  |  ETL"

progress.progress(50)

status.markdown(
    '<p style="color:#2563eb; font-size:0.86rem; margin:4px 0;">Generating summary...</p>',
    unsafe_allow_html=True,
)

if mode_is_ai and reachable:
    summary_text  = ai_summary(selected_model, sku_groups, dry_run)
    summary_label = f"Operation Summary  |  {selected_model}"
else:
    summary_text  = generate_summary_text(sku_groups, dry_run)
    summary_label = "Operation Summary  |  ETL"

progress.progress(75)

status.markdown(
    '<p style="color:#2563eb; font-size:0.86rem; margin:4px 0;">Applying changes to file...</p>',
    unsafe_allow_html=True,
)

if dry_run:
    output_bytes = file_bytes
else:
    try:
        output_bytes = apply_balancing_to_workbook(wb, sku_groups)
    except Exception as e:
        progress.empty()
        status.empty()
        st.error(f"Failed to write updated file: {e}")
        st.stop()

progress.progress(100)
time.sleep(0.2)
status.empty()
progress.empty()

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

st.markdown("---")
st.markdown("### Balancing Complete")

st.markdown(f"""
<div class="analysis-box">
  <div class="analysis-label">{analysis_label}</div>
  {analysis_text}
</div>""", unsafe_allow_html=True)

st.markdown(f"""
<div class="analysis-box">
  <div class="analysis-label">{summary_label}</div>
  {summary_text}
</div>""", unsafe_allow_html=True)

st.markdown("")

r1, r2, r3, r4 = st.columns(4)
mode_color = "#f97316" if dry_run else "#16a34a"
mode_str   = "DRY RUN"  if dry_run else "SAVED"

with r1:
    st.markdown(f"""
<div class="card card-green">
  <div class="metric-label">SKUs Rebalanced</div>
  <div class="metric-value" style="font-size:1.7rem; color:#16a34a;">{result.skus_balanced}</div>
</div>""", unsafe_allow_html=True)
with r2:
    st.markdown(f"""
<div class="card card-green">
  <div class="metric-label">Rows Updated</div>
  <div class="metric-value" style="font-size:1.7rem; color:#16a34a;">{result.rows_updated}</div>
</div>""", unsafe_allow_html=True)
with r3:
    st.markdown(f"""
<div class="card card-green">
  <div class="metric-label">Total Units</div>
  <div class="metric-value" style="font-size:1.7rem; color:#16a34a;">{result.total_stock:,}</div>
</div>""", unsafe_allow_html=True)
with r4:
    st.markdown(f"""
<div class="card card-green">
  <div class="metric-label">Status</div>
  <div class="metric-value" style="font-size:1.7rem; color:{mode_color};">{mode_str}</div>
</div>""", unsafe_allow_html=True)

st.markdown("")

if dry_run:
    st.info(
        "Dry Run is on. Turn it off in the sidebar and click "
        "Balance Stocks Now again to save the file."
    )
else:
    ext   = ".xlsm" if uploaded.name.endswith(".xlsm") else ".xlsx"
    fname = uploaded.name.replace(ext, f"_balanced{ext}")
    st.download_button(
        label     = "Download Balanced Excel File",
        data      = output_bytes,
        file_name = fname,
        mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("")

with st.expander("Detailed Change Log", expanded=True):
    for g in sku_groups:
        rows_html = ""
        for loc, old, new in zip(g.locations, g.old_stocks, g.new_stocks):
            if old != new:
                dot      = '<span class="dot-changed"></span>'
                old_cell = f'<span class="val-old">{old:,}</span>'
                new_cell = f'<span class="val-new">{new:,}</span>'
            else:
                dot      = '<span class="dot-same"></span>'
                old_cell = f'<span class="val-same">{old:,}</span>'
                new_cell = f'<span class="val-same">{new:,}</span>'

            rows_html += f"""
<tr>
  <td style="width:22px;">{dot}</td>
  <td>{loc}</td>
  <td>{old_cell}</td>
  <td style="width:30px; text-align:center;"><span class="arrow">&#8594;</span></td>
  <td>{new_cell}</td>
</tr>"""

        st.markdown(f"""
<div class="card" style="margin-bottom:10px;">
  <span style="font-weight:700; color:#111827;">{g.sku_name}</span>
  <span style="margin-left:10px; color:#9ca3af; font-size:0.8rem;">
    {g.total:,} units across {len(g.locations)} locations
  </span>
  <table class="diff-table" style="margin-top:10px;">
    <thead>
      <tr><th></th><th>Location</th><th>Old Stock</th><th></th><th>New Stock</th></tr>
    </thead>
    <tbody>{rows_html}</tbody>
  </table>
</div>""", unsafe_allow_html=True)