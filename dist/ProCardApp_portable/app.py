import io
import os
import re
import tempfile
import textwrap
from collections import Counter
from contextlib import redirect_stdout
from datetime import datetime
from calendar import monthrange

import pandas as pd
import streamlit as st

from processor import build_file_feed_from_teams_workflow, run_procard_pipeline


st.set_page_config(page_title="ProCard Reconciliation App", layout="wide")

# ── Brand Primary  ────────────────────────────────────
BRAND_PRIME = "#5D1725"
BRAND_GREY = "#777777"
BRAND_LIGHT = "#C1C6C8"

# ── Brand Accent  ─────────────────────────────────────
BRAND_ACCENT = "#A69F88"
BRAND_GREEN = "#8F993E"
BRAND_ORANGE = "#A9431E"
BRAND_GOLD = "#C99700"
BRAND_CREAM = "#DAC79D"

WHITE = "#FFFFFF"
BLACK = "#000000"

STEP_TITLES = [
    "Welcome",
    "Upload Files",
    "Parse & Match",
    "FOAPAL Review",
    "Download Output",
]

FILE_FEED_COLUMN_WIDTHS = [31.14, 1.57, 12.43, 1.29, 28.14, 6.29, 5.29]


def hex_to_rgba(hex_color, alpha):
    h = hex_color.lstrip("#")
    if len(h) != 6:
        return f"rgba(0,0,0,{alpha})"
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def render_helper_box(message, tone="info", target=None, compact=False):
    color_map = {
        "info": BRAND_GOLD,
        "warning": BRAND_GOLD,
        "success": BRAND_GREEN,
        "error": BRAND_ORANGE,
        "prime": BRAND_PRIME,
    }
    base = color_map.get(tone, BRAND_GOLD)
    bg = hex_to_rgba(base, 0.14)
    border = hex_to_rgba(base, 0.45)
    host = target if target is not None else st
    if compact:
        html = (
            f"<div style='text-align:center;'>"
            f"<div style='display:inline-block; border-radius:8px; border:1px solid {border}; "
            f"background:{bg}; padding:0.65rem 0.85rem; color:{BLACK};'>{message}</div>"
            f"</div>"
        )
    else:
        html = (
            f"<div style='border-radius:8px; border:1px solid {border}; "
            f"background:{bg}; padding:0.65rem 0.85rem; color:{BLACK};'>{message}</div>"
        )
    host.markdown(
        html,
        unsafe_allow_html=True,
    )


def apply_brand_theme():
    st.markdown(
        f"""
        <style>
        :root {{
            --brand-prime: {BRAND_PRIME};
            --brand-grey: {BRAND_GREY};
            --brand-light: {BRAND_LIGHT};
            --brand-accent: {BRAND_ACCENT};
            --brand-green: {BRAND_GREEN};
            --brand-orange: {BRAND_ORANGE};
            --brand-gold: {BRAND_GOLD};
            --brand-cream: {BRAND_CREAM};
            --white: {WHITE};
            --black: {BLACK};
        }}

        .stApp {{
            background: var(--white);
            color: var(--black);
        }}

        h1, h2, h3 {{
            color: var(--brand-prime);
        }}

        [data-testid="stSidebar"] {{
            background: var(--brand-light);
        }}

        [data-testid="stMetric"] {{
            background: var(--white);
            border: 1px solid var(--brand-light);
            border-radius: 10px;
            padding: 8px;
        }}

        [data-testid="stAlert"] {{
            border-radius: 8px;
            border: 1px solid var(--brand-accent);
        }}

        div.stButton > button {{
            background: var(--brand-prime);
            color: var(--white);
            border: 1px solid var(--brand-prime);
            border-radius: 8px;
        }}

        div.stButton > button:hover {{
            background: var(--brand-orange);
            border-color: var(--brand-orange);
            color: var(--white);
        }}

        div.stDownloadButton > button {{
            background: var(--brand-prime);
            color: var(--white);
            border: 1px solid var(--brand-prime);
            border-radius: 8px;
        }}

        [data-testid="stProgressBar"] div[role="progressbar"] {{
            background-color: var(--brand-prime) !important;
        }}

        [data-testid="stProgressBar"] > div > div > div {{
            background-color: var(--brand-prime) !important;
        }}

        [data-testid="stProgressBar"] div[role="progressbar"] > div {{
            background-color: var(--brand-prime) !important;
        }}

        .stTextInput input,
        .stSelectbox div[data-baseweb="select"] > div,
        .stFileUploader,
        .stDataFrame {{
            border-color: var(--brand-light) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_top_logo():
    logo_png_path = os.path.join(os.path.dirname(__file__), "logo.png")
    logo_pdf_path = os.path.join(os.path.dirname(__file__), "ScriptState_Maroon(RGB).pdf")
    if os.path.exists(logo_png_path):
        st.image(logo_png_path, width=230)
        return

    if not os.path.exists(logo_pdf_path):
        return

    # Preferred: render first page of PDF to PNG if PyMuPDF is available.
    try:
        import fitz  # type: ignore

        with fitz.open(logo_pdf_path) as doc:
            if doc.page_count > 0:
                page = doc.load_page(0)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                st.image(pix.tobytes("png"), width=230)
                return
    except Exception:
        pass


def init_state():
    defaults = {
        "step": 0,
        "workflow_file": None,
        "zip_file": None,
        "batch_date": default_batch_date(),
        "processing_done": False,
        "processing_log": "",
        "summary": {},
        "teams_workflow": None,
        "bank_review": None,
        "workflow_review": None,
        "workflow_only_rows": None,
        "source_output_file": "",
        "foapal_review_df": None,
        "foapal_current_pos": 0,
        "foapal_review_skipped": False,
        "teams_workflow_original": None,
        "final_output_bytes": None,
        "final_output_name": "",
        "final_output_signature": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def go_next():
    st.session_state.step = min(st.session_state.step + 1, len(STEP_TITLES) - 1)


def default_batch_date():
    today = datetime.now()
    last_day = monthrange(today.year, today.month)[1]
    return today.replace(day=last_day).strftime("%y%m%d")


def is_last_day_batch_date_YYMMDD(value):
    try:
        parsed = datetime.strptime(str(value).strip(), "%y%m%d")
    except Exception:
        return False
    return parsed.day == monthrange(parsed.year, parsed.month)[1]


def reset_upload_step_state():
    st.session_state.workflow_file = None
    st.session_state.zip_file = None
    st.session_state.batch_date = default_batch_date()
    if "workflow_upload" in st.session_state:
        del st.session_state["workflow_upload"]
    if "zip_upload" in st.session_state:
        del st.session_state["zip_upload"]


def reset_parse_step_state():
    st.session_state.processing_done = False
    st.session_state.processing_log = ""
    st.session_state.summary = {}
    st.session_state.teams_workflow = None
    st.session_state.teams_workflow_original = None
    st.session_state.bank_review = None
    st.session_state.workflow_review = None
    st.session_state.workflow_only_rows = None
    st.session_state.source_output_file = ""
    st.session_state.foapal_review_df = None
    st.session_state.foapal_current_pos = 0
    st.session_state.foapal_review_skipped = False
    st.session_state.final_output_bytes = None
    st.session_state.final_output_name = ""
    st.session_state.final_output_signature = None


def reset_foapal_step_state():
    if st.session_state.teams_workflow_original is not None:
        st.session_state.teams_workflow = st.session_state.teams_workflow_original.copy(deep=True)
    st.session_state.foapal_current_pos = 0
    st.session_state.foapal_review_skipped = False
    st.session_state.final_output_bytes = None
    st.session_state.final_output_name = ""
    st.session_state.final_output_signature = None


def go_back():
    current_step = st.session_state.step
    if current_step == 1:
        reset_upload_step_state()
    elif current_step == 2:
        reset_parse_step_state()
    elif current_step == 3:
        reset_foapal_step_state()
    st.session_state.step = max(st.session_state.step - 1, 0)


def is_blank(val):
    return pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan"


def parse_amount_ui(val):
    if is_blank(val):
        return 0.0
    s = str(val).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def normalize_date_mmdd_ui(val):
    s = "" if is_blank(val) else str(val).strip()
    if not s:
        return ""
    m = re.search(r'(\d{1,2})/(\d{1,2})(?:/\d{2,4})?', s)
    if not m:
        return ""
    return f"{int(m.group(1)):02d}/{int(m.group(2)):02d}"


def normalize_merchant_ui(val):
    s = "" if is_blank(val) else str(val).upper()
    return re.sub(r'[^A-Z0-9]+', ' ', s).strip()


def build_row_key(date_val, amount_val, merchant_val):
    return (
        normalize_date_mmdd_ui(date_val),
        round(parse_amount_ui(amount_val), 2),
        normalize_merchant_ui(merchant_val),
    )


def select_rows_by_key_counts(df, key_col, counter_map):
    if df.empty or not counter_map:
        return pd.DataFrame(columns=df.columns)
    chunks = []
    for key, count in counter_map.items():
        if count <= 0:
            continue
        chunks.append(df.loc[df[key_col] == key].head(int(count)))
    return pd.concat(chunks, ignore_index=True) if chunks else pd.DataFrame(columns=df.columns)


def missing_required_mask(df, required_cols):
    return df[required_cols].apply(lambda col: col.map(is_blank)).any(axis=1)


def sanitize_for_output(df):
    """Remove nulls from output sheets to prevent blank/null export cells."""
    cleaned = df.copy()
    return cleaned.fillna("")


def normalize_foapal_code(val):
    s = "" if is_blank(val) else str(val).strip()
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def normalize_editable_foapal(df, required_cols):
    """Keep FOAPAL editor fields as text so Streamlit doesn't coerce them."""
    normalized = df.copy()
    for col in required_cols:
        normalized[col] = normalized[col].map(normalize_foapal_code)
    return normalized


def foapal_sort_key(value):
    return (0, int(value)) if value.isdigit() else (1, value.lower())


def build_foapal_dropdown_options(df, required_cols):
    options_by_col = {}
    value_counts_by_col = {}
    for col in required_cols:
        values = [normalize_foapal_code(val) for val in df[col] if not is_blank(val)]
        counts = pd.Series(values).value_counts() if values else pd.Series(dtype="int64")
        value_counts_by_col[col] = counts
        options_by_col[col] = sorted(set(values), key=foapal_sort_key)

    # Resolve Account/Program overlaps by keeping each code in the column where it appears more often.
    if 'Account:' in options_by_col and 'Program:' in options_by_col:
        account_set = set(options_by_col['Account:'])
        program_set = set(options_by_col['Program:'])
        overlap = account_set & program_set
        for code in overlap:
            account_count = int(value_counts_by_col['Account:'].get(code, 0))
            program_count = int(value_counts_by_col['Program:'].get(code, 0))
            if account_count >= program_count:
                if code in options_by_col['Program:']:
                    options_by_col['Program:'].remove(code)
            else:
                if code in options_by_col['Account:']:
                    options_by_col['Account:'].remove(code)

    for col in required_cols:
        options_by_col[col] = [""] + options_by_col[col]
    return options_by_col


def selectbox_value(label, current_value, options, allow_outside_current=True):
    current = normalize_foapal_code(current_value)
    choices = list(options)
    if allow_outside_current and current and current not in choices:
        choices.append(current)
    selected_index = choices.index(current) if current in choices else 0
    return st.selectbox(label, options=choices, index=selected_index)


def clamp_review_position(missing_ids):
    if not missing_ids:
        st.session_state.foapal_current_pos = 0
        return
    st.session_state.foapal_current_pos = min(
        st.session_state.foapal_current_pos,
        len(missing_ids) - 1,
    )


def render_step_nav():
    completed_steps = st.session_state.step
    last_idx = max(1, len(STEP_TITLES) - 1)
    progress = completed_steps / last_idx
    progress_pct = int(round(progress * 100))
    fill_width_css = "0%" if progress_pct == 0 else f"min(100%, calc({progress_pct}% + 2px))"

    st.subheader("Progress")
    badge_html = []
    for i, title in enumerate(STEP_TITLES):
        if st.session_state.step > i:
            bg = BRAND_PRIME
            fg = WHITE
            border = BRAND_PRIME
        elif st.session_state.step == i:
            bg = WHITE
            fg = BLACK
            border = BRAND_PRIME
        else:
            bg = BRAND_GREY
            fg = WHITE
            border = BRAND_GREY

        left_pct = int(round((i / last_idx) * 100))
        if i == 0:
            transform = "translate(0, -50%)"
        elif i == len(STEP_TITLES) - 1:
            transform = "translate(-100%, -50%)"
        else:
            transform = "translate(-50%, -50%)"

        badge_html.append(
            textwrap.dedent(
                f"""
                <div style="
                    position:absolute;
                    top:50%;
                    left:{left_pct}%;
                    transform:{transform};
                    background:{bg};
                    color:{fg};
                    border:1px solid {border};
                    border-radius:999px;
                    padding:4px 10px;
                    font-size:0.82rem;
                    font-weight:600;
                    white-space:nowrap;
                    z-index:2;
                ">{i + 1}. {title}</div>
                """
            ).strip()
        )

    progress_html = textwrap.dedent(
        f"""
        <div style="position:relative; width:100%; height:52px; margin-top:4px;">
            <div style="
                position:absolute;
                left:0;
                right:0;
                top:50%;
                transform:translateY(-50%);
                height:14px;
                background:{BRAND_LIGHT};
                border-radius:999px;
                overflow:hidden;
                z-index:1;
            ">
                <div style="width:{fill_width_css}; height:100%; background:{BRAND_PRIME}; border-radius:999px;"></div>
            </div>
            {''.join(badge_html)}
        </div>
        """
    ).strip()

    st.markdown(
        progress_html,
        unsafe_allow_html=True,
    )


def render_back_next(can_next=True):
    c1, c_spacer, c2 = st.columns([1, 10, 1])
    with c1:
        st.button("Back", on_click=go_back, disabled=st.session_state.step == 0)
    with c2:
        st.button("Next", on_click=go_next, disabled=not can_next)


def run_processing_pipeline():
    tmp_dir = tempfile.mkdtemp(prefix="procard_app_")
    workflow_bytes = io.BytesIO(st.session_state.workflow_file.getvalue())
    zip_bytes = io.BytesIO(st.session_state.zip_file.getvalue())
    output_file = os.path.join(tmp_dir, "Final_ProCard_Upload.xlsx")

    log_buf = io.StringIO()
    with redirect_stdout(log_buf):
        summary = run_procard_pipeline(
            workflow_csv=workflow_bytes,
            bank_zip=zip_bytes,
            batch_date=st.session_state.batch_date,
            manual_review_duplicates=False,
            prompt_for_missing_foapal=False,
            output_file=output_file,
        )

    output_file = summary.get("output_file", "Final_ProCard_Upload.xlsx")
    if not os.path.isabs(output_file):
        output_file = os.path.join(os.getcwd(), output_file)

    st.session_state.processing_log = log_buf.getvalue()
    st.session_state.summary = summary
    st.session_state.source_output_file = output_file
    st.session_state.teams_workflow = pd.read_excel(output_file, sheet_name="TEAMS WORKFLOW")
    st.session_state.teams_workflow_original = st.session_state.teams_workflow.copy(deep=True)
    st.session_state.bank_review = pd.read_excel(output_file, sheet_name="PDF Parsed Transactions")
    st.session_state.workflow_review = pd.read_excel(output_file, sheet_name="Original Workflow")
    st.session_state.workflow_only_rows = pd.read_excel(output_file, sheet_name="Workflow Not In PDF")
    st.session_state.foapal_review_df = None
    st.session_state.foapal_current_pos = 0
    st.session_state.foapal_review_skipped = False
    st.session_state.final_output_bytes = None
    st.session_state.final_output_name = ""
    st.session_state.final_output_signature = None
    st.session_state.processing_done = True


def build_final_workbook_bytes():
    teams_workflow = sanitize_for_output(st.session_state.teams_workflow)
    file_feed = build_file_feed_from_teams_workflow(teams_workflow, st.session_state.batch_date)
    file_feed = sanitize_for_output(file_feed)
    bank_review = sanitize_for_output(st.session_state.bank_review)
    workflow_review = sanitize_for_output(st.session_state.workflow_review)
    workflow_only_rows = sanitize_for_output(st.session_state.workflow_only_rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        file_feed.to_excel(writer, sheet_name="FILE FEED", index=False, header=False)
        teams_workflow.to_excel(writer, sheet_name="TEAMS WORKFLOW", index=False)
        bank_review.to_excel(writer, sheet_name="PDF Parsed Transactions", index=False)
        workflow_review.to_excel(writer, sheet_name="Original Workflow", index=False)
        workflow_only_rows.to_excel(writer, sheet_name="Workflow Not In PDF", index=False)

        file_feed_ws = writer.sheets.get("FILE FEED")
        if file_feed_ws is not None:
            for col_letter, width in zip(["A", "B", "C", "D", "E", "F", "G"], FILE_FEED_COLUMN_WIDTHS):
                file_feed_ws.column_dimensions[col_letter].width = width
    output.seek(0)
    return output.getvalue()


def build_final_output_signature():
    if st.session_state.teams_workflow is None:
        return None
    teams_for_sig = sanitize_for_output(st.session_state.teams_workflow)
    data_hash = int(pd.util.hash_pandas_object(teams_for_sig, index=True).sum())
    return (
        tuple(FILE_FEED_COLUMN_WIDTHS),
        st.session_state.batch_date,
        data_hash,
        bool(st.session_state.foapal_review_skipped),
    )


init_state()
apply_brand_theme()

header_logo_col, header_title_col = st.columns([1, 7])
with header_logo_col:
    render_top_logo()
with header_title_col:
    st.title("ProCard Reconciliation Web App")
render_step_nav()

step = st.session_state.step

if step == 0:
    st.subheader("Welcome")
    st.write(
        "Welcome to the ProCard reconciliation app. Upload a workflow CSV and a ZIP of statement PDFs, "
        "then walk through parsing, matching, FOAPAL review, and final download."
    )
    render_back_next(can_next=True)

elif step == 1:
    st.subheader("Upload Files")
    st.session_state.workflow_file = st.file_uploader(
        "Upload ProCard workflow CSV",
        type=["csv"],
        key="workflow_upload",
    )
    st.session_state.zip_file = st.file_uploader(
        "Upload ZIP containing bank statement PDFs",
        type=["zip"],
        key="zip_upload",
    )
    st.session_state.batch_date = st.text_input(
        "Batch date (last day of month, YYMMDD)",
        value=st.session_state.batch_date,
    )
    if st.session_state.batch_date and not is_last_day_batch_date_YYMMDD(st.session_state.batch_date):
        render_helper_box("Batch date should be the last day of the month in YYMMDD format.", "warning")

    can_next = st.session_state.workflow_file is not None and st.session_state.zip_file is not None
    render_back_next(can_next=can_next)

elif step == 2:
    st.subheader("Parse & Match")
    if not st.session_state.processing_done:
        parse_info_box = st.empty()
        render_helper_box("Click the button below to begin parsing and matching.", "info", target=parse_info_box)
        st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
        if st.button("Start Parsing and Matching"):
            parse_info_box.empty()
            with st.spinner("Processing files..."):
                run_processing_pipeline()
            st.rerun()

    if st.session_state.processing_done:
        summary = st.session_state.summary or {}

        bank_review = st.session_state.bank_review if st.session_state.bank_review is not None else pd.DataFrame()
        workflow_review = st.session_state.workflow_review if st.session_state.workflow_review is not None else pd.DataFrame()
        workflow_only = st.session_state.workflow_only_rows if st.session_state.workflow_only_rows is not None else pd.DataFrame()

        bank_tx_count = len(bank_review)
        workflow_tx_count = 0
        if not workflow_review.empty:
            workflow_tx_count = int((
                (~workflow_review['Date'].map(is_blank))
                & (~workflow_review['Amount'].map(is_blank))
            ).sum())

        previous_cycle_count = 0
        if not workflow_only.empty and 'Warning Missing' in workflow_only.columns:
            previous_cycle_count = int(
                workflow_only['Warning Missing']
                .fillna('')
                .astype(str)
                .str.startswith('Likely on previous statement')
                .sum()
            )

        pdf_total = float(summary.get('pdf_parsed_total', 0.0))
        comparable_workflow_total = float(summary.get('comparable_workflow_total', 0.0))
        previous_cycle_total = abs(float(summary.get('excluded_previous_statement_total', 0.0)))
        total_difference = float(summary.get('total_difference', pdf_total - comparable_workflow_total))
        statement_cycle = str(summary.get('statement_cycle', '') or '').strip()

        totals_match = abs(total_difference) < 0.005
        if totals_match:
            render_helper_box("Processing complete. No errors detected.", "success")
        else:
            render_helper_box("Processing complete. Errors detected: totals do not match.", "error")

        if statement_cycle:
            st.markdown("<div style='height: 0.35rem;'></div>", unsafe_allow_html=True)
            render_helper_box(
                f"<strong>Reconciliation Snapshot</strong><br>Current Billing Cycle (from PDFs): {statement_cycle}",
                "prime",
            )
            st.markdown("<div style='height: 0.75rem;'></div>", unsafe_allow_html=True)

        if not totals_match:
            workflow_only_in_scope = pd.DataFrame(columns=workflow_only.columns)
            if not workflow_only.empty:
                warning_series = workflow_only['Warning Missing'].fillna('').astype(str) if 'Warning Missing' in workflow_only.columns else pd.Series(["" for _ in range(len(workflow_only))])
                out_of_date_mask = warning_series.str.startswith('Likely on previous statement') | warning_series.str.startswith('Likely on next statement')
                workflow_only_in_scope = workflow_only.loc[~out_of_date_mask].copy()

            bank_cmp = bank_review.copy()
            workflow_cmp = workflow_review.copy()

            if not bank_cmp.empty:
                bank_cmp['_key'] = bank_cmp.apply(lambda r: build_row_key(r.get('Date', ''), r.get('Amount', ''), r.get('Transaction', '')), axis=1)
            if not workflow_cmp.empty:
                workflow_cmp['_key'] = workflow_cmp.apply(lambda r: build_row_key(r.get('Date', ''), r.get('Amount', ''), r.get('Merchant', '')), axis=1)

            if not workflow_only.empty and not workflow_cmp.empty:
                excluded_outdated_keys = set(
                    workflow_only.loc[
                        workflow_only['Warning Missing'].fillna('').astype(str).str.startswith('Likely on previous statement') |
                        workflow_only['Warning Missing'].fillna('').astype(str).str.startswith('Likely on next statement')
                    ].apply(lambda r: build_row_key(r.get('Date', ''), r.get('Amount', ''), r.get('Merchant', '')), axis=1).tolist()
                )
                workflow_cmp = workflow_cmp.loc[~workflow_cmp['_key'].isin(excluded_outdated_keys)].copy()

            bank_counter = Counter(bank_cmp['_key'].tolist()) if '_key' in bank_cmp.columns else Counter()
            workflow_counter = Counter(workflow_cmp['_key'].tolist()) if '_key' in workflow_cmp.columns else Counter()

            bank_only_counter = bank_counter - workflow_counter

            bank_only_rows = select_rows_by_key_counts(bank_cmp, '_key', bank_only_counter)
            if '_key' in bank_only_rows.columns:
                bank_only_rows = bank_only_rows.drop(columns=['_key'])

            if not workflow_only_in_scope.empty:
                st.markdown("<div style='height: 0.5rem;'></div>", unsafe_allow_html=True)
                st.write("Mismatched Workflow rows (not excluded for out-of-date)")
                st.dataframe(workflow_only_in_scope, width='stretch')

            if not bank_only_rows.empty:
                st.markdown("<div style='height: 0.5rem;'></div>", unsafe_allow_html=True)
                st.write("Mismatched PDF rows not found in Workflow")
                st.dataframe(bank_only_rows, width='stretch')

            if workflow_only_in_scope.empty and bank_only_rows.empty:
                render_helper_box("Totals differ, but no in-scope row-level mismatches were identified by key match.", "warning")

        m1, m2, m3 = st.columns(3)
        m1.metric("PDF Transactions", f"{bank_tx_count:,}")
        m2.metric("Workflow Transactions", f"{workflow_tx_count:,}")
        m3.metric("Likely Previous Billing Cycle (Workflow)", f"{previous_cycle_count:,}")

        t1, t2, t3 = st.columns(3)
        t1.metric("PDF Total", f"${pdf_total:,.2f}")
        t2.metric("Workflow Total (Comparable)", f"${comparable_workflow_total:,.2f}")
        t3.metric("Amount Difference (PDF - Workflow)", f"${total_difference:,.2f}")

        st.caption(f"Workflow likely previous billing cycle amount (shown as positive): ${previous_cycle_total:,.2f}")

    render_back_next(can_next=st.session_state.processing_done)

elif step == 3:
    st.subheader("FOAPAL Review")
    if st.session_state.teams_workflow is None:
        render_helper_box("Run Parse & Match first.", "warning")
        st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
        render_back_next(can_next=False)
    else:
        required_cols = ['Fund Code:', 'Organization:', 'Account:', 'Program:', 'AD Code:']
        st.session_state.teams_workflow = normalize_editable_foapal(
            st.session_state.teams_workflow,
            required_cols,
        )
        foapal_option_source = (
            st.session_state.teams_workflow_original
            if st.session_state.teams_workflow_original is not None
            else st.session_state.teams_workflow
        )
        dropdown_options = build_foapal_dropdown_options(foapal_option_source, required_cols)
        missing_ids = list(st.session_state.teams_workflow.index[missing_required_mask(
            st.session_state.teams_workflow,
            required_cols,
        )])
        clamp_review_position(missing_ids)

        if not missing_ids:
            render_helper_box("No missing FOAPAL values found.", "success")
            st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
        else:
            current_row_id = missing_ids[st.session_state.foapal_current_pos]
            current_row = st.session_state.teams_workflow.loc[current_row_id]

            render_helper_box(f"{len(missing_ids)} row(s) still need FOAPAL values.", "warning")
            st.markdown("<div style='height: 0.35rem;'></div>", unsafe_allow_html=True)
            st.caption(f"Reviewing row {st.session_state.foapal_current_pos + 1} of {len(missing_ids)}")

            info_cols = st.columns(4)
            info_cols[0].metric("Date", str(current_row.get('Date of Transaction:', '')))
            info_cols[1].metric("Merchant", str(current_row.get('Merchant:', '')))
            info_cols[2].metric("Card Holder", str(current_row.get('Card Holder:', '')))
            info_cols[3].metric("Amount", str(current_row.get('Amount: $', '')))

            with st.form(f"foapal_form_{current_row_id}"):
                col1, col2, col3, col4, col5 = st.columns(5)
                with col1:
                    fund_code = selectbox_value(
                        "Fund",
                        current_row.get('Fund Code:', ''),
                        dropdown_options['Fund Code:'],
                    )
                with col2:
                    organization = selectbox_value(
                        "Org",
                        current_row.get('Organization:', ''),
                        dropdown_options['Organization:'],
                    )
                with col3:
                    account = selectbox_value(
                        "Account",
                        current_row.get('Account:', ''),
                        dropdown_options['Account:'],
                    )
                with col4:
                    ad_code = selectbox_value(
                        "AD Code",
                        current_row.get('AD Code:', ''),
                        dropdown_options['AD Code:'],
                    )
                with col5:
                    program = selectbox_value(
                        "Program",
                        current_row.get('Program:', ''),
                        dropdown_options['Program:'],
                        allow_outside_current=False,
                    )
                submitted = st.form_submit_button("Save FOAPAL for This Row")

            nav1, nav_spacer, nav2 = st.columns([1, 10, 1])
            with nav1:
                if st.button("Previous Missing Row", disabled=st.session_state.foapal_current_pos == 0):
                    st.session_state.foapal_current_pos -= 1
                    st.rerun()
            with nav2:
                if st.button(
                    "Next Missing Row",
                    disabled=st.session_state.foapal_current_pos >= len(missing_ids) - 1,
                ):
                    st.session_state.foapal_current_pos += 1
                    st.rerun()

            if submitted:
                updates = {
                    'Fund Code:': fund_code,
                    'Organization:': organization,
                    'Account:': account,
                    'Program:': program,
                    'AD Code:': ad_code,
                }
                for col, value in updates.items():
                    st.session_state.teams_workflow.at[current_row_id, col] = "" if is_blank(value) else str(value).strip()

                refreshed_missing_ids = list(st.session_state.teams_workflow.index[missing_required_mask(
                    st.session_state.teams_workflow,
                    required_cols,
                )])
                if current_row_id not in refreshed_missing_ids and refreshed_missing_ids:
                    st.session_state.foapal_current_pos = min(
                        st.session_state.foapal_current_pos,
                        len(refreshed_missing_ids) - 1,
                    )
                elif not refreshed_missing_ids:
                    st.session_state.foapal_current_pos = 0
                render_helper_box("FOAPAL saved.", "success")
                st.rerun()

        if missing_ids:
            st.divider()
            render_helper_box("You can continue now and review the remaining FOAPAL values after download.", "warning")
            st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
            if st.button("Review Rest After Download"):
                st.session_state.foapal_review_skipped = True
                go_next()
                st.rerun()

        has_missing_required = missing_required_mask(st.session_state.teams_workflow, required_cols).any()
        can_next = (not has_missing_required) or st.session_state.foapal_review_skipped
        if has_missing_required and not st.session_state.foapal_review_skipped:
            render_helper_box("All required FOAPAL fields must be filled before proceeding.", "error")
            st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
        render_back_next(can_next=can_next)

elif step == 4:
    st.subheader("Download Output")
    if st.session_state.teams_workflow is None:
        render_helper_box("No processed data available yet.", "warning")
        st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
        render_back_next(can_next=False)
    else:
        required_cols = ['Fund Code:', 'Organization:', 'Account:', 'Program:', 'AD Code:']
        has_missing_required = missing_required_mask(st.session_state.teams_workflow, required_cols).any()
        if has_missing_required and not st.session_state.foapal_review_skipped:
            render_helper_box("Cannot generate final output until all required FOAPAL cells are filled.", "error")
            st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)
            render_back_next(can_next=False)
            st.stop()

        if has_missing_required and st.session_state.foapal_review_skipped:
            render_helper_box("Generating output with missing FOAPAL values because review was skipped.", "warning")
            st.markdown("<div style='height: 0.45rem;'></div>", unsafe_allow_html=True)

        current_output_signature = build_final_output_signature()
        if (
            not st.session_state.final_output_bytes
            or st.session_state.final_output_signature != current_output_signature
        ):
            st.session_state.final_output_bytes = build_final_workbook_bytes()
            st.session_state.final_output_name = f"Final_ProCard_Upload_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            st.session_state.final_output_signature = current_output_signature

        st.download_button(
            label="Download Final Workbook",
            data=st.session_state.final_output_bytes,
            file_name=st.session_state.final_output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click="ignore",
        )
        st.write("FILE FEED Preview")
        file_feed_preview = build_file_feed_from_teams_workflow(
            sanitize_for_output(st.session_state.teams_workflow),
            st.session_state.batch_date,
        )
        st.dataframe(sanitize_for_output(file_feed_preview), width='stretch')
        render_back_next(can_next=False)