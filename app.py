#!/usr/bin/env python3
"""
ממשק Streamlit להמרת טופסי דיווח איכות מים לטופס קליטה.

הרצה:
    streamlit run app.py
"""

import io
import os
import tempfile
from datetime import datetime

import openpyxl
import pandas as pd
import streamlit as st

# Import backend functions from the conversion script
from convert_report_to_intake import (
    load_param_table,
    load_csv_lookup,
    load_well_memory,
    save_well_memory,
    load_historical_data,
    convert_report,
    write_intake_file,
    write_error_report,
)

# ---------------------------------------------------------------------------
# Paths — resolved relative to this script so the app works from any CWD
# ---------------------------------------------------------------------------

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
WELL_MEMORY_PATH = os.path.join(_SCRIPT_DIR, 'config', 'well_codes_memory.csv')


# ---------------------------------------------------------------------------
# Helper: find file in config/ or root
# ---------------------------------------------------------------------------

def find_file(*candidates):
    """Return the first existing path from candidates, or None.
    Searches relative to CWD and relative to the script's directory."""
    for c in candidates:
        if os.path.exists(c):
            return c
        alt = os.path.join(_SCRIPT_DIR, c)
        if os.path.exists(alt):
            return alt
    return None


def _load_temp_xlsx(uploaded_file, loader_fn, *args):
    """Write an uploaded file to a temp path, call loader_fn(path, *args), clean up."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    try:
        return loader_fn(tmp_path, *args)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _load_temp_csv(uploaded_file, loader_fn, *args):
    """Write an uploaded CSV to a temp path, call loader_fn(path, *args), clean up."""
    with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='wb') as tmp:
        tmp.write(uploaded_file.read())
        tmp_path = tmp.name
    try:
        return loader_fn(tmp_path, *args)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


def _convert_file_bytes(stored_bytes, *args, **kwargs):
    """Write stored bytes to a temp file, run convert_report(), clean up."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp.write(stored_bytes)
        tmp_path = tmp.name
    try:
        return convert_report(tmp_path, *args, **kwargs)
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)


# ---------------------------------------------------------------------------
# App config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="המרת טופסי דיווח איכות מים",
    page_icon="💧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# RTL support
st.markdown("""
<style>
    .stApp { direction: rtl; }
    .stMarkdown, .stText, .stAlert, .stDataFrame { direction: rtl; text-align: right; }
    h1, h2, h3, p, li, td, th, label, .stSelectbox, .stMultiSelect {
        direction: rtl; text-align: right;
    }
    .dataframe th, .dataframe td { text-align: right !important; }
    section[data-testid="stSidebar"] { direction: rtl; text-align: right; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Session state defaults
# ---------------------------------------------------------------------------

if 'ref_labs' not in st.session_state:
    st.session_state.ref_labs = None
if 'ref_samplers' not in st.session_state:
    st.session_state.ref_samplers = None
if 'param_map' not in st.session_state:
    st.session_state.param_map = None
if 'well_memory' not in st.session_state:
    st.session_state.well_memory = {}
if 'historical_data' not in st.session_state:
    st.session_state.historical_data = None
if 'conversion_results' not in st.session_state:
    st.session_state.conversion_results = None


# ---------------------------------------------------------------------------
# Sidebar — Reference file management
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("קבצי ייחוס")

    # Parameter table
    st.subheader("טבלת פרמטרים")
    params_file = st.file_uploader(
        "param_table.xlsx", type=['xlsx'], key='params_upload')
    if params_file:
        try:
            st.session_state.param_map = _load_temp_xlsx(params_file, load_param_table)
            st.success(f"נטענו {len(st.session_state.param_map)} פרמטרים")
        except Exception as e:
            st.error(f"שגיאה בטעינת טבלת פרמטרים: {e}")
    elif st.session_state.param_map is None:
        _p = find_file('config/param_table.xlsx', 'param_table.xlsx')
        if _p:
            try:
                st.session_state.param_map = load_param_table(_p)
                st.info(f"נטען אוטומטית: {len(st.session_state.param_map)} פרמטרים")
            except Exception as e:
                st.error(f"שגיאה בטעינת טבלת פרמטרים: {e}")
        else:
            st.warning("נדרש קובץ טבלת פרמטרים")

    st.divider()

    # Lab codes
    st.subheader("קודי מעבדות")
    labs_file = st.file_uploader("lab_codes.csv", type=['csv'], key='labs_upload')
    if labs_file:
        try:
            st.session_state.ref_labs = _load_temp_csv(
                labs_file, load_csv_lookup, 'שם מעבדה', 'קוד מעבדה')
            st.success(f"נטענו {len(st.session_state.ref_labs)} מעבדות")
        except Exception as e:
            st.error(f"שגיאה בטעינת קודי מעבדות: {e}")
    elif st.session_state.ref_labs is None:
        _l = find_file('config/lab_codes.csv', 'lab_codes.csv')
        if _l:
            st.session_state.ref_labs = load_csv_lookup(_l, 'שם מעבדה', 'קוד מעבדה')
            st.info(f"נטען אוטומטית: {len(st.session_state.ref_labs)} מעבדות")

    # Sampler codes
    st.subheader("קודי חברות דיגום")
    samplers_file = st.file_uploader(
        "sampler_codes.csv", type=['csv'], key='samplers_upload')
    if samplers_file:
        try:
            st.session_state.ref_samplers = _load_temp_csv(
                samplers_file, load_csv_lookup, 'שם חברת דיגום', 'קוד חברת דיגום')
            st.success(f"נטענו {len(st.session_state.ref_samplers)} חברות")
        except Exception as e:
            st.error(f"שגיאה בטעינת קודי חברות דיגום: {e}")
    elif st.session_state.ref_samplers is None:
        _s = find_file('config/sampler_codes.csv', 'sampler_codes.csv')
        if _s:
            st.session_state.ref_samplers = load_csv_lookup(
                _s, 'שם חברת דיגום', 'קוד חברת דיגום')
            st.info(f"נטען אוטומטית: {len(st.session_state.ref_samplers)} חברות")

    st.divider()

    # Historical data
    st.subheader("נתונים היסטוריים")
    hist_file = st.file_uploader(
        "קובץ היסטורי (xlsx)", type=['xlsx'], key='hist_upload')
    if hist_file:
        try:
            st.session_state.historical_data = _load_temp_xlsx(
                hist_file, load_historical_data)
            st.success(f"נטענו {len(st.session_state.historical_data)} רשומות היסטוריות")
        except Exception as e:
            st.error(f"שגיאה בטעינת נתונים היסטוריים: {e}")

    st.divider()

    # Well memory — always load from the canonical absolute path
    if not st.session_state.well_memory and os.path.exists(WELL_MEMORY_PATH):
        st.session_state.well_memory = load_well_memory(WELL_MEMORY_PATH)

    st.subheader("זיכרון קודי קידוחים")
    if st.session_state.well_memory:
        st.info(f"{len(st.session_state.well_memory)} רשומות בזיכרון")
        if st.button("נקה זיכרון"):
            st.session_state.well_memory = {}
            st.rerun()
    else:
        st.caption("ריק — ייבנה אוטומטית מהקלדות")


# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------

st.title("💧 המרת טופסי דיווח איכות מים")
st.caption("המרה מטופסי דיווח של אתרי דלק לטופס קליטה למערכת המידע של רשות המים")

# Check prerequisites
if st.session_state.param_map is None:
    st.error("נדרש קובץ טבלת פרמטרים. העלה אותו בסרגל הצד.")
    st.stop()

# ---------------------------------------------------------------------------
# Step 1: Upload reporting files
# ---------------------------------------------------------------------------

st.header("① העלאת טופסי דיווח")

uploaded_files = st.file_uploader(
    "גרור או בחר קבצי דיווח (xlsx)",
    type=['xlsx'],
    accept_multiple_files=True,
    key='report_upload',
)

if not uploaded_files:
    st.info("העלה טופס דיווח אחד או יותר כדי להתחיל")
    st.stop()

st.success(f"הועלו {len(uploaded_files)} קבצים")

# ---------------------------------------------------------------------------
# Step 2: Process files
# ---------------------------------------------------------------------------

if st.button("🔄 עבד קבצים", type="primary", use_container_width=True):
    all_rows = []
    all_results = []
    missing_wells = []
    missing_labs = []
    missing_samplers = []

    progress = st.progress(0, text="מעבד קבצים...")

    for i, uploaded in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files),
                          text=f"מעבד: {uploaded.name}")

        rows, errors, warnings = _convert_file_bytes(
            uploaded.read(),
            st.session_state.param_map,
            ref_labs=st.session_state.ref_labs,
            ref_samplers=st.session_state.ref_samplers,
            interactive=False,
            well_memory=st.session_state.well_memory,
            historical_data=st.session_state.historical_data,
        )

        for err in errors:
            if 'קוד קידוח חסר' in err or 'קוד קידוח לא מספרי' in err:
                missing_wells.append((uploaded.name, err))
            if 'קוד מעבדה חסר' in err:
                missing_labs.append((uploaded.name, err))
            if 'קוד חברת דיגום חסר' in err:
                missing_samplers.append((uploaded.name, err))

        all_rows.extend(rows)
        all_results.append((uploaded.name, errors, warnings, len(rows)))

    progress.empty()

    st.session_state.conversion_results = {
        'all_rows': all_rows,
        'all_results': all_results,
        'missing_wells': missing_wells,
        'missing_labs': missing_labs,
        'missing_samplers': missing_samplers,
        'uploaded_files': [(f.name, f.getvalue()) for f in uploaded_files],
    }
    st.rerun()

# ---------------------------------------------------------------------------
# Step 3: Handle missing codes
# ---------------------------------------------------------------------------

results = st.session_state.conversion_results
if results is None:
    st.stop()

all_rows = results['all_rows']
all_results = results['all_results']
missing_wells = results['missing_wells']
missing_labs = results.get('missing_labs', [])
missing_samplers = results.get('missing_samplers', [])

if missing_labs or missing_samplers or missing_wells:
    st.header("② השלמת נתונים חסרים")

    # Lab codes
    lab_inputs = {}
    if missing_labs:
        st.subheader("קודי מעבדה חסרים")
        st.warning(f"נמצאו {len(missing_labs)} קבצים עם קוד מעבדה חסר")
        for idx, (fname, err_msg) in enumerate(missing_labs):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.text(f"📁 {fname}: {err_msg}")
            with col2:
                code = st.text_input("קוד מעבדה", key=f"lab_code_{idx}",
                                     placeholder="למשל: 6")
                if code:
                    lab_inputs[idx] = code

    # Sampler codes
    sampler_inputs = {}
    if missing_samplers:
        st.subheader("קודי חברות דיגום חסרים")
        st.warning(f"נמצאו {len(missing_samplers)} קבצים עם קוד חברת דיגום חסר")
        for idx, (fname, err_msg) in enumerate(missing_samplers):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.text(f"📁 {fname}: {err_msg}")
            with col2:
                code = st.text_input("קוד חברת דיגום", key=f"sampler_code_{idx}",
                                     placeholder="למשל: 23")
                if code:
                    sampler_inputs[idx] = code

    # Well codes
    well_inputs = {}
    if missing_wells:
        st.subheader("קודי קידוח חסרים")
        st.warning(f"נמצאו {len(missing_wells)} קידוחים עם קוד חסר או לא תקין")
        for idx, (fname, err_msg) in enumerate(missing_wells):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.text(f"📁 {fname}: {err_msg}")
            with col2:
                code = st.text_input("קוד (8 ספרות)", key=f"well_code_{idx}",
                                     max_chars=8, placeholder="12345678")
                if code:
                    well_inputs[idx] = code

    if st.button("🔄 עבד מחדש עם הנתונים שהוקלדו", use_container_width=True):
        # Build per-file overrides for lab/sampler
        file_overrides = {}

        for idx, code_str in lab_inputs.items():
            try:
                file_overrides.setdefault(missing_labs[idx][0], {})['lab'] = int(code_str)
            except ValueError:
                st.error(f"קוד מעבדה לא מספרי: '{code_str}'")

        for idx, code_str in sampler_inputs.items():
            try:
                file_overrides.setdefault(missing_samplers[idx][0], {})['sampler'] = int(code_str)
            except ValueError:
                st.error(f"קוד חברת דיגום לא מספרי: '{code_str}'")

        # Process well codes and update memory
        new_memory_entries = {}
        for idx, code_str in well_inputs.items():
            if len(code_str) != 8 or not code_str.isdigit():
                st.error(f"קוד קידוח חייב להיות 8 ספרות: '{code_str}'")
                continue
            code_int = int(code_str)
            fname = missing_wells[idx][0]
            for stored_name, stored_bytes in results['uploaded_files']:
                if stored_name != fname:
                    continue
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                    tmp.write(stored_bytes)
                    tmp_path = tmp.name
                try:
                    wb = openpyxl.load_workbook(tmp_path, data_only=True)
                    try:
                        site_name = wb.active.cell(row=2, column=2).value or '?'
                    finally:
                        wb.close()
                finally:
                    if os.path.exists(tmp_path):
                        os.unlink(tmp_path)
                err_msg = missing_wells[idx][1]
                if "'" in err_msg:
                    well_name = err_msg.split("'")[1]
                    st.session_state.well_memory[(site_name, well_name)] = code_int
                    new_memory_entries[(site_name, well_name)] = code_int
                break

        if new_memory_entries:
            save_well_memory(st.session_state.well_memory, WELL_MEMORY_PATH)
            st.success(f"נשמרו {len(new_memory_entries)} קודי קידוח חדשים לזיכרון")

        # Re-process all files with overrides
        all_rows = []
        all_results = []
        missing_wells = []
        missing_labs = []
        missing_samplers = []

        for stored_name, stored_bytes in results['uploaded_files']:
            overrides = file_overrides.get(stored_name, {})
            rows, errors, warnings = _convert_file_bytes(
                stored_bytes,
                st.session_state.param_map,
                ref_labs=st.session_state.ref_labs,
                ref_samplers=st.session_state.ref_samplers,
                interactive=False,
                well_memory=st.session_state.well_memory,
                historical_data=st.session_state.historical_data,
                lab_code_override=overrides.get('lab'),
                sampler_code_override=overrides.get('sampler'),
            )

            for err in errors:
                if 'קוד קידוח חסר' in err or 'קוד קידוח לא מספרי' in err:
                    missing_wells.append((stored_name, err))
                if 'קוד מעבדה חסר' in err:
                    missing_labs.append((stored_name, err))
                if 'קוד חברת דיגום חסר' in err:
                    missing_samplers.append((stored_name, err))

            all_rows.extend(rows)
            all_results.append((stored_name, errors, warnings, len(rows)))

        st.session_state.conversion_results = {
            'all_rows': all_rows,
            'all_results': all_results,
            'missing_wells': missing_wells,
            'missing_labs': missing_labs,
            'missing_samplers': missing_samplers,
            'uploaded_files': results['uploaded_files'],
        }
        st.rerun()

# ---------------------------------------------------------------------------
# Step 4: Results summary
# ---------------------------------------------------------------------------

st.header("③ תוצאות")

total_rows = len(all_rows)
total_errors = sum(len(e) for _, e, _, _ in all_results)
total_warnings = sum(len(w) for _, _, w, _ in all_results)

col1, col2, col3, col4 = st.columns(4)
col1.metric("קבצים", len(all_results))
col2.metric("שורות קליטה", total_rows)
col3.metric("שגיאות", total_errors, delta_color="inverse")
col4.metric("אזהרות", total_warnings, delta_color="inverse")

for fname, errors, warnings, row_count in all_results:
    with st.expander(
            f"{'🔴' if errors else '🟢'} {fname} — "
            f"{row_count} שורות, {len(errors)} שגיאות, {len(warnings)} אזהרות"):
        if errors:
            for err in errors:
                st.error(f"✗ {err}")
        if warnings:
            for warn in warnings:
                st.warning(f"⚠ {warn}")
        if not errors and not warnings:
            st.success("עבר בהצלחה")

# ---------------------------------------------------------------------------
# Step 5: Preview and download
# ---------------------------------------------------------------------------

if total_rows > 0:
    st.header("④ תצוגה מקדימה והורדה")

    preview_data = []
    for row in all_rows[:50]:
        date_val = row['C']
        date_str = date_val.strftime('%d.%m.%Y') if isinstance(date_val, datetime) else date_val
        preview_data.append({
            'מקור מים': row['A'],
            'מס ש"ה': row['B'],
            'תאריך': date_str,
            'מוסד דוגם': row['E'],
            'תוצאה': row['G'],
            'פרמטר': row['H'],
            'מעבדה': row['I'],
        })

    st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)
    if total_rows > 50:
        st.caption(f"מוצגות 50 שורות מתוך {total_rows}")

    # Generate output files in memory (no temp disk files)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    intake_buf = io.BytesIO()
    write_intake_file(all_rows, intake_buf)

    error_buf = io.BytesIO()
    write_error_report(all_results, error_buf)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label=f"📥 הורד קובץ קליטה ({total_rows} שורות)",
            data=intake_buf.getvalue(),
            file_name=f"קליטה_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            label="📋 הורד דוח שגיאות",
            data=error_buf.getvalue(),
            file_name=f"שגיאות_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
elif not missing_wells and not missing_labs and not missing_samplers:
    st.error("לא נוצרו שורות קליטה. בדוק את דוח השגיאות.")
