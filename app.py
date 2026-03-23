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
# Helper: find file in config/ or root
# ---------------------------------------------------------------------------

def find_file(*candidates):
    """Return the first existing path from candidates, or None.
    Searches relative to CWD and relative to the script's directory."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for c in candidates:
        if os.path.exists(c):
            return c
        # Also try relative to script directory
        alt = os.path.join(script_dir, c)
        if os.path.exists(alt):
            return alt
    return None

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
    /* Fix dataframe display */
    .dataframe th, .dataframe td { text-align: right !important; }
    /* Sidebar */
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
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(params_file.read())
            tmp_path = tmp.name
        st.session_state.param_map = load_param_table(tmp_path)
        os.unlink(tmp_path)
        st.success(f"נטענו {len(st.session_state.param_map)} פרמטרים")
    elif st.session_state.param_map is None:
        # Try loading from default location
        _p = find_file('config/param_table.xlsx', 'param_table.xlsx')
        if _p:
            st.session_state.param_map = load_param_table(_p)
            st.info(f"נטען אוטומטית: {len(st.session_state.param_map)} פרמטרים")
        else:
            st.warning("נדרש קובץ טבלת פרמטרים")

    st.divider()

    # Lab codes
    st.subheader("קודי מעבדות")
    labs_file = st.file_uploader("lab_codes.csv", type=['csv'], key='labs_upload')
    if labs_file:
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='wb') as tmp:
            tmp.write(labs_file.read())
            tmp_path = tmp.name
        st.session_state.ref_labs = load_csv_lookup(tmp_path, 'שם מעבדה', 'קוד מעבדה')
        os.unlink(tmp_path)
        st.success(f"נטענו {len(st.session_state.ref_labs)} מעבדות")
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
        with tempfile.NamedTemporaryFile(suffix='.csv', delete=False, mode='wb') as tmp:
            tmp.write(samplers_file.read())
            tmp_path = tmp.name
        st.session_state.ref_samplers = load_csv_lookup(
            tmp_path, 'שם חברת דיגום', 'קוד חברת דיגום')
        os.unlink(tmp_path)
        st.success(f"נטענו {len(st.session_state.ref_samplers)} חברות")
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
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(hist_file.read())
            tmp_path = tmp.name
        st.session_state.historical_data = load_historical_data(tmp_path)
        os.unlink(tmp_path)
        st.success(f"נטענו {len(st.session_state.historical_data)} רשומות היסטוריות")

    st.divider()

    # Well memory
    st.subheader("זיכרון קודי קידוחים")
    if st.session_state.well_memory:
        st.info(f"{len(st.session_state.well_memory)} רשומות בזיכרון")
        if st.button("נקה זיכרון"):
            st.session_state.well_memory = {}
            st.rerun()
    else:
        st.caption("ריק — ייבנה אוטומטית מהקלדות")

    # Load from file
    _wm = find_file('config/well_codes_memory.csv', 'well_codes_memory.csv')
    if _wm and not st.session_state.well_memory:
        st.session_state.well_memory = load_well_memory(_wm)


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
# Step 2: Process files and detect missing well codes
# ---------------------------------------------------------------------------

if st.button("🔄 עבד קבצים", type="primary", use_container_width=True):
    all_rows = []
    all_results = []
    missing_wells = []  # [(file, site, well_name, col_idx, raw_code)]

    progress = st.progress(0, text="מעבד קבצים...")

    for i, uploaded in enumerate(uploaded_files):
        progress.progress((i + 1) / len(uploaded_files),
                          text=f"מעבד: {uploaded.name}")

        # Save to temp file
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        rows, errors, warnings = convert_report(
            tmp_path,
            st.session_state.param_map,
            ref_labs=st.session_state.ref_labs,
            ref_samplers=st.session_state.ref_samplers,
            interactive=False,
            well_memory=st.session_state.well_memory,
            historical_data=st.session_state.historical_data,
        )

        os.unlink(tmp_path)

        # Detect missing well codes from errors
        for err in errors:
            if 'קוד קידוח חסר' in err or 'קוד קידוח לא מספרי' in err:
                missing_wells.append((uploaded.name, err))

        all_rows.extend(rows)
        all_results.append((uploaded.name, errors, warnings, len(rows)))

    progress.empty()

    st.session_state.conversion_results = {
        'all_rows': all_rows,
        'all_results': all_results,
        'missing_wells': missing_wells,
        'uploaded_files': [(f.name, f.getvalue()) for f in uploaded_files],
    }
    st.rerun()

# ---------------------------------------------------------------------------
# Step 3: Show results and handle missing well codes
# ---------------------------------------------------------------------------

results = st.session_state.conversion_results
if results is None:
    st.stop()

all_rows = results['all_rows']
all_results = results['all_results']
missing_wells = results['missing_wells']

# --- Missing well codes interactive resolution ---

if missing_wells:
    st.header("② השלמת קודי קידוח חסרים")
    st.warning(f"נמצאו {len(missing_wells)} קידוחים עם קוד חסר או לא תקין")

    well_inputs = {}
    for idx, (fname, err_msg) in enumerate(missing_wells):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.text(f"📁 {fname}: {err_msg}")
        with col2:
            code = st.text_input(
                "קוד (8 ספרות)",
                key=f"well_code_{idx}",
                max_chars=8,
                placeholder="12345678",
            )
            if code:
                well_inputs[idx] = code

    if st.button("🔄 עבד מחדש עם הקודים שהוקלדו", use_container_width=True):
        # Parse entered codes and update well memory
        import openpyxl as _openpyxl

        new_memory_entries = {}
        for idx, code_str in well_inputs.items():
            try:
                code_int = int(code_str)
                if len(code_str) == 8:
                    # Extract site and well name from the uploaded file
                    fname = missing_wells[idx][0]
                    for stored_name, stored_bytes in results['uploaded_files']:
                        if stored_name == fname:
                            with tempfile.NamedTemporaryFile(
                                    suffix='.xlsx', delete=False) as tmp:
                                tmp.write(stored_bytes)
                                tmp_path = tmp.name
                            wb = _openpyxl.load_workbook(tmp_path, data_only=True)
                            ws = wb.active
                            site_name = ws.cell(row=2, column=2).value or '?'
                            # Find the well name from error message
                            err_msg = missing_wells[idx][1]
                            if "'" in err_msg:
                                well_name = err_msg.split("'")[1]
                                st.session_state.well_memory[
                                    (site_name, well_name)] = code_int
                                new_memory_entries[(site_name, well_name)] = code_int
                            wb.close()
                            os.unlink(tmp_path)
                            break
                else:
                    st.error(f"קוד {code_str} אינו 8 ספרות")
            except ValueError:
                st.error(f"ערך לא מספרי: {code_str}")

        if new_memory_entries:
            # Save memory to disk
            save_well_memory(st.session_state.well_memory, 'config/well_codes_memory.csv')
            st.success(f"נשמרו {len(new_memory_entries)} קודים חדשים לזיכרון")

        # Re-process all files
        all_rows = []
        all_results = []
        missing_wells = []

        for stored_name, stored_bytes in results['uploaded_files']:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                tmp.write(stored_bytes)
                tmp_path = tmp.name

            rows, errors, warnings = convert_report(
                tmp_path,
                st.session_state.param_map,
                ref_labs=st.session_state.ref_labs,
                ref_samplers=st.session_state.ref_samplers,
                interactive=False,
                well_memory=st.session_state.well_memory,
                historical_data=st.session_state.historical_data,
            )
            os.unlink(tmp_path)

            for err in errors:
                if 'קוד קידוח חסר' in err or 'קוד קידוח לא מספרי' in err:
                    missing_wells.append((stored_name, err))

            all_rows.extend(rows)
            all_results.append((stored_name, errors, warnings, len(rows)))

        st.session_state.conversion_results = {
            'all_rows': all_rows,
            'all_results': all_results,
            'missing_wells': missing_wells,
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

# Per-file details
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

    # Preview
    preview_data = []
    for row in all_rows[:50]:
        date_val = row['C']
        if isinstance(date_val, datetime):
            date_str = date_val.strftime('%d.%m.%Y')
        else:
            date_str = date_val
        preview_data.append({
            'מקור מים': row['A'],
            'מס ש"ה': row['B'],
            'תאריך': date_str,
            'מוסד דוגם': row['E'],
            'תוצאה': row['G'],
            'פרמטר': row['H'],
            'מעבדה': row['I'],
        })

    st.dataframe(
        pd.DataFrame(preview_data),
        use_container_width=True,
        hide_index=True,
    )
    if total_rows > 50:
        st.caption(f"מוצגות 50 שורות מתוך {total_rows}")

    # Generate files for download
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Intake file
    intake_path = os.path.join(tempfile.gettempdir(), f'קליטה_{timestamp}.xlsx')
    write_intake_file(all_rows, intake_path)
    with open(intake_path, 'rb') as f:
        intake_bytes = f.read()
    os.unlink(intake_path)

    # Error report
    error_path = os.path.join(tempfile.gettempdir(), f'שגיאות_{timestamp}.xlsx')
    write_error_report(all_results, error_path)
    with open(error_path, 'rb') as f:
        error_bytes = f.read()
    os.unlink(error_path)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label=f"📥 הורד קובץ קליטה ({total_rows} שורות)",
            data=intake_bytes,
            file_name=f"קליטה_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            label=f"📋 הורד דוח שגיאות",
            data=error_bytes,
            file_name=f"שגיאות_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
elif not missing_wells:
    st.error("לא נוצרו שורות קליטה. בדוק את דוח השגיאות.")
