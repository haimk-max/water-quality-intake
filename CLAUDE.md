# CLAUDE.md — Water Quality Intake Converter

## Project Overview

Tool for the Israel Water Authority (רשות המים) that converts water quality field reporting forms (Excel) from fuel site monitoring into the Authority's standardized intake format for its information system.

The system handles validation, error detection, fuzzy name matching, interactive well code completion with persistent memory, BH lookup table resolution, historical anomaly detection, and produces both an intake file and an error report.

## Language Policy

- Code, variable names, comments: **English**
- UI text, error/warning messages, Excel headers, file labels: **Hebrew (RTL)**
- User communication: **Hebrew**

---

## Architecture

### Backend: `convert_report_to_intake.py`

CLI tool **and** importable Python module. All conversion logic lives here.

**Core Functions:**

| Function | Purpose |
|---|---|
| `convert_report()` | Main entry: reads reporting Excel → validates → returns intake rows + errors + warnings |
| `load_param_table()` | Loads `param_table.xlsx`: numeric code → text symbol mapping (830 params) |
| `load_csv_lookup()` | Generic CSV loader for lab/sampler reference files |
| `load_well_memory()` / `save_well_memory()` | Persistent well code memory (site+well_name → code). `save_well_memory` creates parent directory if missing. |
| `load_bh_lookup()` | Loads `Sites_missing_BH_codes.xlsx`: (site, well_name) → well_code lookup table |
| `lookup_bh_code()` | Searches BH table with mild fuzzy matching (threshold 0.9 for both site and well name, substring fallback) |
| `load_historical_data()` | Loads historical data for anomaly detection. Auto-detects two formats: intake format (headers row 1) and historical template (headers row 5) |
| `validate_name_code()` | Fuzzy matching of lab/sampler names against reference lists using `difflib.SequenceMatcher` and substring matching |
| `parse_measurement()` | Parses cell values: numbers, `"<X"` below-detection → 0, empty → skip |
| `parse_date()` | Extracts sampling date from row 3 (handles datetime objects and text in multiple formats) |
| `prompt_well_code()` | Interactive CLI prompt for missing well codes (8-digit validation) |
| `fuzzy_match()` | Returns scored matches above threshold, used by `validate_name_code()` |
| `write_intake_file()` | Writes output Excel in intake format (9 columns, DD.MM.YYYY date as text). Accepts file path or `io.BytesIO`. |
| `write_error_report()` | Writes error/warning Excel with color coding (red=error, yellow=warning). Accepts file path or `io.BytesIO`. |

**`convert_report()` parameters:**

```python
convert_report(
    report_path,
    param_map,
    ref_labs=None,
    ref_samplers=None,
    interactive=False,
    well_memory=None,
    historical_data=None,
    lab_code_override=None,       # Manual override when lab code missing
    sampler_code_override=None,   # Manual override when sampler code missing
    bh_lookup=None,               # BH well-code lookup table
)
```

**CLI Usage:**

```bash
# Basic
python convert_report_to_intake.py file1.xlsx file2.xlsx

# Interactive mode (prompts for missing well codes)
python convert_report_to_intake.py file.xlsx --interactive

# Full options
python convert_report_to_intake.py file.xlsx \
  --params param_table.xlsx \
  --labs config/lab_codes.csv \
  --samplers config/sampler_codes.csv \
  --historical historical_data.xlsx \
  --interactive \
  --well-memory config/well_codes_memory.csv \
  --output output.xlsx \
  --error-report errors.xlsx
```

**Auto-discovery:** If `--params`, `--labs`, `--samplers` are not specified, the script looks for default filenames in the report directory and current directory.

---

### Frontend: `app.py`

Streamlit web application wrapping the backend. Designed for a small internal team running on multiple machines.

**Key design decisions:**
- All paths resolved relative to script directory (`_SCRIPT_DIR`, `WELL_MEMORY_PATH`) — not CWD-dependent
- Temp file helpers (`_load_temp_xlsx`, `_load_temp_csv`, `_convert_file_bytes`) use `try/finally` for guaranteed cleanup
- Download files generated in `io.BytesIO` — no disk writes, no multi-user collisions
- Reference files loaded lazily into `session_state`; errors shown in sidebar

**UI Flow:**

1. **Sidebar** — Reference file management: parameter table, lab codes, sampler codes, BH lookup table, historical data, well memory. Files in `config/` or root are auto-loaded on startup.
2. **Step ①** — Upload one or more reporting form Excel files (drag & drop).
3. **Step ②** — If any codes are missing (lab, sampler, or well): inputs appear grouped by type. On submit, overrides are applied and files are re-processed. Well codes are saved to memory; lab/sampler codes are **not** saved (asked fresh each time).
4. **Step ③** — Results dashboard: metrics (files, rows, errors, warnings) + expandable per-file details.
5. **Step ④** — Preview table + download buttons for intake file and error report.

**Helper functions:**
- `find_file(*candidates)` — searches `_SCRIPT_DIR`-relative paths for reference files
- `_load_temp_xlsx(file, fn)` / `_load_temp_csv(file, fn)` — write uploaded file to temp, call fn, clean up
- `_convert_file_bytes(bytes, ...)` — write bytes to temp, call `convert_report()`, clean up

**Run:**

```bash
streamlit run app.py
```

**Deployed at:** `https://water-quality-intake-jejkg9bbhkcegvtepdjbxs.streamlit.app/`

---

## File Structure

```
water-quality-intake/
├── CLAUDE.md                           # This file — project guide
├── README.md                           # User-facing documentation
├── app.py                              # Streamlit frontend
├── convert_report_to_intake.py         # Backend engine
├── param_table.xlsx                    # Parameter mapping table (830 rows)
├── Sites_missing_BH_codes.xlsx         # BH well-code lookup table (~392 entries)
├── requirements.txt                    # Python dependencies
├── .gitignore
├── .streamlit/
│   └── config.toml                     # Streamlit theme + upload size
├── config/
│   ├── lab_codes.csv                   # Lab name → code
│   ├── sampler_codes.csv               # Sampler name → code
│   └── well_codes_memory.csv           # Persistent well code memory (tracked in git)
└── test_files/
    ├── Copy_of_אשקלון_טופס_דיווח_דצמבר_2020.xlsx   # 5 wells, all codes present
    ├── דיווח_09_25.xlsx                              # 1 well, MISSING well code (resolved via BH table)
    └── טופס_דיווח_נס_ציונה_2025.xlsx                 # 1 well, code present
```

**Important:** CSV files must be inside `config/`, NOT in root — Streamlit Cloud reads root CSVs as requirements files and crashes.

---

## Data Formats

### Reporting Form (Input)

Excel file, single sheet named "דיווח מדדי שדה ומעבדה לאתר" or "גיליון1".

**Header section (rows 1-11):**

| Row | Col A | Col B | Col C | Col D+ |
|-----|-------|-------|-------|--------|
| 2 | "שם אתר הדיגום" | Site name | "-" | |
| 3 | "תאריך דיגום" | Date (text or "-") | Date (datetime) | |
| 4 | "מעבדה" | Lab name | Lab code (int) | |
| 5 | "חברת דיגום" | Sampler name | Sampler code (int) | |
| 6 | "שיטת דיגום" | Method name | Method code (int) | |
| 7 | "שם קידוח" | "-" | "-" | Well names (1-20 wells) |
| 8 | "קוד קידוח" | "-" | "-" | Well codes (8-digit int) |
| 9 | | "אמצעי שיקום" | "-" | Remediation codes |
| 10 | | "עומק עד פני המים" | "m" | Depth values (meters) |
| 11 | | "OIL LAYER THIKNESS" | "m" | Oil layer values |

**Data section (rows 13+):**

| Col A | Col B | Col C | Col D+ |
|-------|-------|-------|--------|
| Parameter code (int) | Parameter description | Units | Measurement values per well |

**Measurement values can be:**
- Numeric (int/float): actual measurement
- `"<X"` (string): below detection limit → converted to 0
- Empty/None: parameter not measured → row skipped (no intake row created)
- `0` (explicit): passed as-is to intake
- `"-"`: treated as empty

**Lookup tables** in columns F-W (varying start column) are for human reference only — the script ignores them. Well detection stops after 2 consecutive empty columns in rows 7-8.

### Intake Form (Output)

Flat Excel table. One row per measurement per well.

| Column | Header | Source | Notes |
|--------|--------|--------|-------|
| A | מקור מים | Constant | Always `5` for fuel sites |
| B | מס ש"ה | Row 8 per well column | 8-digit well code |
| C | תאריך דיגום | Row 3, col C | Formatted as `DD.MM.YYYY` text string |
| D | עומק דיגום | — | Always empty for fuel sites |
| E | מוסד דוגם | Row 5, col C | Sampler company code (int) |
| F | סימן | — | Always empty for fuel sites |
| G | תוצאה סופית | Measurement cell | Numeric value; 0 for below-detection |
| H | סמל פרמטר | Col A → param_table.xlsx | Text symbol (e.g., CA, BENZ, MTBE) |
| I | מעבדה | Row 4, col C | Lab code (int) |

### BH Lookup Table (`Sites_missing_BH_codes.xlsx`)

Excel with columns: שם האתר (A), חברה (B), חברה מייעצת (C), שם קידוח של החברה (D), קוד קידוח (E).

Used to resolve missing well codes by (site_name, well_name) with fuzzy matching (threshold 0.9).

### Historical Data Format

Auto-detected. Supports two layouts:

**Format 1 — Intake format** (headers in row 1):
Same columns as output above (מס ש"ה, סמל פרמטר, תוצאה סופית, etc.)

**Format 2 — Historical template** (headers in row 5):
Columns: זיהוי קידוח, שם קידוח, תאריך מדידה, שם פרמטר, ריכוז, סמן, etc.

Detection: scans rows 1-10 for a row containing ≥3 of: קידוח, פרמטר, ריכוז, ש"ה, תאריך, מדידה.

---

## Validation Rules

### Errors (block row creation)

- Missing or unparseable sampling date
- Missing lab code that cannot be resolved from reference or manual override
- Missing sampler code that cannot be resolved from reference or manual override
- Missing well code: not in form, BH table, memory, and not in interactive mode
- Non-numeric well code that cannot be resolved
- Unparseable measurement value (not a number, not `"<X"`, not empty)
- No wells found in the form

### Warnings (informational, rows still created)

- Well code completed from BH lookup table (with match details if fuzzy)
- Well code completed from memory
- Well code entered interactively
- Lab/sampler code entered manually via UI
- Lab/sampler name fuzzy match (close but not exact)
- Lab/sampler name not found in reference
- Lab code mismatch between form and reference
- Parameter code not in param_table.xlsx (rows skipped)
- **ECFD ≥ 100** — likely unit error (mS/cm vs µS/cm)
- **Historical anomaly** — value differs by ≥2 orders of magnitude (×100 or ÷100) from previous measurement for same well+parameter

### Special Parameters

| Row | Description | Param Symbol | Notes |
|-----|-------------|-------------|-------|
| 10 | עומק עד פני המים | WDEP | Water depth in meters |
| 11 | OIL LAYER THIKNESS | OILTH | Oil layer thickness |

---

## Well Code Resolution (Priority Order)

When a well has no valid 8-digit code in row 8:

1. **BH lookup table** — search `Sites_missing_BH_codes.xlsx` by (site_name, well_name) with fuzzy matching (SequenceMatcher ≥0.9 or substring). Result saved to memory.
2. **Memory lookup** — check `well_codes_memory.csv` for (site_name, well_name) match
3. **Interactive prompt** — if `--interactive` flag / Streamlit UI: ask user for code, saved to memory
4. **Error** — if none of the above: report error, skip well

Memory is keyed by `(site_name, well_name)` where site_name comes from B2.
`well_codes_memory.csv` is tracked in git so memory persists across machines and deployments.

## Lab/Sampler Code Resolution

When lab or sampler code is missing or cannot be resolved from the reference CSV:

1. **Reference CSV** — exact match → use ref code; fuzzy match → warning + use ref code
2. **Manual override** — Streamlit UI prompts user to enter code per file (not saved)
3. **Error** — if neither: report error, skip all rows for that file

---

## Reference Files

### config/lab_codes.csv

```csv
שם מעבדה,קוד מעבדה
אמינולב,6
בקטוכם,8
...
```

Multiple names can map to the same code (e.g., אמינולב and אמינולאב both → 6).

### config/sampler_codes.csv

```csv
שם חברת דיגום,קוד חברת דיגום
אתגר,23
ידע מים,25
...
```

### param_table.xlsx

Excel with columns: זיהוי פרמטר (int), סמל פרמטר (text), תאור מקוצר, יחידת מידה 1, יחידת מידה 2, סווג כימי.

Examples: 1→CA, 2→Mg, 84→BENZ, 212→ECFD, 2009→MTBE

---

## Dependencies

- Python 3.10+
- streamlit ≥ 1.30.0
- openpyxl ≥ 3.1.0
- pandas ≥ 2.0.0
- Standard library: argparse, csv, os, re, sys, datetime, difflib, pathlib, io, tempfile

---

## Testing

No automated test suite. Manual testing with files in `test_files/`:

| File | Wells | Codes | Tests |
|------|-------|-------|-------|
| Copy_of_אשקלון_...xlsx | 5 | All present | Multi-well conversion, below-detection values (0 and "<X"), full parameter set |
| דיווח_09_25.xlsx | 1 | **Missing** — resolved via BH table | BH fuzzy lookup, memory save |
| טופס_דיווח_נס_ציונה_2025.xlsx | 1 | Present | Basic single-well conversion |

**Regression test command:**

```bash
python convert_report_to_intake.py \
  test_files/Copy_of_אשקלון_טופס_דיווח_דצמבר_2020.xlsx \
  test_files/טופס_דיווח_נס_ציונה_2025.xlsx \
  --params param_table.xlsx \
  --output /dev/null --error-report /dev/null
```

Expected: 0 errors, 0 warnings.

**BH lookup test:**

```bash
python convert_report_to_intake.py \
  test_files/דיווח_09_25.xlsx \
  --params param_table.xlsx \
  --output /dev/null --error-report /dev/null
```

Expected: 0 errors, 1 warning (well code completed from BH table).

---

## Code Style

- Python 3.10+, inline type hints
- Functions should be focused, documented with docstrings
- Hebrew strings only in user-facing output (UI labels, error messages, Excel headers)
- No unnecessary comments — code should be self-explanatory
- Error handling: collect errors/warnings in lists, don't raise exceptions during conversion
- All file I/O uses explicit UTF-8 encoding
- Temp files always cleaned up via `try/finally`
- Download buffers use `io.BytesIO` (never write to shared temp dir)
