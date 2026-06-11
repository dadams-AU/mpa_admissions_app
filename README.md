# Admissions Excel File Updater

A Streamlit app to keep the MPA applicant tracking spreadsheet in sync with the university admissions system export. Upload your latest university CSV and your master Excel sheet — the app shows you exactly what changed before you download.

**No coding required. No data leaves your computer.**

---

## Features

- Drag-and-drop file upload
- Shows a preview of all status changes and new applicants before download
- Excludes `37POSCPMA` (non-MPA) records automatically
- Fixes only `prog_actn` on existing records — your hand-entered notes are untouched
- Output filename includes today's date (e.g. `Fall 2026 MPA Applicants Unified 06112026.xlsx`)
- Runs locally on Mac or Windows

---

## Getting Started

### 1. Install Prerequisites

Python 3.8+ is required. ([Download Python](https://www.python.org/downloads/))

```bash
pip install -r requirements.txt
```

### 2. Launch the App

```bash
streamlit run admissions_app.py
```

A browser window will open automatically.

**Mac shortcut:** Double-click `Run Admissions App.command` (first run: `chmod +x "Run Admissions App.command"`).

### 3. Upload Your Files

| File | What it is |
|---|---|
| **University CSV** | Latest export from the university admissions system (e.g. `FUL_SR_MQ_GRAD_DEPT_RPT.csv`) |
| **Current Excel file** | Your master MPA applicant tracking sheet (`.xlsx`) |

### 4. Review and Download

The app displays:
- **Records updated** — existing applicants whose `prog_actn` changed
- **New applicants added** — IDs in the CSV not yet in your Excel
- **Excluded** — records filtered out due to `37POSCPMA` acad_plan

Review the change tables, then click **Download** to get the updated file.

---

## Requirements

- Both files must have an `id` column
- The University CSV must have `prog_actn` and `acad_plan` columns
- The Excel file must be `.xlsx` format

---

## File Structure

```
admissions_excell_file_updater/
├── admissions_app.py
├── requirements.txt
├── README.md
├── LICENSE
├── .gitignore
└── Run Admissions App.command   # Mac double-click launcher
```

---

## License

MIT License — see [LICENSE](LICENSE) for details.
