import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

EXCLUDED_PLAN = "37POSCPMA"

st.set_page_config(page_title="Admissions Excel File Updater")
st.title("Admissions Excel File Updater")
st.write(
    "Upload the University CSV and the current Excel file. "
    "The app will show you exactly what changed before you download."
)

csv_file = st.file_uploader("Upload University CSV file", type=["csv"])
excel_file = st.file_uploader("Upload current Excel file", type=["xlsx"])

if not (csv_file and excel_file):
    st.info("Upload both files to begin.")
    st.stop()

def normalize_cols(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

try:
    df_univ = normalize_cols(pd.read_csv(csv_file))
    df_excel = normalize_cols(pd.read_excel(excel_file, engine="openpyxl"))
except Exception as e:
    st.error(f"Could not read files: {e}")
    st.stop()

# Validate required columns
for label, df in [("University CSV", df_univ), ("Excel file", df_excel)]:
    if "id" not in df.columns:
        st.error(f"'{label}' is missing an `id` column. Check the file and try again.")
        st.stop()

if "prog_actn" not in df_univ.columns:
    st.error("University CSV is missing a `prog_actn` column.")
    st.stop()

if "acad_plan" not in df_univ.columns:
    st.error("University CSV is missing an `acad_plan` column (used to exclude non-MPA records).")
    st.stop()

# Parse date columns
for col in ["application_date", "department_review_date"]:
    if col in df_univ.columns:
        df_univ[col] = pd.to_datetime(df_univ[col], errors="coerce")

# Filter out non-MPA plan
excluded_count = (df_univ["acad_plan"] == EXCLUDED_PLAN).sum()
df_univ = df_univ[df_univ["acad_plan"] != EXCLUDED_PLAN]

# Identify updates (NaN-safe comparison)
merged = df_univ.merge(df_excel[["id", "prog_actn"]], on="id", how="inner", suffixes=("_new", "_existing"))
changed = merged[
    merged["prog_actn_new"].fillna("") != merged["prog_actn_existing"].fillna("")
][["id", "prog_actn_existing", "prog_actn_new"]].rename(
    columns={"prog_actn_existing": "old", "prog_actn_new": "new"}
)

# Identify new applicants (in CSV but not in Excel)
new_ids = df_univ[~df_univ["id"].isin(df_excel["id"])]

# Apply updates
df_result = df_excel.copy()
if not changed.empty:
    for _, row in changed.iterrows():
        df_result.loc[df_result["id"] == row["id"], "prog_actn"] = row["new"]

# Append new applicants — only keep columns already in the Excel
if not new_ids.empty:
    cols_to_add = [c for c in df_excel.columns if c in new_ids.columns]
    df_result = pd.concat([df_result, new_ids[cols_to_add]], ignore_index=True)

# --- Summary ---
st.subheader("Summary")
col1, col2, col3 = st.columns(3)
col1.metric("Records updated", len(changed))
col2.metric("New applicants added", len(new_ids))
col3.metric(f"Excluded ({EXCLUDED_PLAN})", excluded_count)

if not changed.empty:
    st.subheader("Status Changes")
    st.dataframe(changed.reset_index(drop=True), use_container_width=True)

if not new_ids.empty:
    st.subheader("New Applicants")
    display_cols = [c for c in ["id", "last", "first_name", "prog_actn"] if c in new_ids.columns]
    st.dataframe(new_ids[display_cols].reset_index(drop=True), use_container_width=True)

# --- Download ---
today = date.today().strftime("%m%d%Y")
output_name = f"Fall 2026 MPA Applicants Unified {today}.xlsx"

output = BytesIO()
df_result.to_excel(output, index=False, engine="openpyxl")
output.seek(0)

st.divider()
st.download_button(
    label=f"Download {output_name}",
    data=output,
    file_name=output_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
