import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Admissions Excel File Updater")

st.title("Admissions Excel File Updater")
st.write(
    """
    Upload the University CSV file and the existing Excel file.  
    The app will update the Excel sheet, removing applicants with '37POSCPMA', and let you download the new version.
    """
)

csv_file = st.file_uploader("Upload University CSV file", type=["csv"])
excel_file = st.file_uploader("Upload current Excel file", type=["xlsx"])

if csv_file and excel_file:
    try:
        df_univ = pd.read_csv(csv_file)
        df_univ.columns = df_univ.columns.str.strip().str.lower().str.replace(" ", "_")
        date_cols = ['application_date', 'department_review_date']
        for col in date_cols:
            if col in df_univ.columns:
                df_univ[col] = pd.to_datetime(df_univ[col], errors='coerce')
        df_univ = df_univ[df_univ["acad_plan"] != "37POSCPMA"]

        df_excel = pd.read_excel(excel_file, engine="openpyxl")
        df_excel.columns = df_excel.columns.str.strip().str.lower().str.replace(" ", "_")

        if 'id' not in df_univ.columns or 'id' not in df_excel.columns:
            st.error("ID column missing in one of the files. Please check your columns.")
        else:
            merged_df = df_univ.merge(df_excel, on='id', how='left', suffixes=('_new', '_existing'))
            updates = merged_df[merged_df['prog_actn_new'] != merged_df['prog_actn_existing']]
            new_applicants = df_univ[~df_univ['id'].isin(df_excel['id'])]

            for idx, row in updates.iterrows():
                df_excel.loc[df_excel['id'] == row['id'], 'prog_actn'] = row['prog_actn_new']

            df_excel = pd.concat([df_excel, new_applicants], ignore_index=True)

            # Create downloadable Excel file in memory
            output = BytesIO()
            df_excel.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)

            st.success("Excel file successfully updated! Download below:")
            st.download_button(
                label="Download Updated Excel File",
                data=output,
                file_name="Updated_Applicants.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Something went wrong: {e}")

else:
    st.info("Upload both the University CSV file and the Excel file to begin.")
