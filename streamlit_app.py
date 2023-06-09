import pandas as pd
import streamlit as st
import os
import tempfile
import numpy as np

def read_excel_file(file):
    return pd.read_excel(file, engine='openpyxl')

def compare_dataframes(df1, df2):
    diff_locations = (df1 != df2) & ~(df1.isna() & df2.isna())
    return diff_locations

def generate_report(df1, df2, diff_locations):
    report = df1.copy()
    for col in df1.columns:
        report[col] = report[col].mask(diff_locations[col], df2[col].astype(str) + ' -> ' + df1[col].astype(str))
    report = report.replace({np.nan: None})
    return report, diff_locations

def save_report(report, diff_locations, output_file):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    report.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    format_yellow = workbook.add_format({"bg_color": "#FFFF00"})
    
    for i, col in enumerate(report.columns):
        for row in range(len(report)):
            if diff_locations.at[row, col]:  # Highlight only cells with differences
                worksheet.write(row + 1, i, report.at[row, col], format_yellow)
            else:
                if report.at[row, col] is not None:  # Write only cells with actual values
                    worksheet.write(row + 1, i, report.at[row, col])

    writer.close()

def summary_of_differences(diff_locations):
    total_columns = len(diff_locations.columns)
    columns_with_diff = (diff_locations.any(axis=0)).sum()

    total_rows = len(diff_locations)
    rows_with_diff = (diff_locations.any(axis=1)).sum()

    return total_columns, columns_with_diff, total_rows, rows_with_diff


st.set_page_config(page_title="Excel Files Comparison", page_icon=None, layout="centered", initial_sidebar_state="auto")
st.title("Excel Files Comparison")

#write a description with bullet points in markdown streamlit
st.markdown("""## Usage

1. Upload the first Excel file (e.g., file1.xlsx) using the file uploader.
2. Upload the second Excel file (e.g., file2.xlsx) using the second file uploader.
3. The application will compare the files and display a summary of differences in the browser.
4. You can download the difference report as an Excel file by clicking the "Download Report" button. Cells that are different will be highlighted and show the differences.

## File uploads
""")

uploaded_file1 = st.file_uploader("Choose the first Excel file (e.g., file1.xlsx)", type=['xlsx', 'xls'])
uploaded_file2 = st.file_uploader("Choose the second Excel file (e.g., file2.xlsx)", type=['xlsx', 'xls'])

if uploaded_file1 and uploaded_file2:
    st.write("Comparing files...")
    df1 = pd.read_excel(uploaded_file1)
    df2 = pd.read_excel(uploaded_file2)

    diff_locations = compare_dataframes(df1, df2)
    report, diff_locations = generate_report(df1, df2, diff_locations)

    with tempfile.NamedTemporaryFile(mode="wb", suffix=".xlsx", delete=False) as tmpfile:
        save_report(report, diff_locations, tmpfile.name)

    st.markdown("**Summary of differences:**")


    total_columns, columns_with_diff, total_rows, rows_with_diff = summary_of_differences(diff_locations)

    st.write(f"Total columns: {total_columns}, columns with differences: {columns_with_diff}")
    st.write(f"Total rows: {total_rows}, rows with differences: {rows_with_diff}")


    st.markdown("**Download report as Excel file:**")
    with open(tmpfile.name, "rb") as f:
        download_link = st.download_button("Download Report", data=f.read(), file_name="difference_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    os.unlink(tmpfile.name)
else:
    st.warning("Please upload both Excel files.")