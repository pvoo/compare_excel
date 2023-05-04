import pandas as pd
import streamlit as st
import os
import tempfile
import numpy as np

def read_excel_file(file):
    return pd.read_excel(file, engine='openpyxl')

def compare_dataframes(df1, df2):
    diff_locations = (df1 != df2) & ~(df1.isna() & df2.isna())
    diff_locations[df1.columns[:3]] = False  # Don't mark the first three columns as different
    return diff_locations

def generate_report(df1, df2, diff_locations):
    report = df1.copy()
    for col in df1.columns:
        if col not in df1.columns[:3]:  # Skip the first three columns
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

    writer.save()

st.set_page_config(page_title="Excel Files Comparison", page_icon=None, layout="centered", initial_sidebar_state="auto")
st.title("Excel Files Comparison")

file1_name = st.text_input("Enter the name of the first Excel file (e.g., file1.xlsx)", value='ProductsExcel_v1.xlsx')
file2_name = st.text_input("Enter the name of the second Excel file (e.g., file2.xlsx)",value='ProductsExcel_v2.xlsx')

if file1_name and file2_name:
    if os.path.exists(file1_name) and os.path.exists(file2_name):
        st.write("Comparing files...")
        df1 = read_excel_file(file1_name)
        df2 = read_excel_file(file2_name)

        diff_locations = compare_dataframes(df1, df2)
        report, diff_locations = generate_report(df1, df2, diff_locations)

        with tempfile.NamedTemporaryFile(mode="wb", suffix=".xlsx", delete=False) as tmpfile:
            save_report(report, diff_locations, tmpfile.name)

        st.write("Difference report:")
        st.write(report)

        st.markdown("Download report as Excel file:")
        with open(tmpfile.name, "rb") as f:
            download_link = st.download_button("Download Report", data=f.read(), file_name="difference_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        os.unlink(tmpfile.name)
    else:
        st.error("Files not found. Make sure the files are in the correct folder.")
