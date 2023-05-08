# Excel Files Comparison

This is a Streamlit web application that compares two Excel files and generates a difference report. The report shows cells with differences between the two files and highlights them in yellow. 

## Dependencies

Make sure to install the following dependencies:

```bash
pip install pandas numpy streamlit openpyxl xlsxwriter
```

## How to Run

To run the Streamlit app, use the following command:

```bash
streamlit run app.py
```

Then, open the provided URL in a web browser to start using the application.

## Usage

1. Upload the first Excel file (e.g., file1.xlsx) using the file uploader.
2. Upload the second Excel file (e.g., file2.xlsx) using the second file uploader.
3. The application will compare the files and display the difference report in the browser.
4. You can download the difference report as an Excel file by clicking the "Download Report" button.

## Code Overview

### Functions

- `read_excel_file(file)`: Reads an Excel file and returns a pandas DataFrame.
- `compare_dataframes(df1, df2)`: Compares two DataFrames and returns a DataFrame of the same shape with `True` at the locations where the data differs and `False` elsewhere.
- `generate_report(df1, df2, diff_locations)`: Generates a difference report DataFrame with the format `cell_value_in_df2 -> cell_value_in_df1` for cells with differences.
- `save_report(report, diff_locations, output_file)`: Saves the difference report to an Excel file with yellow background for cells with differences.

### Streamlit App

The Streamlit app is structured as follows:

1. Set page configurations and display the title.
2. Create file uploaders for both Excel files.
3. If both files are uploaded, read them into DataFrames and compare them.
4. Generate and display the difference report.
5. Create a download button for the difference report as an Excel file.
6. Delete the temporary file used for the download.