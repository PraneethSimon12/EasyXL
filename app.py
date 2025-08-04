import streamlit as st
import pandas as pd
import pyodbc
import xlsxwriter  # For streaming Excel
import os
from datetime import datetime
import win32com.client  # For Outlook automation

# --- Folder Setup ---
SQL_FOLDER = 'sql_files'
OUTPUT_FOLDER = 'output_reports'
os.makedirs(SQL_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- DB Connection Details ---
SERVER = 'USER\\SQLEXPRESS'  # ✅ Replace with your server name
DATABASE = 'OnlineRetailDb'  # ✅ Replace with your database name

# --- Query Execution (pandas) ---
def run_query(query):
    conn = pyodbc.connect(
        f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;'
    )
    df = pd.read_sql(query, conn)
    conn.close()
    return df

# --- Excel Export (pandas) ---
def export_excel(df, report_name):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{report_name}_{timestamp}.xlsx"
    output_path = os.path.join(OUTPUT_FOLDER, file_name)

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- Sheet 1: Data ---
        df.to_excel(writer, index=False, sheet_name='Data')
        worksheet = writer.sheets['Data']

        # Header formatting
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
            worksheet.set_column(i, i, max_len)

        # --- Sheet 2: Summary ---
        summary_ws = workbook.add_worksheet("Summary")

        # Basic info
        summary_ws.write("A1", "Report Name")
        summary_ws.write("B1", report_name)
        summary_ws.write("A2", "Generated On")
        summary_ws.write("B2", timestamp)
        summary_ws.write("A3", "Total Rows")
        summary_ws.write("B3", len(df))
        summary_ws.write("A4", "Total Columns")
        summary_ws.write("B4", len(df.columns))

        row = 6  # Start summary table here

        # --- Numeric Column Stats ---
        numeric_cols = df.select_dtypes(include='number').columns

        if not numeric_cols.empty:
            summary_ws.write(row, 0, "Numeric Summary")
            row += 1
            summary_ws.write_row(row, 0, ["Column", "Total", "Average", "Min", "Max"])
            row += 1
            for col in numeric_cols:
                summary_ws.write(row, 0, col)
                summary_ws.write(row, 1, df[col].sum())
                summary_ws.write(row, 2, df[col].mean())
                summary_ws.write(row, 3, df[col].min())
                summary_ws.write(row, 4, df[col].max())
                row += 1
            row += 2

        # --- Categorical Column Value Counts ---
        cat_cols = df.select_dtypes(include='object').columns

        if not cat_cols.empty:
            summary_ws.write(row, 0, "Categorical Value Counts (Top 5)")
            row += 1
            for col in cat_cols:
                summary_ws.write(row, 0, f"{col}")
                row += 1
                value_counts = df[col].value_counts().head(5)
                summary_ws.write_row(row, 0, ["Value", "Count"])
                row += 1
                for val, count in value_counts.items():
                    summary_ws.write(row, 0, str(val))
                    summary_ws.write(row, 1, int(count))
                    row += 1
                row += 1  # Space before next column

    return output_path
# --- Streaming Query Execution (for large queries) ---
def run_query_streaming_to_excel(query, report_name):
    conn = pyodbc.connect(
        f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;'
    )
    cursor = conn.cursor()
    cursor.execute(query)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{report_name}_{timestamp}_streamed.xlsx"
    output_path = os.path.join(OUTPUT_FOLDER, file_name)

    workbook = xlsxwriter.Workbook(output_path)
    data_ws = workbook.add_worksheet("Data")

    # --- Header ---
    columns = [col[0] for col in cursor.description]
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
    for col_idx, col_name in enumerate(columns):
        data_ws.write(0, col_idx, col_name, header_format)

    data_ws.freeze_panes(1, 0)
    data_ws.autofilter(0, 0, 1, len(columns) - 1)

    # --- Stream rows + collect for summary ---
    row_idx = 1
    all_rows = []

    while True:
        rows = cursor.fetchmany(1000)
        if not rows:
            break
        for row in rows:
            row = list(row)
            if len(row) == len(columns):
                all_rows.append(row)
                for col_idx, value in enumerate(row):
                    data_ws.write(row_idx, col_idx, value)
                row_idx += 1
            else:
                print(f"Skipped row with mismatched length: {row}")

    for i in range(len(columns)):
        data_ws.set_column(i, i, 20)

    cursor.close()
    conn.close()

    # --- Create Summary Sheet ---
    df = pd.DataFrame(all_rows, columns=columns)
    summary_ws = workbook.add_worksheet("Summary")

    summary_ws.write("A1", "Report Name")
    summary_ws.write("B1", report_name)
    summary_ws.write("A2", "Generated On")
    summary_ws.write("B2", timestamp)
    summary_ws.write("A3", "Total Rows")
    summary_ws.write("B3", len(df))
    summary_ws.write("A4", "Total Columns")
    summary_ws.write("B4", len(df.columns))

    row = 6
    numeric_cols = df.select_dtypes(include='number').columns
    if not numeric_cols.empty:
        summary_ws.write(row, 0, "Numeric Summary")
        row += 1
        summary_ws.write_row(row, 0, ["Column", "Total", "Average", "Min", "Max"])
        row += 1
        for col in numeric_cols:
            summary_ws.write(row, 0, col)
            summary_ws.write(row, 1, df[col].sum())
            summary_ws.write(row, 2, df[col].mean())
            summary_ws.write(row, 3, df[col].min())
            summary_ws.write(row, 4, df[col].max())
            row += 1
        row += 2

    cat_cols = df.select_dtypes(include='object').columns
    if not cat_cols.empty:
        summary_ws.write(row, 0, "Categorical Value Counts (Top 5)")
        row += 1
        for col in cat_cols:
            summary_ws.write(row, 0, f"{col}")
            row += 1
            value_counts = df[col].value_counts().head(5)
            summary_ws.write_row(row, 0, ["Value", "Count"])
            row += 1
            for val, count in value_counts.items():
                summary_ws.write(row, 0, str(val))
                summary_ws.write(row, 1, int(count))
                row += 1
            row += 1

    workbook.close()
    return output_path

# --- Outlook Email Draft ---
def create_outlook_email(subject, body, attachment_path):
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(os.path.abspath(attachment_path))
    mail.Display()  # Opens Outlook draft

# --- Streamlit UI ---
st.title("SQL to Excel Report Generator")

uploaded_file = st.file_uploader("Upload your SQL file", type=["sql"])
use_streaming = st.checkbox("Use streaming mode for large queries")

if uploaded_file:
    query_text = uploaded_file.read().decode("utf-8")
    report_name = uploaded_file.name.replace('.sql', '')
    st.code(query_text, language='sql')

    if st.button("Run and Generate Report"):
        try:
            if use_streaming:
                excel_path = run_query_streaming_to_excel(query_text, report_name)
                st.success("Large query streamed and Excel report generated!")
            else:
                df = run_query(query_text)
                st.success("Query executed successfully!")
                st.dataframe(df.head(10))
                excel_path = export_excel(df, report_name)

            with open(excel_path, 'rb') as f:
                st.download_button(
                    label="Download Excel Report",
                    data=f,
                    file_name=os.path.basename(excel_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # --- Email Draft Button ---
            st.markdown("---")
            st.subheader("Email Report")

            if st.button("Create Email Draft in Outlook"):
                subject = f"Automated Report: {report_name}"
                body = f"""Hi,

Please find attached the Excel report: {report_name}.

Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Regards,
Automation Bot"""
                create_outlook_email(subject, body, excel_path)
                st.success("Outlook draft created successfully!")

        except Exception as e:
            st.error(f"Error: {e}")
