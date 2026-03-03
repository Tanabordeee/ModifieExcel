import streamlit as st
from openpyxl import load_workbook
from datetime import datetime, timedelta
import random
import tempfile

st.title("🎲 Random DateTime Excel Tool")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    sheet_name = st.text_input("Sheet Name")
    column_name = st.text_input("Column Name")

    start_range = st.text_input("Start (YYYY-MM-DD HH:MM:SS)")
    end_range = st.text_input("End (YYYY-MM-DD HH:MM:SS)")

    if st.button("Random & Download"):

        wb = load_workbook(uploaded_file)

        if sheet_name not in wb.sheetnames:
            st.error("Sheet not found")
        else:
            ws = wb[sheet_name]

            headers = [cell.value for cell in ws[1]]
            if column_name not in headers:
                st.error("Column not found")
            else:
                col_index = headers.index(column_name)

                start_dt = datetime.strptime(start_range, "%Y-%m-%d %H:%M:%S")
                end_dt = datetime.strptime(end_range, "%Y-%m-%d %H:%M:%S")

                rows = []
                for row in ws.iter_rows(min_row=2):
                    cell = row[col_index]
                    rows.append(cell)

                total_seconds = int((end_dt - start_dt).total_seconds())

                random_dates = [
                    start_dt + timedelta(seconds=random.randint(0, total_seconds))
                    for _ in rows
                ]

                random_dates.sort()

                for cell, new_dt in zip(rows, random_dates):
                    original_format = cell.number_format
                    cell.value = new_dt
                    cell.number_format = original_format

                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    wb.save(tmp.name)
                    st.download_button(
                        "Download File",
                        data=open(tmp.name, "rb"),
                        file_name="output.xlsx"
                    )