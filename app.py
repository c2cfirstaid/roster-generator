
import re
from datetime import datetime
from io import BytesIO
import pandas as pd
import zipfile
import streamlit as st
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.title("Daily Class Roster Generator")
st.write("Upload your Bookeo export file for tomorrow's classes. We'll generate one roster per class, ready to upload to Connecteam.")

uploaded_file = st.file_uploader("Upload Bookeo Export (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Main')

    course_type_col = df.columns[12]  # Column M
    course_level_col = df.columns[61]  # Column BJ
    start_time_col = "Start"

    def extract_location_and_category(text):
        match = re.match(r"(.+?)\s*\((.+)\)", str(text))
        if match:
            category = match.group(1).strip()
            location = match.group(2).strip()
            if "First Aid & CPR/AED" in category:
                category = "First Aid & CPR/AED"
            return category, location
        return "Unknown", "Unknown"

    df[['Course Category', 'Location']] = df[course_type_col].apply(
        lambda x: pd.Series(extract_location_and_category(x))
    )

    df['Start Time'] = pd.to_datetime(df[start_time_col])
    grouped = df.groupby(['Location', 'Course Category', 'Start Time'])

    # Load static info block (rows 16–36 from sample)
    info_block = pd.read_excel("Mar 27_Markham_Blended_Recert_Patrick.xlsx", sheet_name="Roster", header=None).iloc[15:36].reset_index(drop=True)

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for (location, category, start_time), group in grouped:
            # Prepare student data
            student_data = pd.DataFrame({
                "Pass (P) or Fail (F) or No-show (N)": "",
                "First Name": group["First name (participant)"],
                "Last Name": group["Last name (participant)"],
                "Course Taken": group[course_level_col],
                "QR Code / Paper Test\n(specify which)": "",
                "First Aid Kit\n(HIGHLIGHT)\ndelivered = green\nnot delivered = red": "",
                "Textbook\n(HIGHLIGHT)\ndelivered = green\nnot delivered = red": "",
                "Notes for Instructor": ""
            })

            wb = Workbook()
            ws = wb.active
            ws.title = "Roster"

            # Add student data starting at row 1
            for r_idx, row in enumerate(dataframe_to_rows(student_data, index=False, header=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            # Add empty rows if needed to reach row 21
            current_row = len(student_data) + 1
            if current_row < 21:
                current_row = 21

            # Add info block starting at row 21
            for i, row in info_block.iterrows():
                for j, cell in enumerate(row):
                    ws.cell(row=current_row + i, column=j + 1, value=cell)

            # Name and write the file
            date_str = start_time.strftime("%b %d")
            time_str = start_time.strftime("%I%p").lstrip('0')
            clean_course = re.sub(r'\W+', '', category.replace(" ", ""))
            clean_location = re.sub(r'\W+', '', location)
            file_name = f"{date_str}_{clean_location}_{clean_course}_{time_str}.xlsx"

            temp_buffer = BytesIO()
            wb.save(temp_buffer)
            zipf.writestr(file_name, temp_buffer.getvalue())

    st.success("Rosters generated!")
    st.download_button(
        label="📥 Download All Rosters (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Rosters.zip",
        mime="application/zip"
    )
