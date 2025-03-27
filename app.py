
import re
from datetime import datetime
from io import BytesIO
import pandas as pd
import zipfile
import streamlit as st

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

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for (location, category, start_time), group in grouped:
            roster_data = pd.DataFrame({
                "Pass (P) or Fail (F) or No-show (N)": "",
                "First Name": group["First name (participant)"],
                "Last Name": group["Last name (participant)"],
                "Course Taken": group[course_level_col],
                "QR Code / Paper Test\n(specify which)": "",
                "First Aid Kit\n(HIGHLIGHT)\ndelivered = green\nnot delivered = red": "",
                "Textbook\n(HIGHLIGHT)\ndelivered = green\nnot delivered = red": "",
                "Notes for Instructor": ""
            })

            date_str = start_time.strftime("%b %d")
            time_str = start_time.strftime("%I%p").lstrip('0')
            clean_course = re.sub(r'\W+', '', category.replace(" ", ""))
            clean_location = re.sub(r'\W+', '', location)
            file_name = f"{date_str}_{clean_location}_{clean_course}_{time_str}.xlsx"

            excel_buffer = BytesIO()
            roster_data.to_excel(excel_buffer, index=False)
            zipf.writestr(file_name, excel_buffer.getvalue())

    st.success("Rosters generated!")
    st.download_button(
        label="📥 Download All Rosters (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Rosters.zip",
        mime="application/zip"
    )
