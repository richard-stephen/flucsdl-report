import io
import re

from bs4 import BeautifulSoup
import pandas as pd
import streamlit as st

ROOM_HEADING_RE = re.compile(r"^Room\s+(\S+)\s+\((.+)\)\s*$")

def html_processor(html_content):

    soup = BeautifulSoup(html_content, 'html.parser')

    # Step 3: Find all tables and process them
    sections = soup.find_all(string= lambda text: text.startswith("Room"))
    room_data = []
    for section in sections[2:]:
        match = ROOM_HEADING_RE.match(section.text)
        if match:
            room_id = match.group(1).strip()
            room_name = match.group(2).strip()
        else:
            room_id = ""
            room_name = section.text.strip()
        next_element = section.find_next()
        if next_element:
            if next_element.name == 'hr':
                print(f"{section} is skipped since no analysis is performed")
                continue
            paragraph =  next_element.find_next('p')
            if paragraph:
                table = paragraph.find('table')
                if table:
                    rows = table.find_all('tr')
                    third_row = rows[2]
                    cells = third_row.find_all('td')
                    mindf = float(cells[2].text.strip().replace('%',""))
                    avgdf = float(cells[3].text.strip().replace('%',""))
                    maxdf = float(cells[4].text.strip().replace('%',""))            
                    room_data.append([room_id, room_name, mindf, avgdf, maxdf])
    df = pd.DataFrame(room_data, columns=["Room ID", "Room name", "Min Daylight Factor (%)", "Avg Daylight Factor (%)", "Max Daylight Factor (%)"])
    return df


st.title("Daylight Analysis Report Converter")

uploaded_file = st.file_uploader("Upload a FlucsDL HTML Report", type=["htm", "html"])
if uploaded_file:
    html_content = uploaded_file.read().decode('windows-1252')
    df = html_processor(html_content)

    st.write("### Extracted Data")
    st.dataframe(df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Daylight Factors")
    buffer.seek(0)

    st.markdown("### 🎉 Your file is ready for download!")
    st.download_button(
        label="📥 Download Excel File",
        data=buffer,
        file_name="Room_daylight_factors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

