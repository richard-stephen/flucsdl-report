import re

from bs4 import BeautifulSoup
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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


def format_worksheet(ws):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    text_align = Alignment(horizontal="left", vertical="center")
    number_align = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = border
            if cell.row > 1:
                cell.alignment = text_align if cell.column <= 2 else number_align

    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for cell in ws[letter]:
            if cell.value is None:
                continue
            longest_line = max(len(line) for line in str(cell.value).splitlines())
            if longest_line > max_len:
                max_len = longest_line
        ws.column_dimensions[letter].width = min(max_len + 2, 40)


st.title("Daylight Analysis Report Converter")

uploaded_file = st.file_uploader("Upload a FlucsDL HTML Report", type=["htm", "html"])
if uploaded_file:
    html_content = uploaded_file.read().decode('windows-1252')
    df = html_processor(html_content)

    st.write("### Extracted Data")
    st.dataframe(df)

    # Save DataFrame to an Excel file
    excel_file_path = "Room_daylight_factors.xlsx"
    with pd.ExcelWriter(excel_file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Daylight Factors")
        format_worksheet(writer.sheets["Daylight Factors"])

    with open(excel_file_path, "rb") as file:
        st.markdown("### 🎉 Your file is ready for download!")
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name="Room_daylight_factors.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

