from bs4 import BeautifulSoup
import pandas as pd
import streamlit as st

def html_processor(html_content):

    soup = BeautifulSoup(html_content, 'html.parser')

    # Step 3: Find all tables and process them
    sections = soup.find_all(string= lambda text: text.startswith("Room"))
    room_data = []
    for section in sections[2:]:
        section_name = section.text
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
                    room_data.append([section_name,mindf,avgdf,maxdf])
    df = pd.DataFrame(room_data, columns=["Room details", "Min Daylight Factor (%)", "Avg Daylight Factor (%)", "Max Daylight Factor (%)"])
    return df

st.title("Daylight Analysis Report Converter")

uploaded_file = st.file_uploader("Upload a FlucsDL HTML Report", type=["htm", "html"])
if uploaded_file:
    html_content = uploaded_file.read().decode('windows-1252')
    df = html_processor(html_content)

    st.write("### Extracted Data")
    st.dataframe(df)

    # Save DataFrame to an Excel file
    excel_file_path = "Room_daylight_factors.xlsx"
    df.to_excel(excel_file_path, index=False)

    with open(excel_file_path, "rb") as file:
        st.markdown("### 🎉 Your file is ready for download!")
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name="Room_daylight_factors.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

