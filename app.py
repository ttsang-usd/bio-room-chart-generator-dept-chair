import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from datetime import datetime
import io

# --- Core Logic from your script ---
# (Slightly adapted to work with Streamlit's file objects)

def load_schedule_data(uploaded_file):
    """Loads data from an uploaded CSV or Excel file."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        else:
            st.error("Unsupported file format. Please upload a CSV or Excel file.")
            return None
        return df
    except Exception as e:
        st.error(f"Error reading the file: {e}")
        return None

def parse_time(time_str):
    if pd.isna(time_str) or time_str == '':
        return None
    try:
        time_str = str(time_str).strip()
        # Handle formats like "14:30" or "1430"
        if ':' in time_str:
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
        else:
            time_str = time_str.zfill(4) # Pad with leading zero if needed e.g. 930 -> 0930
            hours = int(time_str[:-2])
            minutes = int(time_str[-2:])
        return hours * 60 + minutes
    except:
        return None

def format_time_12hr(time_str):
    if pd.isna(time_str) or time_str == '':
        return ''
    time_obj = datetime.strptime(str(int(float(time_str))).zfill(4), '%H%M')
    return time_obj.strftime('%I:%M %p')

def get_day_of_week(row):
    days = []
    for day_char, day_full in [('M', 'Monday'), ('T', 'Tuesday'), ('W', 'Wednesday'), ('R', 'Thursday'), ('F', 'Friday')]:
        if row.get(day_char) == day_char:
            days.append(day_full)
    return days

def process_schedule_data(df):
    room_schedule = {}
    required_columns = ['BLDG', 'ROOM', 'BEGIN', 'END', 'SUBJ', 'CRSE #', 'TITLE', 'LAST NAME', 'M', 'T', 'W', 'R', 'F']
    
    # Check for missing columns
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"The uploaded file is missing the following required columns: {', '.join(missing_cols)}")
        st.info("Please use the provided template and ensure all column headers match exactly.")
        return None

    df_cleaned = df.dropna(subset=['BLDG', 'ROOM', 'BEGIN', 'END'])
    df_cleaned['Course'] = df_cleaned['SUBJ'].astype(str) + ' ' + df_cleaned['CRSE #'].astype(str)
    df_cleaned['Instructor'] = df_cleaned['LAST NAME'].astype(str)
    df_cleaned['Days'] = df_cleaned.apply(get_day_of_week, axis=1)

    for _, row in df_cleaned.iterrows():
        room_name = f"{row['BLDG']} {row['ROOM']}"
        begin_time = parse_time(row['BEGIN'])
        end_time = parse_time(row['END'])
        
        if begin_time is None or end_time is None:
            continue

        for day in row['Days']:
            if day not in room_schedule:
                room_schedule[day] = {}
            if room_name not in room_schedule[day]:
                room_schedule[day][room_name] = []

            is_morning = begin_time < 720 # 12:00 PM in minutes
            
            room_schedule[day][room_name].append({
                'Begin': format_time_12hr(row['BEGIN']),
                'End': format_time_12hr(row['END']),
                'Course': row['Course'],
                'Title': row['TITLE'],
                'Instructor': row['Instructor'],
                'BeginMinutes': begin_time,
                'IsMorning': is_morning
            })
            
    # Sort entries by time
    for day in room_schedule:
        for room in room_schedule[day]:
            room_schedule[day][room].sort(key=lambda x: x['BeginMinutes'])
    return room_schedule


def create_room_use_chart(room_schedule):
    doc = Document()
    doc.add_heading('Classroom Use Chart', 0)
    
    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    all_rooms = sorted(list(set(room for day_rooms in room_schedule.values() for room in day_rooms.keys())))
    
    df_data = []
    for day in days_of_week:
        row_data = {'Day': day}
        if day in room_schedule:
            for room in all_rooms:
                row_data[room] = room_schedule[day].get(room, [])
        else:
             for room in all_rooms:
                row_data[room] = []
        df_data.append(row_data)
        
    df_chart = pd.DataFrame(df_data)
    
    if df_chart.empty:
        doc.add_paragraph("No schedule data was processed to create a chart.")
        return doc

    cols = ['Day'] + all_rooms
    df_chart = df_chart[cols]

    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(cols):
        hdr_cells[i].text = col_name

    for _, row in df_chart.iterrows():
        row_cells = table.add_row().cells
        for j, col_name in enumerate(cols):
            val = row[col_name]
            para = row_cells[j].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if col_name == 'Day':
                run = para.add_run(val)
                run.font.name = 'Times New Roman'
                run.bold = True
                run.font.size = Pt(20)
            else:
                if isinstance(val, list) and val:
                    morning = [v for v in val if v['IsMorning']]
                    afternoon = [v for v in val if not v['IsMorning']]
                    if len(morning) == 0 and len(afternoon) == 1:
                        row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
                    elif len(afternoon) == 0 and len(morning) == 1:
                        row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.TOP
                    else:
                        row_cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for idx, v in enumerate(morning + afternoon):
                        if idx > 0:
                            para.add_run("\n\n")
                        text = f"{v['Begin']}-{v['End']}\n{v['Course']}\n{v['Title']}\n{v['Instructor']}"
                        run = para.add_run(text)
                        run.font.name = 'Times New Roman'
                        run.bold = True
                        run.font.size = Pt(9)
                        run.font.color.rgb = RGBColor(0, 0, 0)
    return doc

# --- Streamlit App UI ---

st.set_page_config(page_title="Bio Room Use Chart Generator", layout="wide")

st.title("🧪 Room Use Chart Generator for Bio Dept Chair")

st.header("Upload Class Schedule File (.xlsx or .csv")
uploaded_file = st.file_uploader(
    "Upload your class schedule (CSV or Excel)",
    type=['csv', 'xlsx', 'xls'],
    label_visibility="collapsed"
)

# --- File Processing Logic ---
if uploaded_file is not None:
    with st.spinner('Processing your file...'):
        df = load_schedule_data(uploaded_file)
        
        if df is not None:
            room_schedule = process_schedule_data(df)
            
            if room_schedule:
                st.success("File processed successfully! Your document is ready for download.")
                
                doc = create_room_use_chart(room_schedule)
                bio = io.BytesIO()
                doc.save(bio)
                
                st.download_button(
                    label="Download Word Document",
                    data=bio.getvalue(),
                    file_name=f"Room_Use_Chart_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("Could not generate a chart. Please check that your file contains the correct data and column headers.")

# --- Instructions Section ---
st.markdown("---") 

st.header("How to Use This App")

# Template data for download button
template_csv = """SUBJ,CRSE #,SEC #,TITLE,ATTRIBUTE,UNITS,M,T,W,R,F,BEGIN,END,BLDG,ROOM,ENROLLMENT,LAST NAME,FIRST NAME
"""

st.markdown("""
**Step 1: Get the Template**
- Click the button below to download the required template. This ensures your data is in the correct format.
""")

st.download_button(
   label="Download Template CSV",
   data=template_csv,
   file_name="Template_Schedule.csv",
   mime="text/csv",
)

st.markdown("""
**Step 2: Prepare Your Data**
- Open the downloaded template in Excel.
- ❗Copy your class schedule data into the appropriate columns. The data type in each column MUST match the column headers. ❗

**Step 3: Upload Your File**
- Save your edited file as either CSV or Excel.
- Upload your file using the uploader at the top of the page.

**Step 4: Download Your Chart**
- If the file is processed successfully, a blue **"Download Word Document"** button will appear at the top of the page. Click it to get your room use chart.
""")


