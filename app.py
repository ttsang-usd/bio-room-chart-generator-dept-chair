import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from datetime import datetime
import io
import os

# --- Helper Functions to match original formatting ---

def format_time_condensed(time_str):
    """Formats time to H:MM without AM/PM and leading zeros."""
    if pd.isna(time_str) or time_str == '':
        return ''
    try:
        time_obj = datetime.strptime(str(int(float(time_str))).zfill(4), '%H%M')
        # Format as 12-hour, then remove leading zero if it exists (e.g., 08:30 -> 8:30)
        formatted_time = time_obj.strftime('%I:%M')
        if formatted_time.startswith('0'):
            return formatted_time[1:]
        return formatted_time
    except (ValueError, TypeError):
        return ''

def abbreviate_title(title):
    """Shortens course titles to match the original document's format."""
    if pd.isna(title):
        return ''
    title = str(title)
    # A dictionary of replacements to shorten long course titles
    replacements = {
        'Anatomy & Physiology II': 'A & P II',
        'Earth/Life Sci for Educators': 'Life Sci Ed',
        'Biology Capstone Seminar': 'Capstone',
        'Genomes and Evolution Lab': 'Genome Evol L',
        'Genomes and Evolution': 'Genome Evol',
        'Bioenergetics and Systems Lab': 'Bioenergetics L',
        'Bioenergetics and Systems': 'Bioenergetics',
        'Medical Microbiology': 'Med Micro',
        'Science in the Public Domain': 'SCI Pub Dom',
        'Research Methods': 'Res Meth',
        'Ecology of Communities': 'Ecol Comm',
        'Cell Physiology Lab': 'Cell Phys L',
        'Molecular Techniques': 'Molec Tech',
        'Life’s Changes & Challenges': 'Life Change Bio',
        'Comparative Anatomy': 'Comp Anat'
    }
    for old, new in replacements.items():
        title = title.replace(old, new)
    return title

# --- Core Logic from your script (adapted for new formatting) ---

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
        if ':' in time_str:
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
        else:
            time_str = time_str.zfill(4)
            hours = int(time_str[:-2])
            minutes = int(time_str[-2:])
        return hours * 60 + minutes
    except:
        return None

def get_day_of_week(row):
    days = []
    for day_char, day_full in [('M', 'Monday'), ('T', 'Tuesday'), ('W', 'Wednesday'), ('R', 'Thursday'), ('F', 'Friday')]:
        if str(row.get(day_char)).strip() == day_char:
            days.append(day_full)
    return days

def process_schedule_data(df):
    room_schedule = {}
    required_columns = ['BLDG', 'ROOM', 'BEGIN', 'END', 'SUBJ', 'CRSE #', 'TITLE', 'LAST NAME', 'M', 'T', 'W', 'R', 'F']
    
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"The uploaded file is missing the following required columns: {', '.join(missing_cols)}")
        st.info("Please use the provided template and ensure all column headers match exactly.")
        return None

    df_cleaned = df.dropna(subset=['BLDG', 'ROOM', 'BEGIN', 'END'])
    df_cleaned['Course'] = df_cleaned['SUBJ'].astype(str) + df_cleaned['CRSE #'].astype(str)
    df_cleaned['Instructor'] = df_cleaned['LAST NAME'].astype(str).str.upper()
    df_cleaned['Days'] = df_cleaned.apply(get_day_of_week, axis=1)

    for _, row in df_cleaned.iterrows():
        try:
            room_name = f"{row['BLDG'].replace('SCST', 'ST')}{int(float(row['ROOM']))}"
        except (ValueError, TypeError):
            continue

        begin_time = parse_time(row['BEGIN'])
        
        if begin_time is None:
            continue

        for day in row['Days']:
            if day not in room_schedule:
                room_schedule[day] = {}
            if room_name not in room_schedule[day]:
                room_schedule[day][room_name] = []

            is_morning = begin_time < 720
            
            room_schedule[day][room_name].append({
                'Begin': row['BEGIN'],
                'End': row['END'],
                'Course': row['Course'],
                'Title': row['TITLE'],
                'Instructor': row['Instructor'],
                'BeginMinutes': begin_time,
                'IsMorning': is_morning
            })
            
    for day in room_schedule:
        for room in room_schedule[day]:
            room_schedule[day][room].sort(key=lambda x: x['BeginMinutes'])
    return room_schedule


def create_room_use_chart(room_schedule):
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    doc.add_heading('Room Use Chart for the Biology Laboratories', 0).alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p_legend = doc.add_paragraph()
    p_legend.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_b = p_legend.add_run('B = Morning  ')
    run_b.bold = True
    run_b.font.color.rgb = RGBColor(0, 0, 255)
    run_g = p_legend.add_run('G = Afternoon')
    run_g.bold = True
    run_g.font.color.rgb = RGBColor(0, 128, 0)

    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    all_rooms = ['ST225', 'ST227', 'ST229', 'ST242', 'ST325', 'ST327', 'ST330', 'ST429']
    
    table = doc.add_table(rows=1, cols=len(all_rooms) + 1)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].width = Inches(0.5)

    p_hdr_legend = hdr_cells[0].paragraphs[0]
    run_b_hdr = p_hdr_legend.add_run('B=Morning ')
    run_b_hdr.font.color.rgb = RGBColor(0, 0, 255)
    run_g_hdr = p_hdr_legend.add_run('G=Afternoon')
    run_g_hdr.font.color.rgb = RGBColor(0, 128, 0)

    for i, col_name in enumerate(all_rooms, 1):
        hdr_cells[i].text = col_name
        hdr_cells[i].width = Inches(1.25)

    for day in days_of_week:
        row_cells = table.add_row().cells
        row_cells[0].text = day[:3]
        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for j, room_name in enumerate(all_rooms, 1):
            val = room_schedule.get(day, {}).get(room_name, [])
            para = row_cells[j].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if val:
                for idx, v in enumerate(val):
                    if idx > 0:
                        para.add_run("\n")
                    
                    begin = format_time_condensed(v['Begin'])
                    end = format_time_condensed(v['End'])
                    title = abbreviate_title(v['Title'])
                    text = f"{begin}-{end}{v['Course']}{title}{v['Instructor']}"
                    
                    run = para.add_run(text)
                    font = run.font
                    font.name = 'Times New Roman'
                    font.bold = True
                    font.size = Pt(9)
                    
                    if v['IsMorning']:
                        font.color.rgb = RGBColor(0, 0, 255)
                    else:
                        font.color.rgb = RGBColor(0, 128, 0)
    return doc

# --- Streamlit App UI ---

st.set_page_config(page_title="Bio Room Use Chart Generator", layout="wide")

st.title("Bio Room Use Chart Generator for Dept Chair")
st.write("This tool converts a class schedule into a Room Use Chart.")

st.header("Upload Your Schedule File")
uploaded_file = st.file_uploader(
    "Upload your class schedule (CSV or Excel)",
    type=['csv', 'xlsx', 'xls'],
    label_visibility="collapsed"
)

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

st.markdown("---") 

st.header("How to Use This App")

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
- Open the downloaded template (`Template_Schedule.csv`) in Excel or any spreadsheet software.
- **Crucial:** Copy your class schedule data into the appropriate columns. The column headers in your file **must exactly match** the template headers.

**Step 3: Upload Your File**
- Save your edited file as either CSV or Excel.
- Drag and drop or browse to upload your file using the uploader at the top of the page.

**Step 4: Download Your Chart**
- If the file is processed successfully, a blue **"Download Word Document"** button will appear at the top of the page. Click it to get your room use chart.
""")

