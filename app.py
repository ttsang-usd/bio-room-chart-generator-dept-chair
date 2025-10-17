import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from datetime import datetime
import io

# --- Helper Functions to match original formatting ---

def format_time_condensed(time_str):
    """Formats time to H:MM without AM/PM and leading zeros."""
    if pd.isna(time_str) or time_str == '':
        return ''
    try:
        time_obj = datetime.strptime(str(int(float(time_str))).zfill(4), '%H%M')
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

def correct_instructor_name(name):
    """Corrects specific instructor names to match the original doc."""
    if pd.isna(name): return ''
    name = str(name).strip().upper()
    corrections = {
        'NYHOLT DE PRADA': 'PRADA',
        'RECART GONZALEZ': 'RECART-GONZALEZ',
        'FLEMING-DAVIES': 'FLEMING-DAVIES' # Ensures consistent formatting
    }
    return corrections.get(name, name)


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
    if pd.isna(time_str) or time_str == '': return None
    try:
        time_str = str(time_str).strip().zfill(4)
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
        return None

    df_cleaned = df.dropna(subset=['BLDG', 'ROOM', 'BEGIN', 'END'])
    df_cleaned['Course'] = df_cleaned['SUBJ'].astype(str) + df_cleaned['CRSE #'].astype(str)
    df_cleaned['Instructor'] = df_cleaned['LAST NAME'].apply(correct_instructor_name)
    df_cleaned['Days'] = df_cleaned.apply(get_day_of_week, axis=1)

    for _, row in df_cleaned.iterrows():
        try:
            room_name = f"{row['BLDG'].replace('SCST', 'ST')}{int(float(row['ROOM']))}"
        except (ValueError, TypeError):
            continue
        begin_time = parse_time(row['BEGIN'])
        if begin_time is None: continue

        for day in row['Days']:
            if day not in room_schedule: room_schedule[day] = {}
            if room_name not in room_schedule[day]: room_schedule[day][room_name] = []
            
            room_schedule[day][room_name].append({
                'Begin': row['BEGIN'],
                'End': row['END'],
                'Course': row['Course'],
                'Title': row['TITLE'],
                'Instructor': row['Instructor'],
                'BeginMinutes': begin_time,
                'IsMorning': begin_time < 720
            })
            
    for day in room_schedule:
        for room in room_schedule[day]:
            room_schedule[day][room].sort(key=lambda x: x['BeginMinutes'])
    return room_schedule


def create_room_use_chart(room_schedule):
    doc = Document()
    # Set to Landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    # Set margins
    for sec in doc.sections:
        sec.top_margin = Inches(0.5)
        sec.bottom_margin = Inches(0.5)
        sec.left_margin = Inches(0.5)
        sec.right_margin = Inches(0.5)

    # Title
    heading = doc.add_heading('Room Use Chart for the Biology Laboratories', 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in heading.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True

    # Legend
    p_legend = doc.add_paragraph()
    p_legend.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_b = p_legend.add_run('B = Morning  ')
    run_b.font.name = 'Times New Roman'
    run_b.font.size = Pt(10)
    run_b.bold = True
    run_b.font.color.rgb = RGBColor(0, 0, 255)
    run_g = p_legend.add_run('G = Afternoon')
    run_g.font.name = 'Times New Roman'
    run_g.font.size = Pt(10)
    run_g.bold = True
    run_g.font.color.rgb = RGBColor(0, 128, 0)
    
    # Table
    days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    all_rooms = ['ST225', 'ST227', 'ST229', 'ST242', 'ST325', 'ST327', 'ST330', 'ST429']
    
    table = doc.add_table(rows=1, cols=len(all_rooms) + 1)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].width = Inches(1.0)
    
    # Table Header Content
    p_hdr_legend = hdr_cells[0].paragraphs[0]
    p_hdr_legend.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_b_hdr = p_hdr_legend.add_run('B=Morning\nG=Afternoon')
    run_b_hdr.font.name = 'Times New Roman'
    run_b_hdr.font.size = Pt(9)
    run_b_hdr.bold = True

    for i, col_name in enumerate(all_rooms, 1):
        p_hdr = hdr_cells[i].paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_hdr.add_run(col_name)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        hdr_cells[i].width = Inches(1.25)

    # Table Body
    for day in days_of_week:
        row_cells = table.add_row().cells
        p_day = row_cells[0].paragraphs[0]
        p_day.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_day = p_day.add_run(day[:3])
        run_day.font.name = 'Times New Roman'
        run_day.font.size = Pt(10)
        run_day.bold = True
        
        for j, room_name in enumerate(all_rooms, 1):
            val = room_schedule.get(day, {}).get(room_name, [])
            para = row_cells[j].paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if val:
                for idx, v in enumerate(val):
                    if idx > 0: para.add_run("\n")
                    
                    begin = format_time_condensed(v['Begin'])
                    end = format_time_condensed(v['End'])
                    title = abbreviate_title(v['Title'])
                    text = f"{begin}-{end}\n{v['Course']}\n{title}\n{v['Instructor']}"
                    
                    run = para.add_run(text)
                    font = run.font
                    font.name = 'Times New Roman'
                    font.bold = True
                    font.size = Pt(9)
                    font.color.rgb = RGBColor(0, 0, 255) if v['IsMorning'] else RGBColor(0, 128, 0)
    return doc

# --- Streamlit App UI ---
st.set_page_config(page_title="Bio Room Use Chart Generator", layout="wide")
st.title("Bio Room Use Chart Generator for Dept Chair")
st.write("This tool converts a class schedule into a Room Use Chart.")
st.header("Upload Your Schedule File")
uploaded_file = st.file_uploader("", type=['csv', 'xlsx', 'xls'], label_visibility="collapsed")

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
st.markdown("**Step 1: Get the Template**\n- Click the button below to download the required template.")
st.download_button(label="Download Template CSV", data=template_csv, file_name="Template_Schedule.csv", mime="text/csv")
st.markdown("""
**Step 2: Prepare Your Data**
- Open the downloaded template (`Template_Schedule.csv`) in Excel or any spreadsheet software.
- **Crucial:** Copy your class schedule data into the appropriate columns. The column headers **must exactly match** the template.

**Step 3: Upload Your File**
- Save your edited file and upload it using the uploader at the top of the page.

**Step 4: Download Your Chart**
- If successful, a blue **"Download Word Document"** button will appear. Click it to save your chart.
""")

