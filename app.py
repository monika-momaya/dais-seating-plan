import io
from dataclasses import dataclass
from typing import List

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

st.set_page_config(page_title="Seating Plan Generator", layout="wide")


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)


def set_cell_border(cell, color="000000", size="8"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.first_child_found_in("w:tcBorders")
    if tc_borders is None:
        tc_borders = OxmlElement('w:tcBorders')
        tc_pr.append(tc_borders)
    for edge in ('top', 'left', 'bottom', 'right'):
        element = tc_borders.find(qn(f'w:{edge}'))
        if element is None:
            element = OxmlElement(f'w:{edge}')
            tc_borders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), size)
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color)


def style_paragraph(paragraph, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    paragraph.alignment = align
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(size)
        run.bold = bold
        if color:
            run.font.color.rgb = RGBColor.from_string(color)


def set_landscape(doc: Document):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)


def create_document(event_meta: dict, df: pd.DataFrame) -> io.BytesIO:
    doc = Document()
    set_landscape(doc)

    title = doc.add_paragraph()
    run = title.add_run(event_meta['title'])
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 102, 153)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for key in ['subtitle', 'time_text']:
        if event_meta.get(key):
            p = doc.add_paragraph(event_meta[key])
            style_paragraph(p, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color='666666')

    doc.add_paragraph("")

    seat_df = df.sort_values('display_order').reset_index(drop=True)
    seat_count = max(len(seat_df), 1)
    seat_table = doc.add_table(rows=2, cols=seat_count)
    seat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    seat_table.autofit = False

    usable_width = 10.5
    cell_width = usable_width / seat_count

    for i, row in seat_df.iterrows():
        top = seat_table.cell(0, i)
        bottom = seat_table.cell(1, i)
        top.width = Inches(cell_width)
        bottom.width = Inches(cell_width)
        top.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        bottom.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        top.text = str(row['seat_no'])
        bottom.text = str(row['code'])
        style_paragraph(top.paragraphs[0], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_paragraph(bottom.paragraphs[0], bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(top, 'F5EFD6')
        set_cell_shading(bottom, 'F5EFD6')
        set_cell_border(top)
        set_cell_border(bottom)

    doc.add_paragraph("")
    detail_table = doc.add_table(rows=1, cols=3)
    detail_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    detail_table.autofit = False
    hdr = detail_table.rows[0].cells
    headers = ['Sr. No.', 'Code', 'Dignitaries on Dais']
    widths = [0.7, 0.8, 8.9]
    for i, text in enumerate(headers):
        hdr[i].width = Inches(widths[i])
        hdr[i].text = text
        style_paragraph(hdr[i].paragraphs[0], bold=True, size=10)
        set_cell_border(hdr[i])

    detail_df = df.sort_values('serial_no').reset_index(drop=True)
    for _, row in detail_df.iterrows():
        cells = detail_table.add_row().cells
        values = [row['serial_no'], row['code'], f"{row['name']}, {row['designation']}"]
        for i, value in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].text = str(value)
            style_paragraph(cells[i].paragraphs[0], bold=(i < 2), size=10)
            set_cell_border(cells[i])

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def template_df():
    return pd.DataFrame([
        {"serial_no": 1, "seat_no": 1, "display_order": 9, "code": "CM", "name": "Shri Example Person", "designation": "Chief Guest"},
        {"serial_no": 2, "seat_no": 2, "display_order": 10, "code": "NG", "name": "Shri Second Person", "designation": "Guest of Honour"},
        {"serial_no": 3, "seat_no": 3, "display_order": 8, "code": "GS", "name": "Shri Third Person", "designation": "Minister"},
    ])


st.title("Free Seating Plan Generator")
st.caption("Build seating plans and export a landscape .docx file using only free Python libraries.")

with st.sidebar:
    st.header("Event details")
    event_title = st.text_input("Title", "INAUGURAL SEATING PLAN (TENTATIVE)")
    subtitle = st.text_input("Subtitle / Venue", "")
    time_text = st.text_input("Date / Time line", "")
    st.markdown("### Input format")
    st.write("Use CSV/Excel columns: serial_no, seat_no, display_order, code, name, designation")

uploaded = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"])

if uploaded:
    if uploaded.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
else:
    df = template_df()

required_cols = ["serial_no", "seat_no", "display_order", "code", "name", "designation"]
missing = [c for c in required_cols if c not in df.columns]

if missing:
    st.error(f"Missing columns: {', '.join(missing)}")
else:
    st.subheader("Edit dignitary data")
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

    st.subheader("Preview")
    preview_cols = st.columns(len(edited) if len(edited) <= 12 else 6)
    seat_preview = edited.sort_values('display_order').reset_index(drop=True)
    for i, row in seat_preview.head(len(preview_cols)).iterrows():
        with preview_cols[i % len(preview_cols)]:
            st.markdown(f"**{row['seat_no']}**  ")
            st.caption(str(row['code']))

    output = create_document(
        {"title": event_title, "subtitle": subtitle, "time_text": time_text},
        edited
    )

    st.download_button(
        "Download Word document",
        data=output.getvalue(),
        file_name="seating-plan.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    csv_data = edited.to_csv(index=False).encode('utf-8')
    st.download_button("Download sample CSV", data=csv_data, file_name="seating-input.csv", mime="text/csv")
