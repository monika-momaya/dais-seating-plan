import io
import base64
from pathlib import Path
import pandas as pd
import streamlit as st

try:
    from streamlit_sortables import sort_items
except Exception:
    sort_items = None
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

st.set_page_config(page_title="Seating Plan Generator", layout="wide")


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)


def set_cell_border(cell, color="000000", size="8"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right'):
        el = borders.find(qn(f'w:{edge}'))
        if el is None:
            el = OxmlElement(f'w:{edge}')
            borders.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), size)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)

def set_cell_margins(cell, top=60, start=60, bottom=60, end=60):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for m, v in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tc_mar.find(qn(f'w:{m}'))
        if node is None:
            node = OxmlElement(f'w:{m}')
            tc_mar.append(node)
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right'):
        el = borders.find(qn(f'w:{edge}'))
        if el is None:
            el = OxmlElement(f'w:{edge}')
            borders.append(el)
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), size)
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)


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


def set_landscape(doc):
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.35)
    section.bottom_margin = Inches(0.35)
    section.left_margin = Inches(0.35)
    section.right_margin = Inches(0.35)


def chair_svg_data_uri(fill="#f5efd6", stroke="#555555"):
    svg = """<svg xmlns='http://www.w3.org/2000/svg' width='80' height='80' viewBox='0 0 80 80'>
  <g fill='none' stroke='{stroke}' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'>
    <rect x='24' y='16' width='32' height='20' rx='4' fill='{fill}'/>
    <rect x='20' y='34' width='40' height='20' rx='5' fill='{fill}'/>
    <path d='M26 54v14M54 54v14M24 68h6M50 68h6M18 38v10M62 38v10'/>
  </g>
</svg>""".format(fill=fill, stroke=stroke)
    return 'data:image/svg+xml;base64,' + base64.b64encode(svg.encode()).decode()


def render_preview_html(df, title, subtitle, time_text):
    preview = df.sort_values('display_order').reset_index(drop=True)
    seats = []
    for _, r in preview.iterrows():
        seats.append(f"""
        <div class='seat'>
          <div class='chair'>🪑</div>
          <div class='seatno'>{r['seat_no']}</div>
          <div class='code'>{r['code']}</div>
        </div>
        """)
    head = f"""
    <div class='sheet-head'>
      <div class='sheet-head-left'>{title}</div>
      <div class='sheet-head-right'>
        <div>{subtitle}</div>
        <div>{time_text}</div>
      </div>
    </div>
    """
    return head + "<div class='seatrow'>" + "".join(seats) + "</div>"


def create_document(event_meta, df):
    doc = Document()
    set_landscape(doc)

    top = doc.add_table(rows=1, cols=2)
    top.alignment = WD_TABLE_ALIGNMENT.CENTER
    top.autofit = False
    top.cell(0,0).width = Inches(2.2)
    top.cell(0,1).width = Inches(8.8)
    if event_meta.get('logo_path'):
        p = top.cell(0,0).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run()
        run.add_picture(event_meta['logo_path'], width=Inches(1.3))
    else:
        top.cell(0,0).text = ''
    right = top.cell(0,1)
    title = right.paragraphs[0]
    run = title.add_run(event_meta['title'])
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 102, 153)
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for key in ['subtitle', 'time_text']:
        if event_meta.get(key):
            p = right.add_paragraph(event_meta[key])
            style_paragraph(p, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT, color='666666')

    doc.add_paragraph("")
    seat_df = df.sort_values('display_order').reset_index(drop=True)
    seat_count = max(len(seat_df), 1)
    seat_table = doc.add_table(rows=3, cols=seat_count)
    seat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    seat_table.autofit = False
    usable_width = 10.5
    cell_width = usable_width / seat_count

    for i, row in seat_df.iterrows():
        for r in range(3):
            seat_table.cell(r, i).width = Inches(cell_width)
            seat_table.cell(r, i).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        icon = seat_table.cell(0, i)
        num = seat_table.cell(1, i)
        code = seat_table.cell(2, i)
        icon.text = '🪑'
        num.text = str(row['seat_no'])
        code.text = str(row['code'])
        style_paragraph(icon.paragraphs[0], size=14, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_paragraph(num.paragraphs[0], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_paragraph(code.paragraphs[0], bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(num, 'F5EFD6')
        set_cell_shading(code, 'F5EFD6')
        set_cell_border(icon, size='4')
        set_cell_border(num, size='6')
        set_cell_border(code, size='6')
        set_cell_margins(icon, 40, 40, 40, 40)
        set_cell_margins(num, 40, 40, 40, 40)
        set_cell_margins(code, 40, 40, 40, 40)

    doc.add_paragraph("")
    detail_table = doc.add_table(rows=1, cols=3)
    detail_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    detail_table.autofit = False
    headers = ['Sr. No.', 'Code', 'Dignitaries on Dais']
    widths = [0.7, 0.8, 8.9]
    for i, text in enumerate(headers):
        cell = detail_table.rows[0].cells[i]
        cell.width = Inches(widths[i])
        cell.text = text
        style_paragraph(cell.paragraphs[0], bold=True, size=9)
        set_cell_shading(cell, 'DDDDDD')
        set_cell_border(cell, size='6')
        set_cell_margins(cell, 50, 60, 50, 60)

    for _, row in df.sort_values('serial_no').iterrows():
        cells = detail_table.add_row().cells
        values = [row['serial_no'], row['code'], f"{row['name']}, {row['designation']}"]
        for i, value in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].text = str(value)
            style_paragraph(cells[i].paragraphs[0], bold=(i < 2), size=9)
            set_cell_margins(cells[i], 40, 60, 40, 60)
            set_cell_border(cells[i])

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def sample_df():
    return pd.DataFrame([
        {"serial_no": 1, "seat_no": 1, "display_order": 3, "code": "CM", "name": "Shri Example Person", "designation": "Chief Guest"},
        {"serial_no": 2, "seat_no": 2, "display_order": 1, "code": "NG", "name": "Shri Second Person", "designation": "Guest of Honour"},
        {"serial_no": 3, "seat_no": 3, "display_order": 4, "code": "GS", "name": "Shri Third Person", "designation": "Minister"},
        {"serial_no": 4, "seat_no": 4, "display_order": 2, "code": "DCM", "name": "Shri Fourth Person", "designation": "Deputy Chief Minister"},
        {"serial_no": 5, "seat_no": 5, "display_order": 5, "code": "PB", "name": "Shri Fifth Person", "designation": "Minister"},
    ])


st.markdown("""
<style>
.sheet-head{display:flex;justify-content:space-between;align-items:flex-start;margin:4px 0 10px 0;padding:6px 2px}
.sheet-head-left{font-size:18px;font-weight:700;color:#0a72b8;text-align:left;max-width:58%}
.sheet-head-left:after{content:''}
.sheet-head-right{font-size:13px;color:#666;text-align:right;line-height:1.5}
.seatrow{display:flex;flex-wrap:wrap;gap:10px;padding:12px 8px 8px;border:1px solid #e5e5e5;border-radius:14px;background:#fcfcfc;align-items:flex-end;min-height:130px}
.seat{min-width:88px;text-align:center;border:1px solid #d8d8d8;border-radius:14px;padding:10px 8px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.04)}
.chair{font-size:30px;line-height:1.1;margin-bottom:4px}
.seatno{font-weight:700;font-size:18px;line-height:1.1;font-family:Arial, sans-serif}
.code{font-size:11px;color:#666;margin-top:2px;font-family:Arial, sans-serif}
.helper{font-size:13px;color:#666}\n.seatrow.no-wrap{flex-wrap:nowrap;overflow-x:auto}
</style>
""", unsafe_allow_html=True)

st.title("Free Seating Plan Generator")
st.caption("Live preview + Word export using only free open-source libraries.")

with st.sidebar:
    st.header("Event details")
    title = st.text_input("Title", "INAUGURAL SEATING PLAN (TENTATIVE)")
    subtitle = st.text_input("Subtitle / Venue", "")
    time_text = st.text_input("Date / Time line", "")
    logo = st.file_uploader("Upload logo", type=["png", "jpg", "jpeg", "webp"])
    show_icons = st.toggle("Show chair icons in preview", value=True)
    wrap_preview = st.toggle("Wrap preview rows", value=True)
    if logo is not None:
        st.image(logo, use_container_width=True)
    st.markdown("### Input columns")
    st.code("serial_no, seat_no, display_order, code, name, designation")

uploaded = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"])
if uploaded:
    if uploaded.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded)
    else:
        df = pd.read_excel(uploaded)
else:
    df = sample_df()

required = ["serial_no", "seat_no", "display_order", "code", "name", "designation"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing columns: {', '.join(missing)}")
    st.stop()

st.subheader("Edit data")
edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

st.subheader("Seat ordering")
seat_labels = [f"{r.seat_no} · {r.code} · {r.name}" for r in edited.sort_values('display_order').itertuples(index=False)]
if sort_items is not None:
    reordered = sort_items(seat_labels, direction="horizontal")
    if reordered:
        order_map = {label: i + 1 for i, label in enumerate(reordered)}
        edited = edited.copy()
        edited['display_order'] = edited.apply(lambda r: order_map.get(f"{r.seat_no} · {r.code} · {r.name}", r.display_order), axis=1)
        st.success("Seat order updated by drag and drop.")
    else:
        st.info("Drag items to reorder the dais sequence.")
else:
    st.info("Install streamlit-sortables to enable drag and drop reordering. Using display_order fallback.")
        st.caption("You can drag to reorder if the component is installed; otherwise edit display_order.")

st.subheader("Live preview")
preview_data = edited.sort_values('display_order').reset_index(drop=True)
if show_icons:
    st.markdown(render_preview_html(preview_data), unsafe_allow_html=True)
else:
    cols = st.columns(min(len(preview_data), 12) or 1)
    for i, row in preview_data.iterrows():
        with cols[i % len(cols)]:
            st.markdown(f"**{row['seat_no']}**")
            st.caption(str(row['code']))

st.markdown("### Current seat order")
order_df = preview_data[['seat_no', 'code', 'name']].reset_index(drop=True)
st.dataframe(order_df, use_container_width=True, hide_index=True)

logo_path = None
if logo is not None:
    logo_path = f"output/seating-plan-app/{logo.name}"
    Path(logo_path).write_bytes(logo.getbuffer())

out = create_document({'title': title, 'subtitle': subtitle, 'time_text': time_text, 'logo_path': logo_path}, edited.sort_values('display_order'))

st.download_button(
    'Download Word document',
    data=out.getvalue(),
    file_name='seating-plan.docx',
    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
)

st.download_button(
    'Download CSV template',
    data=edited.to_csv(index=False).encode('utf-8'),
    file_name='seating-plan.csv',
    mime='text/csv'
)
