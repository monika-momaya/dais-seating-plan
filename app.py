import io
import json
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


def set_paragraph_single_spacing(paragraph):
    fmt = paragraph.paragraph_format
    fmt.line_spacing = 1
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)


def style_paragraph(paragraph, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    paragraph.alignment = align
    set_paragraph_single_spacing(paragraph)
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        run.font.name = "Arial"
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


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_border(cell, color="000000", size="8"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), size)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def set_cell_margins(cell, top=60, start=60, bottom=60, end=60):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in [("top", top), ("start", start), ("bottom", bottom), ("end", end)]:
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def history_path():
    path = Path("output/seating-plan-app/history.json")
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def load_history():
    if "history" not in st.session_state:
        st.session_state["history"] = []
    p = history_path()
    if p.exists() and not st.session_state["history"]:
        try:
            st.session_state["history"] = json.loads(p.read_text())
        except Exception:
            st.session_state["history"] = []


def save_current_history(title, subtitle, time_text, df):
    path = history_path()
    record = {"title": title, "subtitle": subtitle, "time_text": time_text, "rows": df.to_dict(orient="records")}
    hist = st.session_state.get("history", [])
    hist.insert(0, record)
    limit = int(st.session_state.get("history_limit", 5))
    st.session_state["history"] = hist[:limit]
    path.write_text(json.dumps(st.session_state["history"], indent=2, default=str))


def create_document(event_meta, df):
    doc = Document()
    set_landscape(doc)

    top = doc.add_table(rows=1, cols=2)
    top.alignment = WD_TABLE_ALIGNMENT.CENTER
    top.autofit = False
    top.cell(0, 0).width = Inches(2.2)
    top.cell(0, 1).width = Inches(8.8)

    if event_meta.get("logo_path") and Path(event_meta["logo_path"]).exists():
        p = top.cell(0, 0).paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.add_run().add_picture(event_meta["logo_path"], width=Inches(1.3))
    else:
        top.cell(0, 0).text = ""

    right = top.cell(0, 1)
    title = right.paragraphs[0]
    run = title.add_run(event_meta["title"])
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0, 102, 153)
    title.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_paragraph_single_spacing(title)

    for key in ["subtitle", "time_text"]:
        if event_meta.get(key):
            p = right.add_paragraph(event_meta[key])
            style_paragraph(p, size=10, align=WD_ALIGN_PARAGRAPH.RIGHT, color="666666")

    doc.add_paragraph("")
    seat_df = df.sort_values("display_order").reset_index(drop=True)
    seat_table = doc.add_table(rows=3, cols=max(len(seat_df), 1))
    seat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    seat_table.autofit = False
    cell_width = 10.5 / max(len(seat_df), 1)

    for i, row in seat_df.iterrows():
        for r in range(3):
            seat_table.cell(r, i).width = Inches(cell_width)
            seat_table.cell(r, i).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        num = seat_table.cell(1, i)
        code = seat_table.cell(2, i)
        num.text = str(row["seat_no"])
        code.text = str(row["code"])
        style_paragraph(num.paragraphs[0], bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        style_paragraph(code.paragraphs[0], bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(num, "F5EFD6")
        set_cell_shading(code, "F5EFD6")
        set_cell_border(num, size="6")
        set_cell_border(code, size="6")
        set_cell_margins(num, 40, 40, 40, 40)
        set_cell_margins(code, 40, 40, 40, 40)

    doc.add_paragraph("")
    detail_table = doc.add_table(rows=1, cols=3)
    detail_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    detail_table.autofit = False
    headers = ["Sr. No.", "Code", "Inaugural Dignitaries on Dais:"]
    widths = [0.7, 0.8, 8.9]

    for i, hdr in enumerate(headers):
        cell = detail_table.rows[0].cells[i]
        cell.width = Inches(widths[i])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.text = hdr
        style_paragraph(cell.paragraphs[0], bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT)
        set_cell_shading(cell, "DDDDDD")
        set_cell_border(cell, size="6")
        set_cell_margins(cell, 50, 60, 50, 60)

    for _, row in df.iterrows():
        cells = detail_table.add_row().cells
        values = [row["serial_no"], row["code"], row["name"]]
        for i, value in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cells[i].text = str(value)
            style_paragraph(cells[i].paragraphs[0], bold=(i < 2), size=9, align=WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_border(cells[i], size="6")
            set_cell_margins(cells[i], 40, 60, 40, 40)

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


st.markdown("<style>.seat-card{border:1px solid #d8d8d8;border-radius:14px;padding:12px 8px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.04);text-align:center;min-height:90px;display:flex;flex-direction:column;justify-content:center;align-items:center}</style>", unsafe_allow_html=True)

st.title("Free Seating Plan Generator")
st.caption("Word-style preview + export using only free open-source libraries.")

with st.sidebar:
    st.header("Event details")
    title = st.text_input("Title", "INAUGURAL SEATING PLAN (TENTATIVE)")
    subtitle = st.text_input("Subtitle / Venue", "")
    time_text = st.text_input("Date / Time line", "")
    logo = st.file_uploader("Upload logo", type=["png", "jpg", "jpeg", "webp"])
    history_limit = st.selectbox("Store last histories", [5, 10], index=0)
    st.session_state["history_limit"] = history_limit
    if logo is not None:
        st.image(logo, use_container_width=True)

load_history()

uploaded = st.file_uploader("Upload CSV or XLSX", type=["csv", "xlsx"])
if uploaded:
    df = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
else:
    df = sample_df()

required = ["serial_no", "seat_no", "display_order", "code", "name", "designation"]
missing = [c for c in required if c not in df.columns]
if missing:
    st.error(f"Missing columns: {', '.join(missing)}")
    st.stop()

st.subheader("Edit data")
edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)

st.subheader("Reorder seats")
seat_labels = [f"{r.seat_no} · {r.code} · {r.name}" for r in edited.sort_values("display_order").itertuples(index=False)]
if sort_items is not None:
    reordered = sort_items(seat_labels, direction="horizontal")
    if reordered:
        order_map = {label: i + 1 for i, label in enumerate(reordered)}
        edited = edited.copy()
        edited["display_order"] = edited.apply(lambda r: order_map.get(f"{r.seat_no} · {r.code} · {r.name}", r.display_order), axis=1)
        st.success("Preview order updated by drag and drop.")
    else:
        st.info("Drag the seat cards to change the preview order.")
else:
    st.info("Install streamlit-sortables to enable drag and drop reordering.")
    st.caption("Without the component, edit display_order manually.")

st.subheader("Word preview")
preview_data = edited.sort_values("display_order").reset_index(drop=True)
left, right = st.columns([1, 2])
with left:
    if logo is not None:
        st.image(logo, width=160)
with right:
    st.markdown(f"### {title}")
    if subtitle:
        st.write(subtitle)
    if time_text:
        st.write(time_text)

seat_cols = st.columns(min(len(preview_data), 12) or 1)
for i, row in preview_data.iterrows():
    with seat_cols[i % len(seat_cols)]:
        st.markdown(f"<div class='seat-card'><div style='font-size:24px;font-weight:800;color:#666;line-height:1'>{row['seat_no']}</div><div style='font-size:16px;color:#666;margin-top:6px;font-weight:600'>{row['code']}</div></div>", unsafe_allow_html=True)

st.markdown("### Current seat order")
order_df = preview_data[["seat_no", "code", "name"]].reset_index(drop=True)
st.dataframe(order_df, use_container_width=True, hide_index=True)

st.markdown("### Recent histories")
if st.session_state.get("history"):
    for idx, item in enumerate(st.session_state["history"][:st.session_state.get("history_limit", 5)]):
        with st.expander(f"{idx+1}. {item.get('title', '')}", expanded=(idx == 0)):
            st.write(item.get("subtitle", ""))
            st.write(item.get("time_text", ""))
            hist_df = pd.DataFrame(item.get("rows", []))
            if not hist_df.empty:
                st.dataframe(hist_df, use_container_width=True, hide_index=True)
else:
    st.caption("No saved history yet.")

logo_path = None
if logo is not None:
    logo_path = f"output/seating-plan-app/{logo.name}"
    Path(logo_path).write_bytes(logo.getbuffer())

save_current_history(title, subtitle, time_text, edited.sort_values("display_order"))

out = create_document({"title": title, "subtitle": subtitle, "time_text": time_text, "logo_path": logo_path}, edited.sort_values("display_order"))

st.download_button("Download Word document", data=out.getvalue(), file_name="seating-plan.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
st.download_button("Download CSV template", data=edited.to_csv(index=False).encode("utf-8"), file_name="seating-plan.csv", mime="text/csv")
