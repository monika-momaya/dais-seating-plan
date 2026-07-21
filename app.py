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

st.set_page_config(page_title="Dais Seating Plan", layout="wide")


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


def compute_auto_display_order(n):
    """Return mapping protocol_rank -> display_order (1..n, left to right).
    Rank 1 sits right-middle, rank 2 sits left-middle, then alternate outward:
    odd ranks (3,5,7,...) extend the right side, even ranks (4,6,8,...) extend the left side.
    For odd n, rank 1 ends up as the exact center seat.
    """
    left = []
    right = []
    for k in range(1, n + 1):
        if k == 1:
            right.append(k)
        elif k == 2:
            left.append(k)
        elif k % 2 == 1:
            right.append(k)
        else:
            left.append(k)
    left_to_right_positions = list(reversed(left)) + right
    mapping = {rank: pos + 1 for pos, rank in enumerate(left_to_right_positions)}
    return mapping


def compute_two_row_layout(n, per_row=12):
    row1_count = min(n, per_row)
    row2_count = n - row1_count
    row1_ranks = list(range(1, row1_count + 1))
    row2_ranks = list(range(row1_count + 1, n + 1))
    row1_order = compute_auto_display_order(len(row1_ranks))
    row2_order = compute_auto_display_order(len(row2_ranks))
    assignment = {}
    for rank in row1_ranks:
        assignment[rank] = {"row": 1, "seat": row1_order[rank]}
    for rank in row2_ranks:
        local_rank = rank - row1_count
        assignment[rank] = {"row": 2, "seat": row2_order[local_rank]}
    return assignment


def compute_three_table_layout(n, center_size=6):
    center_count = min(n, center_size)
    assignment = {}
    if n <= center_count:
        center_order = compute_auto_display_order(n)
        for rank in range(1, n + 1):
            assignment[rank] = {"table": "Center", "seat": center_order[rank]}
        return assignment

    center_ranks = list(range(1, center_count + 1))
    remaining = list(range(center_count + 1, n + 1))
    left_ranks, right_ranks = [], []
    for i, rank in enumerate(remaining):
        if i % 2 == 0:
            right_ranks.append(rank)
        else:
            left_ranks.append(rank)

    center_order = compute_auto_display_order(len(center_ranks))
    left_order = compute_auto_display_order(len(left_ranks))
    right_order = compute_auto_display_order(len(right_ranks))

    for rank in center_ranks:
        assignment[rank] = {"table": "Center", "seat": center_order[rank]}
    for i, rank in enumerate(left_ranks, start=1):
        assignment[rank] = {"table": "Left", "seat": left_order[i]}
    for i, rank in enumerate(right_ranks, start=1):
        assignment[rank] = {"table": "Right", "seat": right_order[i]}
    return assignment


def apply_layout(df, mode, per_row=12, center_size=6):
    n = len(df)
    df = df.sort_values("serial_no").reset_index(drop=True)
    if mode == "Two Rows":
        assignment = compute_two_row_layout(n, per_row=per_row)
        df["group"] = df["serial_no"].map(lambda r: f"Row {assignment[r]['row']}")
        df["group_seat"] = df["serial_no"].map(lambda r: assignment[r]["seat"])
        df["group_order"] = df["serial_no"].map(lambda r: assignment[r]["row"])
    elif mode == "Three Round Tables":
        assignment = compute_three_table_layout(n, center_size=center_size)
        table_order = {"Left": 0, "Center": 1, "Right": 2}
        df["group"] = df["serial_no"].map(lambda r: assignment[r]["table"])
        df["group_seat"] = df["serial_no"].map(lambda r: assignment[r]["seat"])
        df["group_order"] = df["group"].map(table_order)
    else:
        df["group"] = "Dais"
        df["group_seat"] = df["display_order"]
        df["group_order"] = 0
    return df


HONORIFICS = {
    "shri", "smt", "ms", "mrs", "mr", "dr", "miss", "shrimati",
    "prof", "professor", "hon", "hon'ble", "honble", "kumari", "km",
}


def strip_honorifics(name):
    words = []
    for w in name.split():
        clean = w.strip(".,").lower()
        if clean not in HONORIFICS:
            words.append(w.strip(".,"))
    return words


def base_code(words):
    return "".join(w[0].upper() for w in words[:3] if w)


def disambiguate_codes(names):
    word_lists = [strip_honorifics(n) for n in names]
    codes = [base_code(w) for w in word_lists]
    seen_count = {}
    final_codes = []
    for i, code in enumerate(codes):
        words = word_lists[i]
        if codes.count(code) == 1:
            final_codes.append(code)
            continue
        extra_len = 2
        candidate = code
        while True:
            key = code
            occurrence = seen_count.get(key, 0)
            if occurrence == 0:
                candidate = code
            else:
                last_word = words[-1] if words else ""
                candidate = base_code(words[:-1]) + last_word[:extra_len].upper() if words else code
                if candidate in final_codes:
                    extra_len += 1
                    continue
            if candidate not in final_codes:
                break
            extra_len += 1
        seen_count[key] = seen_count.get(key, 0) + 1
        final_codes.append(candidate)
    return final_codes


def parse_pasted_names(text):
    rows = []
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    names = []
    designations = []
    manual_codes = []
    for line in lines:
        parts = [p.strip() for p in line.split("|")]
        names.append(parts[0] if len(parts) > 0 else "")
        designations.append(parts[1] if len(parts) > 1 else "")
        manual_codes.append(parts[2] if len(parts) > 2 else None)

    auto_codes = disambiguate_codes(names)
    codes = [mc if mc else ac for mc, ac in zip(manual_codes, auto_codes)]

    for i, (name, designation, code) in enumerate(zip(names, designations, codes), start=1):
        rows.append({"serial_no": i, "code": code, "name": name, "designation": designation})

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    order_map = compute_auto_display_order(len(df))
    df["seat_no"] = df["serial_no"]
    df["display_order"] = df["serial_no"].map(order_map)
    return df


def create_document(event_meta, df, layout_mode="Single Row"):
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


    def style_seat_cell(cell, top_text="", bottom_text=""):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.text = ""
        if top_text != "":
            p1 = cell.paragraphs[0]
            p1.text = str(top_text)
            style_paragraph(p1, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
        p2 = cell.add_paragraph(str(bottom_text)) if bottom_text != "" else cell.add_paragraph("")
        style_paragraph(p2, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cell, "F5EFD6")
        set_cell_border(cell, size="6")
        set_cell_margins(cell, 60, 60, 60, 60)

    def render_seat_row(seat_df):
        seat_table = doc.add_table(rows=2, cols=max(len(seat_df), 1))
        seat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        seat_table.autofit = False
        cell_width = 10.5 / max(len(seat_df), 1)
        for i, row in seat_df.reset_index(drop=True).iterrows():
            for r in range(2):
                seat_table.cell(r, i).width = Inches(cell_width)
            style_seat_cell(seat_table.cell(0, i), row["seat_no"], row["code"])
            seat_table.cell(1, i).text = ""
            set_cell_border(seat_table.cell(1, i), size="0")

    def render_round_table_block(title_text, grp_df):
        grp_df = grp_df.sort_values("group_seat").reset_index(drop=True)
        label_p = doc.add_paragraph(title_text)
        style_paragraph(label_p, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")
        outer = doc.add_table(rows=3, cols=3)
        outer.alignment = WD_TABLE_ALIGNMENT.CENTER
        outer.autofit = False
        widths = [3.2, 4.1, 3.2]
        for c in range(3):
            for r in range(3):
                outer.cell(r, c).width = Inches(widths[c])
                set_cell_margins(outer.cell(r, c), 40, 40, 40, 40)
        for r in range(3):
            for c in range(3):
                outer.cell(r, c).text = ""
        center = outer.cell(1, 1)
        center.text = ""
        p1 = center.paragraphs[0]
        p1.text = title_text.replace(" Table", "")
        style_paragraph(p1, bold=True, size=12, align=WD_ALIGN_PARAGRAPH.CENTER)
        p2 = center.add_paragraph(f"{len(grp_df)} Dignitaries")
        style_paragraph(p2, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")
        set_cell_shading(center, "E9E2C7")
        set_cell_border(center, size="10")

        positions = [(0,1), (1,2), (2,1), (1,0), (0,2), (2,2), (2,0), (0,0)]
        extra_rows = []
        for i, row in grp_df.iterrows():
            if i < len(positions):
                rr, cc = positions[i]
                style_seat_cell(outer.cell(rr, cc), row["seat_no"], row["code"])
            else:
                extra_rows.append(row)
        if extra_rows:
            extra_label = doc.add_paragraph("Additional seats")
            style_paragraph(extra_label, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")
            render_seat_row(pd.DataFrame(extra_rows))

    if layout_mode == "Single Row":
        render_seat_row(df.sort_values("display_order"))
    elif layout_mode == "Two Rows":
        for grp in ["Row 1", "Row 2"]:
            grp_df = df[df["group"] == grp].sort_values("group_seat")
            if grp_df.empty:
                continue
            label_p = doc.add_paragraph(grp)
            style_paragraph(label_p, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")
            render_seat_row(grp_df)
    elif layout_mode == "Three Round Tables":
        display_groups = ["Left", "Center", "Right"]
        trio = doc.add_table(rows=1, cols=3)
        trio.style = "Table Grid"
        trio.alignment = WD_TABLE_ALIGNMENT.CENTER
        trio.autofit = False
        trio.allow_autofit = False
        col_widths = [Inches(2.85), Inches(3.10), Inches(2.85)]
        for c in range(3):
            trio.cell(0, c).width = col_widths[c]
            trio.cell(0, c).text = ""
        positions = [(0,0), (0,1), (0,2), (1,0), (1,2), (2,0), (2,1), (2,2)]
        for idx, grp in enumerate(display_groups):
            grp_df = df[df["group"] == grp].sort_values("group_seat").reset_index(drop=True)
            cell = trio.cell(0, idx)
            cell.text = ""
            label = cell.paragraphs[0]
            label.text = f"{grp} Table"
            style_paragraph(label, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")
            inner = cell.add_table(rows=3, cols=3)
            inner.style = "Table Grid"
            inner.alignment = WD_TABLE_ALIGNMENT.CENTER
            inner.autofit = False
            inner.allow_autofit = False
            widths = [0.85, 1.10 if grp == "Center" else 0.88, 0.85]
            for c in range(3):
                for r in range(3):
                    inner.cell(r, c).width = Inches(widths[c])
                    inner.cell(r, c).text = ""
            center = inner.cell(1, 1)
            center.text = "MAIN" if grp == "Center" else grp
            style_paragraph(center.paragraphs[0], bold=True, size=12 if grp == "Center" else 11, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(center, "E9E2C7")
            set_cell_border(center, size="8")
            for j, row in grp_df.iterrows():
                if j >= len(positions):
                    break
                rr, cc = positions[j]
                inner.cell(rr, cc).text = ""
                p1 = inner.cell(rr, cc).paragraphs[0]
                p1.text = str(row["seat_no"])
                style_paragraph(p1, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
                p2 = inner.cell(rr, cc).add_paragraph(str(row["code"]))
                style_paragraph(p2, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
                set_cell_shading(inner.cell(rr, cc), "F5EFD6")
                set_cell_border(inner.cell(rr, cc), size="5")
                set_cell_margins(inner.cell(rr, cc), 18, 18, 18, 18)

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

    for _, row in df.sort_values("serial_no").iterrows():
        cells = detail_table.add_row().cells
        values = [row["serial_no"], row["code"], row["name"]]
        for i, value in enumerate(values):
            cells[i].width = Inches(widths[i])
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cells[i].text = str(value)
            style_paragraph(cells[i].paragraphs[0], bold=(i < 2), size=9, align=WD_ALIGN_PARAGRAPH.CENTER if i < 2 else WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_border(cells[i], size="6")
            set_cell_margins(cells[i], 40, 60, 40, 60)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def round_table_svg(label, seats, size=220, center_r=58, seat_r=82):
    import math
    cx = cy = size / 2
    seat_positions = []
    for i, seat in enumerate(seats):
        angle = -90 + (360 / max(len(seats), 1)) * i
        rad = math.radians(angle)
        x = cx + seat_r * math.cos(rad)
        y = cy + seat_r * math.sin(rad)
        seat_positions.append((x, y, seat))
    svg = [f'<svg viewBox="0 0 {size} {size}" width="100%" height="100%" xmlns="http://www.w3.org/2000/svg">']
    svg.append(f'<circle cx="{cx}" cy="{cy}" r="{center_r}" fill="#ddd9d1" stroke="#777" stroke-width="1.5"/>')
    svg.append(f'<text x="{cx}" y="{cy+7}" text-anchor="middle" font-size="26" font-weight="700" fill="#e26b2c">{label}</text>')
    for x, y, seat in seat_positions:
        svg.append(f'<circle cx="{x}" cy="{y}" r="18" fill="#f5efd5" stroke="#000" stroke-width="1"/>')
        svg.append(f'<text x="{x}" y="{y+5}" text-anchor="middle" font-size="12" font-weight="700" fill="#111">{seat}</text>')
    svg.append('</svg>')
    return ''.join(svg)


def round_table_card_html(grp_name, grp_df):
    seats = [f"{r.seat_no}<br><span style='font-size:11px'>{r.code}</span>" for r in grp_df.itertuples(index=False)]
    size = 250 if grp_name == "Center" else 220
    center_r = 64 if grp_name == "Center" else 56
    seat_r = 92 if grp_name == "Center" else 80
    svg = round_table_svg(grp_name[0], [r.seat_no for r in grp_df.itertuples(index=False)], size=size, center_r=center_r, seat_r=seat_r)
    return f"<div style='display:flex;flex-direction:column;align-items:center;gap:8px'>{svg}</div>"

def sample_text():
    return "Shri Bhupendrabhai Patel | Hon'ble Chief Minister, Government of Gujarat | CM\n" \
           "Shri Nitin Gadkari | Hon'ble Minister of Road Transport and Highways, Government of India | NG\n" \
           "Shri Harsh Sanghavi | Hon'ble Deputy Chief Minister, Government of Gujarat | DCM\n" \
           "Shri Pradeep Batra | Hon'ble Transport Minister, Government of Uttarakhand | PB\n" \
           "Shri Arjun Singh | Hon'ble Minister for Transport and Labour, Government of West Bengal | AS"


st.markdown("<style>.seat-card{border:1px solid #d8d8d8;border-radius:14px;padding:12px 8px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.04);text-align:center;min-height:90px;display:flex;flex-direction:column;justify-content:center;align-items:center}</style>", unsafe_allow_html=True)

st.title("Dais Seating Plan")
st.caption("Paste dignitary names in protocol order. Seating is auto-generated, with manual override available.")

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

st.subheader("Paste dignitary names (in protocol order, one per line)")
st.caption("Format: Name | Designation | Code (Designation and Code are optional)")
pasted = st.text_area("Dignitary list", value=sample_text(), height=200)

df = parse_pasted_names(pasted)
if df.empty:
    st.warning("Paste at least one name to continue.")
    st.stop()

edited = df.copy()

st.subheader("Seating layout")
layout_mode = st.radio(
    "Choose dais layout",
    ["Single Row", "Two Rows", "Three Round Tables"],
    horizontal=True,
    help="Single Row: default left-right seating. Two Rows: automatic split, front row gets priority. Three Round Tables: top dignitaries in Center, rest split Left/Right.",
)

per_row = 12
center_size = 6
if layout_mode == "Two Rows":
    per_row = st.number_input("Max seats in front row", min_value=2, max_value=40, value=12, step=1)
elif layout_mode == "Three Round Tables":
    center_size = st.number_input("Number of dignitaries at Center table", min_value=1, max_value=len(edited), value=min(6, len(edited)), step=1)

edited = apply_layout(edited, layout_mode, per_row=per_row, center_size=center_size)

if layout_mode == "Single Row":
    st.subheader("Reorder seats (optional manual drag and drop)")
    seat_labels = [f"{r.seat_no} · {r.code} · {r.name}" for r in edited.sort_values("display_order").itertuples(index=False)]
    if sort_items is not None:
        reordered = sort_items(seat_labels, direction="horizontal")
        if reordered:
            order_map = {label: i + 1 for i, label in enumerate(reordered)}
            edited["display_order"] = edited.apply(lambda r: order_map.get(f"{r.seat_no} · {r.code} · {r.name}", r.display_order), axis=1)
            st.success("Preview order updated by drag and drop.")
        else:
            st.info("Drag the seat cards to change the preview order.")
    else:
        st.info("Install streamlit-sortables to enable drag and drop reordering.")
        st.caption("Without the component, edit the pasted list and re-paste to change order.")

st.subheader("Word preview")
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

if layout_mode == "Single Row":
    preview_data = edited.sort_values("display_order").reset_index(drop=True)
    seat_cols = st.columns(min(len(preview_data), 12) or 1)
    for i, row in preview_data.iterrows():
        with seat_cols[i % len(seat_cols)]:
            st.markdown(f"<div class='seat-card'><div style='font-size:24px;font-weight:800;color:#666;line-height:1'>{row['seat_no']}</div><div style='font-size:16px;color:#666;margin-top:6px;font-weight:600'>{row['code']}</div></div>", unsafe_allow_html=True)
    st.markdown("### Current seat order (left to right)")
    order_df = preview_data[["seat_no", "code", "name"]].reset_index(drop=True)
    st.dataframe(order_df, use_container_width=True, hide_index=True)
elif layout_mode == "Two Rows":
    for grp_name in ["Row 1", "Row 2"]:
        grp_df = edited[edited["group"] == grp_name].sort_values("group_seat").reset_index(drop=True)
        if grp_df.empty:
            continue
        st.markdown(f"#### {grp_name}")
        seat_cols = st.columns(min(len(grp_df), 12) or 1)
        for i, row in grp_df.iterrows():
            with seat_cols[i % len(seat_cols)]:
                st.markdown(f"<div class='seat-card'><div style='font-size:22px;font-weight:800;color:#666;line-height:1'>{row['seat_no']}</div><div style='font-size:15px;color:#666;margin-top:6px;font-weight:600'>{row['code']}</div></div>", unsafe_allow_html=True)
    st.markdown("### Current seat order by group")
    order_df = edited.sort_values(["group_order", "group_seat"])[["group", "seat_no", "code", "name"]].reset_index(drop=True)
    st.dataframe(order_df, use_container_width=True, hide_index=True)
else:
    display_groups = ["Left", "Center", "Right"]
    table_cols = st.columns([1, 1.18, 1])
    for idx, grp_name in enumerate(display_groups):
        grp_df = edited[edited["group"] == grp_name].sort_values("group_seat").reset_index(drop=True)
        with table_cols[idx]:
            st.markdown(f"#### {grp_name} Table")
            st.markdown(round_table_card_html(grp_name, grp_df), unsafe_allow_html=True)
    st.markdown("### Current seat order by group")
    order_df = edited.sort_values(["group_order", "group_seat"])[["group", "seat_no", "code", "name"]].reset_index(drop=True)
    st.dataframe(order_df, use_container_width=True, hide_index=True)

st.markdown("### Dignitary list (as pasted, protocol order 1 to N)")
protocol_df = edited.sort_values("serial_no")[["serial_no", "code", "name", "designation"]].reset_index(drop=True)
st.dataframe(protocol_df, use_container_width=True, hide_index=True)

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

save_current_history(title, subtitle, time_text, edited)

out = create_document({"title": title, "subtitle": subtitle, "time_text": time_text, "logo_path": logo_path}, edited, layout_mode=layout_mode)

st.download_button("Download Word document", data=out.getvalue(), file_name="seating-plan.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
st.download_button("Download CSV template", data=edited.to_csv(index=False).encode("utf-8"), file_name="seating-plan.csv", mime="text/csv")
