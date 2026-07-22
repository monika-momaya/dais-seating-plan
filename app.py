from pathlib import Path
p = Path('output/app.py')
text = p.read_text()
start = text.find('    elif layout_mode == "Three Round Tables":')
end = text.find('\n    doc.add_paragraph("")', start)
new_block = '''    elif layout_mode == "Three Round Tables":
        display_groups = ["Left", "Center", "Right"]
        trio = doc.add_table(rows=1, cols=3)
        trio.style = "Table Grid"
        trio.alignment = WD_TABLE_ALIGNMENT.CENTER
        trio.autofit = False
        trio.allow_autofit = False
        trio.columns[0].width = Inches(3.0)
        trio.columns[1].width = Inches(3.3)
        trio.columns[2].width = Inches(3.0)

        templates = {
            "Left": {"n_cols": 3, "top": [8, 6, 10], "bottom": [12, 14], "label_cols": (1, 1)},
            "Center": {"n_cols": 3, "top": [2, 1, 3], "bottom": [4, 5], "label_cols": (1, 1)},
            "Right": {"n_cols": 3, "top": [9, 7, 11], "bottom": [13, 15], "label_cols": (1, 1)},
            6: {"n_cols": 4, "top": [4, 1, 2, 3], "bottom": [5, 6], "label_cols": (1, 2)},
            8: {"n_cols": 4, "top": [4, 1, 2, 3], "bottom": [5, 6, 7, 8], "label_cols": (1, 2)},
        }

        def write_cell(cell, seat):
            if seat is None:
                cell.text = ""
                return
            seat_no, code = seat
            cell.text = ""
            p1 = cell.paragraphs[0]
            p1.text = str(seat_no)
            style_paragraph(p1, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            p2 = cell.add_paragraph(code)
            style_paragraph(p2, bold=True, size=8, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(cell, "F5EFD6")
            set_cell_border(cell, size="5")
            set_cell_margins(cell, 10, 10, 10, 10)

        def lookup_map(rows):
            return {int(r.seat_no): r.code for r in rows}

        for idx, grp in enumerate(display_groups):
            grp_df = df[df["group"] == grp].sort_values("group_seat")
            seat_rows = [(int(r.seat_no), r.code) for r in grp_df.itertuples(index=False)]
            n = len(seat_rows)
            tmpl = templates.get(grp) if n == 5 else templates.get(n, templates[grp])
            lookup = lookup_map(grp_df.itertuples(index=False))

            cell = trio.cell(0, idx)
            cell.text = ""
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            spacer = cell.paragraphs[0]
            spacer.text = ""
            style_paragraph(spacer, size=2 if grp == "Center" else 12)
            label = cell.add_paragraph()
            label.text = f"{grp} Table"
            style_paragraph(label, bold=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666")

            inner = cell.add_table(rows=2, cols=tmpl["n_cols"])
            inner.style = "Table Grid"
            inner.alignment = WD_TABLE_ALIGNMENT.CENTER
            inner.autofit = False
            inner.allow_autofit = False
            widths = [0.82] * tmpl["n_cols"]
            widths[tmpl["n_cols"] // 2] = 1.05 if grp == "Center" else 0.9
            for c in range(tmpl["n_cols"]):
                for r in range(2):
                    inner.cell(r, c).width = Inches(widths[c])
                    inner.cell(r, c).text = ""

            top = tmpl["top"]
            bottom = tmpl["bottom"]
            for seat_no, col in zip(top, range(len(top))):
                write_cell(inner.cell(0, col), (seat_no, lookup.get(seat_no, "")))
            bottom_cols = [0, 2] if tmpl["n_cols"] == 3 else [0, 3]
            for seat_no, col in zip(bottom, bottom_cols):
                write_cell(inner.cell(1, col), (seat_no, lookup.get(seat_no, "")))

            a, b = tmpl["label_cols"]
            label_cell = inner.cell(1, a)
            if a != b:
                label_cell = label_cell.merge(inner.cell(1, b))
            label_cell.text = "MAIN" if grp == "Center" else grp
            style_paragraph(label_cell.paragraphs[0], bold=True, size=12 if grp == "Center" else 11, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(label_cell, "E9E2C7")
            set_cell_border(label_cell, size="8")
            set_cell_margins(label_cell, 10, 10, 10, 10)
'''
text = text[:start] + new_block + text[end:]
p.write_text(text)
Path('output/app_fixed.py').write_text(text)
Path('output/seating-plan-app/app.py').write_text(text)
print('locked A/B sequences')
