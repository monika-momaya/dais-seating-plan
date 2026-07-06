# Seating Plan Generator

Free Streamlit app for seating-plan previews and Word export.

## Free features
- Live preview
- Chair-style layout
- Landscape DOCX export
- CSV/XLSX input
- Optional logo upload

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Next upgrade
Add a free draggable seat-order component for direct reordering in the UI.

## Draggable ordering
- Recommended free options: streamlit-sortables or streamlit-draggable-list.
- They are open-source and support browser-side reordering for seat order.
- We can wire one into the app state next.

## Drag and drop
- The app now supports free drag-and-drop ordering through streamlit-sortables when installed.
- If the component is unavailable, the app falls back to display_order.

## Preview polish
- Preview can wrap to multiple rows or stay in one row.
- Longer names still export cleanly in DOCX with smaller table font.

## Header and branding
- The exported DOCX now supports a two-column top header block with logo on the left and text on the right.
- The preview mirrors that branding direction for consistency.

## Borders, spacing, typography
- Borders are the thin lines around each box/cell in the Word table.
- Spacing means the padding inside those boxes and the gaps between the seat row and the detail table.
- Typography means font family, size, weight, and alignment.

## Latest package
- Free Streamlit app with live chair preview.
- Optional drag-and-drop ordering via streamlit-sortables.
- Landscape Word export with logo, header, borders, and spacing controls.
- No paid plugin required.
