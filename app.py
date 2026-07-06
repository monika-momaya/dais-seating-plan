from pathlib import Path
p = Path('output/seating-plan-app/app.py')
text = p.read_text()
old = '''    if sort_items is not None:
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
'''
new = '''    if sort_items is not None:
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
'''
text = text.replace(old, new)
p.write_text(text)
print('fixed indentation')