st.subheader("Seat ordering")
seat_labels = [f"{r.seat_no} · {r.code} · {r.name}" for r in edited.sort_values('display_order').itertuples(index=False)]
if sort_items is not None:
    reordered = sort_items(seat_labels, direction="horizontal")
    if reordered:
        order_map = {label: i + 1 for i, label in enumerate(reordered)}
        edited = edited.copy()
        edited['display_order'] = edited.apply(
            lambda r: order_map.get(f"{r.seat_no} · {r.code} · {r.name}", r.display_order),
            axis=1
        )
        st.success("Seat order updated by drag and drop.")
    else:
        st.info("Drag items to reorder the dais sequence.")
else:
    st.info("Install streamlit-sortables to enable drag and drop reordering. Using display_order fallback.")
    st.caption("You can drag to reorder if the component is installed; otherwise edit display_order.")

st.subheader("Live preview")
