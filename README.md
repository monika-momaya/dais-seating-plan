# Seating Plan Generator

This is a completely free Streamlit app for generating landscape Word seating plans.

## Free stack
- Streamlit
- pandas
- openpyxl
- python-docx

## Features
- Upload CSV or XLSX
- Edit dignitary list in browser
- Control dais display order separately from serial number
- Export landscape `.docx`
- No paid plugin or paid API required

## Required columns
- `serial_no`
- `seat_no`
- `display_order`
- `code`
- `name`
- `designation`

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes
- This version uses only free open-source libraries.
- You can deploy it free on Streamlit Community Cloud if the app size and usage fit their free limits.
- You can also run it on your own machine at zero software cost.
