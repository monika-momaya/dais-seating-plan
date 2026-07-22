"""Minute-to-Minute Programme — Streamlit Web App v2
=============================================
New in v2:
  • Logo upload → displayed beside event details
  • File renamed to "M2M Programme for <<Event Name>>"
  • Word doc (.docx) download option
  • Excel download (existing)

Run locally:   streamlit run m2m_app_v2.py
Deploy:        Push to GitHub → share.streamlit.io
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io, base64
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
from pathlib import Path

MAX_HISTORY = 3

def save_generated_files(meta, excel_bytes, word_bytes, programme_items=None):
    """Keep the last MAX_HISTORY generated file-sets in session memory (bytes included)."""
    if "generated_history" not in st.session_state:
        st.session_state.generated_history = []
    entry = dict(meta)
    entry["excel_bytes"] = excel_bytes
    entry["word_bytes"] = word_bytes
    st.session_state.generated_history.insert(0, entry)
    st.session_state.generated_history = st.session_state.generated_history[:MAX_HISTORY]
    st.session_state.last_entered = {
        "event_name": meta.get("event_name",""),
        "event_date": meta.get("event_date",""),
        "venue": meta.get("venue",""),
        "start_time": meta.get("start_time",""),
        "programme_items": programme_items or [],
    }

def render_history(limit=MAX_HISTORY):
    history = st.session_state.get("generated_history", [])[:limit]
    if not history:
        st.caption("No generated history yet in this session. Generate a file below and it will appear here.")
        return
    st.markdown("### 🕘 Recent Generated Versions (last 3, this session)")
    for idx, item in enumerate(history):
        with st.container():
            cols = st.columns([3,1,1])
            with cols[0]:
                st.markdown(
                    f"**{item.get('file_name','Unnamed file')}**  \n"
                    f"{item.get('timestamp','')} · {item.get('event_name','')} · "
                    f"{item.get('event_date','')} · {item.get('venue','')}"
                )
            with cols[1]:
                if item.get("excel_bytes"):
                    st.download_button(
                        "⬇️ Excel",
                        data=item["excel_bytes"],
                        file_name=f"{item.get('file_name','Programme')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"hist_excel_{idx}",
                        use_container_width=True,
                    )
            with cols[2]:
                if item.get("word_bytes"):
                    st.download_button(
                        "⬇️ Word",
                        data=item["word_bytes"],
                        file_name=f"{item.get('file_name','Programme')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"hist_word_{idx}",
                        use_container_width=True,
                    )
        st.markdown("---")

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Minute-to-Minute Programme Planner",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #FAFAFA; }
    .stApp { font-family: 'Arial', sans-serif; }    .section-header {
        background: #C9A84C; color: white; padding: 0.5rem 1rem;
        border-radius: 6px; font-weight: bold; margin: 1rem 0 0.5rem 0; font-size: 0.95rem;
    }
    .time-pill {
        background: #F2DADA; color: #7B1B1B; padding: 0.2rem 0.6rem; border-radius: 20px;
        font-weight: bold; font-size: 0.85rem; font-family: monospace;
    }
    .script-box {
        background: #F8F9FA; border-left: 4px solid #7B1B1B;
        padding: 0.8rem 1rem; border-radius: 0 6px 6px 0; margin: 0.3rem 0;
        font-size: 0.88rem; line-height: 1.6;
    }
    .script-box-kn {
        background: #FFF8EE; border-left: 4px solid #C9A84C;
        padding: 0.8rem 1rem; border-radius: 0 6px 6px 0; margin: 0.3rem 0;
        font-size: 0.88rem; line-height: 1.6;
    }
    .summary-card {
        background: #2C2C2C; color: #C9A84C; padding: 1rem 1.5rem;
        border-radius: 8px; text-align: center; font-weight: bold;
    }
    .event-info-panel {
        background: white; border: 1px solid #F5E6C8; border-radius: 8px;
        padding: 1rem 1.5rem; margin-bottom: 1rem; display: flex; align-items: center; gap: 1.5rem;
    }
    .stDownloadButton button {
        background: #7B1B1B !important; color: white !important; border: none !important;
        border-radius: 6px !important; font-weight: bold !important;
    }
    .stDownloadButton button:hover { background: #A52828 !important; }
    .dl-word button { background: #1F4E79 !important; }
    .dl-word button:hover { background: #2E75B6 !important; }
</style>
""", unsafe_allow_html=True)

# ── Script library ────────────────────────────────────────────────────────────
SCRIPTS = [
    (["welcome address","welcome speech"],
     "We will now have the Welcome Address.\nI request [Name], [Designation], to kindly address the gathering.\n[After speech] We thank [Name] for those inspiring words.",
     "ಈಗ ಸ್ವಾಗತ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n[ಹೆಸರು], [ಹುದ್ದೆ] ಅವರನ್ನು ಸಭೆಯನ್ನು ಉದ್ದೇಶಿಸಿ ಮಾತನಾಡಲು ವಿನಂತಿಸುತ್ತೇವೆ.",
     "अब हम स्वागत भाषण की ओर बढ़ेंगे।\nहम [नाम], [पदनाम] से सभा को संबोधित करने का अनुरोध करते हैं।\n[भाषण के पश्चात] हम [नाम] का इन प्रेरणादायक शब्दों के लिए धन्यवाद करते हैं।"),

    (["keynote","inaugural address","chief guest"],
     "We are privileged to have the Keynote / Inaugural Address.\nI request [Full Name & Designation]\nto kindly deliver the Address.\n[After speech] Please join me in a round of applause.",
     "ಈಗ ಮುಖ್ಯ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n[ಪೂರ್ಣ ಹೆಸರು ಮತ್ತು ಹುದ್ದೆ] ಅವರನ್ನು ಭಾಷಣ ಮಾಡಲು ವಿನಂತಿಸುತ್ತೇವೆ.\n[ಭಾಷಣದ ನಂತರ] ಚಪ್ಪಾಳೆಯೊಂದಿಗೆ ಅಭಿನಂದಿಸೋಣ.",
     "हमें मुख्य / उद्घाटन भाषण का सौभाग्य प्राप्त है।\nहम मुख्य अतिथि, [पूर्ण नाम एवं पदनाम], से\nभाषण देने का अनुरोध करते हैं।\n[भाषण के पश्चात] कृपया तालियों से उनका अभिनंदन करें।"),

    (["welcome","dais","dignitar"],
     "Respected dignitaries, honoured guests, and friends — a very warm welcome to you all.\nWe now request our distinguished guests to kindly proceed to the dais and take their seats.\n[MC reads out names one by one as each dignitary is escorted to the stage]",
     "ಗೌರವಾನ್ವಿತ ಅತಿಥಿಗಳೇ, ಗಣ್ಯ ಮಹನೀಯರೇ ಮತ್ತು ಆತ್ಮೀಯ ಬಂಧುಗಳೇ — ನಿಮಗೆ ಹೃತ್ಪೂರ್ವಕ ಸ್ವಾಗತ.\nನಾವು ಗಣ್ಯ ಅತಿಥಿಗಳನ್ನು ವೇದಿಕೆಯಲ್ಲಿ ಆಸೀನರಾಗಲು ವಿನಂತಿಸುತ್ತೇವೆ.\n[ಪ್ರತಿಯೊಬ್ಬ ಅತಿಥಿಯ ಹೆಸರನ್ನು ಓದಿ ಅವರನ್ನು ಸ್ವಾಗತಿಸಿ]",
     "सम्मानित अतिथिगण, गणमान्य महानुभाव और मित्रगण — आप सभी का हार्दिक स्वागत है।\nहम अपने सम्मानित अतिथियों से मंच पर पधार कर आसन ग्रहण करने का अनुरोध करते हैं।\n[MC प्रत्येक अतिथि का नाम पुकारते हुए उन्हें मंच तक ले जाएं]"),

    (["naada","nadageethe","state anthem"],
     "We shall now commence with the Naada Geethe — the State Anthem of Karnataka.\nI request all those present to please rise.\n[Naada Geethe plays]\nThank you. Please be seated.",
     "ನಾವು ಕರ್ನಾಟಕ ನಾಡಗೀತೆಯೊಂದಿಗೆ ಕಾರ್ಯಕ್ರಮವನ್ನು ಆರಂಭಿಸುತ್ತೇವೆ.\nಎಲ್ಲರೂ ದಯವಿಟ್ಟು ಎದ್ದು ನಿಲ್ಲಬೇಕೆಂದು ವಿನಂತಿಸುತ್ತೇನೆ.\n[ನಾಡಗೀತೆ ಹಾಡಲಾಗುವುದು]\nಧನ್ಯವಾದಗಳು. ದಯವಿಟ್ಟು ಕುಳಿತುಕೊಳ್ಳಿ.",
     "हम अब नाडगीते — कर्नाटक के राज्य गीत के साथ कार्यक्रम आरंभ करेंगे।\nउपस्थित सभी सदस्यों से खड़े होने का अनुरोध है।\n[नाडगीते बजाया जाएगा]\nधन्यवाद। कृपया बैठ जाएं।"),

    (["national anthem","jana gana","jai hind"],
     "We will now conclude with the National Anthem.\nI request everyone to please rise.\n[National Anthem plays]\nJai Hind! Jai Karnataka!\nThank you all. The programme now stands concluded.",
     "ರಾಷ್ಟ್ರಗೀತೆಯೊಂದಿಗೆ ಕಾರ್ಯಕ್ರಮವನ್ನು ಮುಕ್ತಾಯಗೊಳಿಸುತ್ತೇವೆ.\nಎಲ್ಲರೂ ದಯವಿಟ್ಟು ಎದ್ದು ನಿಲ್ಲಬೇಕೆಂದು ವಿನಂತಿಸುತ್ತೇನೆ.\n[ರಾಷ್ಟ್ರಗೀತೆ]\nಜೈ ಹಿಂದ್! ಜೈ ಕರ್ನಾಟಕ!",
     "अब हम राष्ट्रगान के साथ कार्यक्रम का समापन करेंगे।\nसभी से खड़े होने का अनुरोध है।\n[राष्ट्रगान]\nजय हिंद!\nआप सभी का धन्यवाद। कार्यक्रम का समापन होता है।"),

    (["lamp","lighting","deepa","inauguration of"],
     "We will now proceed to the auspicious lighting of the lamp.\nI request [Chief Guest Name & Designation] and the distinguished guests\nto kindly come forward for the lamp lighting ceremony.",
     "ಈಗ ದೀಪ ಪ್ರಜ್ವಲನ ಕಾರ್ಯಕ್ರಮ ನಡೆಯಲಿದೆ.\n[ಮುಖ್ಯ ಅತಿಥಿ ಹೆಸರು] ಮತ್ತು ಗಣ್ಯರನ್ನು ದೀಪ ಬೆಳಗಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ.",
     "अब हम दीप प्रज्वलन के शुभ कार्यक्रम की ओर बढ़ेंगे।\nहम [मुख्य अतिथि का नाम एवं पद] तथा गणमान्य अतिथियों से\nदीप प्रज्वलन हेतु मंच पर पधारने का अनुरोध करते हैं।"),

    (["perspective","industry","context setting","biotech","startup"],
     "We will now hear the perspective from [Name], [Designation].\n[After speech] Thank you, [Name], for those valuable insights.",
     "ಈಗ [ಹೆಸರು], [ಹುದ್ದೆ] ಅವರಿಂದ ದೃಷ್ಟಿಕೋನ ಕೇಳಲಿದ್ದೇವೆ.\n[ಭಾಷಣದ ನಂತರ] ಧನ್ಯವಾದಗಳು, [ಹೆಸರು] ಅವರಿಗೆ.",
     "अब हम [नाम], [पदनाम] से उनके विचार सुनेंगे।\n[भाषण के पश्चात] [नाम] का इन मूल्यवान विचारों हेतु धन्यवाद।"),

    (["introduction","recorded message","ambassador","h.e."],
     "We are privileged to have an introduction by [Name], [Designation],\nfollowed by a recorded message from [Speaker Name].",
     "ಈಗ [ಹೆಸರು] ಅವರಿಂದ ಪರಿಚಯ ಮತ್ತು [ಸಂದೇಶ ನೀಡಿದ ಹೆಸರು] ಅವರ ಸಂದೇಶ ಪ್ರಸ್ತುತಿಯಾಗಲಿದೆ.",
     "हमें [नाम], [पदनाम] द्वारा परिचय का सौभाग्य प्राप्त है,\nइसके पश्चात [वक्ता का नाम] का रिकॉर्डेड संदेश प्रस्तुत किया जाएगा."),

    (["release","policy","souvenir","publication","launch"],
     "We will now proceed to the release of [Name of Publication / Policy].\nI request [Chief Guest / Name] and the distinguished guests to come forward.",
     "ಈಗ [ಪ್ರಕಾಶನ / ನೀತಿ ಹೆಸರು] ಬಿಡುಗಡೆ ಮಾಡಲಾಗುವುದು.\n[ಮುಖ್ಯ ಅತಿಥಿ] ಮತ್ತು ಗಣ್ಯರನ್ನು ಮುಂದೆ ಬರಲು ಕೋರುತ್ತೇವೆ.",
     "अब हम [प्रकाशन / नीति का नाम] के विमोचन की ओर बढ़ेंगे।\nहम [मुख्य अतिथि / नाम] तथा गणमान्य अतिथियों से मंच पर पधारने का अनुरोध करते हैं।"),

    (["felicitat","honour","award","memento"],
     "We will now proceed to the felicitation of our distinguished guests.\nI request [Presenter], [Designation], to kindly felicitate [Guest Name].",
     "ಈಗ ಗಣ್ಯ ಅತಿಥಿಗಳ ಸನ್ಮಾನ ಕಾರ್ಯಕ್ರಮ ನಡೆಯಲಿದೆ.\n[ಹೆಸರು] ಅವರನ್ನು [ಸನ್ಮಾನಿಸಲ್ಪಡುವ ಹೆಸರು] ಅವರನ್ನು ಸನ್ಮಾನಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ.",
     "अब हम अपने गणमान्य अतिथियों के सम्मान समारोह की ओर बढ़ेंगे।\nहम [प्रस्तुतकर्ता], [पदनाम] से [अतिथि का नाम] को सम्मानित करने का अनुरोध करते हैं।"),

    (["cultural","dance","music","performance","song"],
     "We will now be treated to a cultural performance by [Name / Group].\nPlease enjoy.",
     "ಈಗ [ಕಲಾವಿದ / ತಂಡ] ಅವರಿಂದ ಸಾಂಸ್ಕೃತಿಕ ಕಾರ್ಯಕ್ರಮ ಪ್ರಸ್ತುತಿಯಾಗಲಿದೆ.\nದಯವಿಟ್ಟು ಆನಂದಿಸಿ.",
     "अब हम [नाम / समूह] द्वारा एक सांस्कृतिक प्रस्तुति का आनंद लेंगे।\nकृपया आनंद उठाएं।"),

    (["vote of thanks","vote of thank"],
     "We will now have the Vote of Thanks.\nI request [Name], [Designation], to kindly propose the Vote of Thanks.",
     "ಈಗ ವಂದನಾರ್ಪಣೆ ನಡೆಯಲಿದೆ.\n[ಹೆಸರು], [ಹುದ್ದೆ] ಅವರನ್ನು ವಂದನಾರ್ಪಣೆ ಸಲ್ಲಿಸಲು ವಿನಂತಿಸುತ್ತೇವೆ.",
     "अब हम धन्यवाद प्रस्ताव की ओर बढ़ेंगे।\nहम [नाम], [पदनाम] से धन्यवाद प्रस्ताव प्रस्तुत करने का अनुरोध करते हैं।"),

    (["tea","coffee","lunch","break","networking","refreshment"],
     "We will now take a short break. The next session commences at [Time].\nRefreshments are available at [Location].",
     "ನಾವು ಈಗ ವಿರಾಮ ತೆಗೆದುಕೊಳ್ಳಲಿದ್ದೇವೆ.\nಮುಂದಿನ ಅಧಿವೇಶನ [ಸಮಯ]ಕ್ಕೆ ಪ್ರಾರಂಭವಾಗಲಿದೆ.",
     "अब हम एक संक्षिप्त विराम लेंगे। अगला सत्र [समय] पर आरंभ होगा।\nजलपान [स्थान] पर उपलब्ध है।"),

    (["panel","discussion","roundtable","session"],
     "We will now move to the Panel Discussion.\nModerator: [Name, Designation]. I request the panellists to take their seats.",
     "ಈಗ ಸಮಿತಿ ಚರ್ಚೆ ನಡೆಯಲಿದೆ.\nನಿರ್ವಾಹಕರು: [ಹೆಸರು, ಹುದ್ದೆ].",
     "अब हम पैनल चर्चा की ओर बढ़ेंगे।\nसंचालक: [नाम, पदनाम]। हम पैनलिस्ट्स से अपने आसन ग्रहण करने का अनुरोध करते हैं।"),

    (["address by","address from","address"],
     "We will now have an address by [Name], [Designation].\nI request [Name] to kindly come forward.\n[After speech] Thank you, [Name].",
     "ಈಗ [ಹೆಸರು], [ಹುದ್ದೆ] ಅವರಿಂದ ಭಾಷಣ ನಡೆಯಲಿದೆ.\n[ಹೆಸರು] ಅವರನ್ನು ಮುಂದೆ ಬರಲು ವಿನಂತಿಸುತ್ತೇವೆ.",
     "अब [नाम], [पदनाम] द्वारा संबोधन होगा।\nहम [नाम] से मंच पर पधारने का अनुरोध करते हैं।\n[भाषण के पश्चात] धन्यवाद, [नाम]।"),
]

LANGUAGE_OPTIONS = {
    "English + ಕನ್ನಡ (Kannada)": {"code": "kan", "label": "ಕನ್ನಡ", "font": "Nirmala UI"},
    "English + हिन्दी (Hindi)": {"code": "hin", "label": "हिन्दी", "font": "Nirmala UI"},
}

import re as _re

def extract_name_designation(item):
    if not item:
        return None, None
    matches = list(_re.finditer(r'\bby\b', item, flags=_re.IGNORECASE))
    if not matches:
        return None, None
    last_by = matches[-1]
    rest = item[last_by.end():].strip()
    if not rest:
        return None, None
    rest = rest.rstrip('.').strip()
    if ',' in rest:
        name_part, designation = rest.split(',', 1)
        name_part = name_part.strip()
        designation = designation.strip()
    else:
        name_part = rest.strip()
        designation = ""
    if not name_part:
        return None, None
    return name_part, designation

def fill_placeholders(template, name, designation):
    if not name:
        return template
    full = f"{name}, {designation}" if designation else name
    replacements = [
        (r'\[Full Name & Designation\]', full),
        (r'\[Name & Designation\]', full),
        (r'\[Name, Designation\]', full),
        (r'\[Name / Designation\]', full),
        (r'\[Name\]\s*,\s*\[Designation\]', full),
        (r'\[Name\],\s*\[Designation\]', full),
        (r'\[Presenter Name\]', name),
        (r'\[Presenter\]', name),
        (r'\[Guest Name\]', name),
        (r'\[Chief Guest Name & Designation\]', full),
        (r'\[Chief Guest / Name\]', name),
        (r'\[Name\]', name),
        (r'\[Designation\]', designation if designation else ""),
        (r'\[ಹೆಸರು\],?\s*\[ಹುದ್ದೆ\]', full),
        (r'\[ಪೂರ್ಣ ಹೆಸರು ಮತ್ತು ಹುದ್ದೆ\]', full),
        (r'\[ಹೆಸರು\]', name),
        (r'\[ಹುದ್ದೆ\]', designation if designation else ""),
        (r'\[नाम\],?\s*\[पदनाम\]', full),
        (r'\[पूर्ण नाम एवं पदनाम\]', full),
        (r'\[नाम\s*/\s*पदनाम\]', full),
        (r'\[नाम\]', name),
        (r'\[पदनाम\]', designation if designation else ""),
    ]
    result = template
    for pattern, value in replacements:
        result = _re.sub(pattern, value, result)
    result = _re.sub(r',\s*,', ',', result)
    result = _re.sub(r',\s*\n', '\n', result)
    return result

def get_script(item, lang_code="kan"):
    t = item.lower()
    name, designation = extract_name_designation(item)
    for entry in SCRIPTS:
        keywords, eng, kan, hin = entry
        if any(k in t for k in keywords):
            regional = hin if lang_code == "hin" else kan
            eng_filled = fill_placeholders(eng, name, designation)
            regional_filled = fill_placeholders(regional, name, designation) if name else regional
            return eng_filled, regional_filled
    eng = f"We will now have — {item}.\nI request [Name / Designation] to kindly come forward.\n[MC Note: Add specific announcement text]"
    kan = f"ಈಗ — {item}.\n[ಹೆಸರು] ಅವರನ್ನು ಮುಂದೆ ಬರಲು ವಿನಂತಿಸುತ್ತೇವೆ."
    hin = f"अब — {item}.\n[नाम / पदनाम] से मंच पर पधारने का अनुरोध है।"
    eng_filled = fill_placeholders(eng, name, designation)
    regional = hin if lang_code == "hin" else kan
    regional_filled = fill_placeholders(regional, name, designation) if name else regional
    return eng_filled, regional_filled

def is_address(item):
    return any(k in item.lower() for k in ["address","keynote","remarks","speech","perspective","introduction"])

def parse_time(t):
    for fmt in ["%I:%M %p","%I:%M%p","%H:%M","%I %p"]:
        try: return datetime.strptime(t.strip().upper(), fmt)
        except: pass
    return None

def fmt_time(dt): return dt.strftime("%I:%M %p").lstrip("0")
def fmt_slot(s,e): return f"{fmt_time(s)} – {fmt_time(e)}"

def safe_filename(name):
    import re
    return re.sub(r'[\\\\/*?:"<>|]','', name).strip() or "Programme"

def normalize_item_text(text):
    """Convert pasted programme text into a clean single line.
    - Real line breaks become commas.
    - If pasted text loses line breaks, insert a comma before common designation words
      when they are stuck to the previous word.
    """
    if text is None:
        return ""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")
    parts = [p.strip(" ,\t") for p in text.split("\n") if p.strip()]
    text = ", ".join(parts)

    # If line breaks are lost during paste, restore likely separators before common titles/designations.
    stuck_words = [
        "Chairperson", "Chairman", "Secretary", "Minister", "Hon'ble", "Hon'ble",
        "Managing Director", "Director", "Principal Secretary", "Additional Chief Secretary",
        "Deputy Chief Minister", "Chief Minister", "Vice Chancellor", "Commissioner",
        "Ambassador", "President", "Founder", "Co-Founder", "CEO", "CTO", "COO"
    ]
    for word in stuck_words:
        pattern = r'(?<![,\s])(' + re.escape(word) + r')'
        text = re.sub(pattern, r', \1', text)

    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r",\s*,+", ", ", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",\s*(,\s*)+", ", ", text)
    return text.strip(" ,")

def compute_slots(rows, start_dt):
    current = start_dt
    for row in rows:
        end = current + timedelta(minutes=int(row.get("duration", 0) or 0))
        row["start_str"] = fmt_time(current)
        row["slot"] = fmt_slot(current, end)
        current = end
    return rows

def build_word(event_name, event_date, venue, rows, logo_bytes=None, lang_code="kan"):
    try:
        from docx import Document as DocxDoc
        from docx.shared import Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
        from docx.enum.table import WD_ALIGN_VERTICAL
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError as e:
        return None, f"python-docx not available: {e}"

    try:
        doc = DocxDoc()
        normal_style = doc.styles["Normal"]
        normal_style.font.name = "Calibri"
        normal_style.font.size = Pt(11)
        rpr = normal_style.element.get_or_add_rPr()
        rFonts = rpr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rpr.append(rFonts)
        rFonts.set(qn("w:ascii"), "Calibri")
        rFonts.set(qn("w:hAnsi"), "Calibri")
        rFonts.set(qn("w:eastAsia"), "Calibri")

        for section in doc.sections:
            section.top_margin = Cm(1.6)
            section.bottom_margin = Cm(1.6)
            section.left_margin = Cm(1.8)
            section.right_margin = Cm(1.8)

        def rgb(hex_str):
            h = hex_str.lstrip("#")
            return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

        def set_cell_bg(cell, hex_color):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), hex_color)
            tcPr.append(shd)

        def no_borders(cell):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement("w:tcBorders")
            for side in ["top", "left", "bottom", "right"]:
                el = OxmlElement(f"w:{side}")
                el.set(qn("w:val"), "nil")
                tcBorders.append(el)
            tcPr.append(tcBorders)

        if logo_bytes:
            hdr_tbl = doc.add_table(rows=1, cols=2)
            logo_cell, title_cell = hdr_tbl.rows[0].cells
            logo_cell.width = Cm(5.0)
            title_cell.width = Cm(11.6)
            no_borders(logo_cell)
            no_borders(title_cell)
            lp = logo_cell.paragraphs[0]
            lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            lp.add_run().add_picture(io.BytesIO(logo_bytes), width=Cm(4.5))
            tp = title_cell.paragraphs[0]
            tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            tp.paragraph_format.line_spacing = Pt(24)
            tp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            tp.paragraph_format.space_before = Pt(0)
            tp.paragraph_format.space_after = Pt(4)
            tr = tp.add_run(f"Minute-to-Minute Programme for {event_name or 'Event'}")
            tr.bold = True
            tr.font.size = Pt(20)
            tr.font.color.rgb = rgb("1A2B4C")
            tr.font.name = "Calibri"
        else:
            title_para = doc.add_paragraph()
            title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            title_para.paragraph_format.line_spacing = Pt(24)
            title_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            title_para.paragraph_format.space_before = Pt(0)
            title_para.paragraph_format.space_after = Pt(4)
            tr = title_para.add_run(f"Minute-to-Minute Programme for {event_name or 'Event'}")
            tr.bold = True
            tr.font.size = Pt(20)
            tr.font.color.rgb = rgb("1A2B4C")
            tr.font.name = "Calibri"

        details = []
        if event_date:
            details.append(f"Date: {event_date}")
        if venue:
            details.append(f"Venue: {venue}")
        if rows:
            details.append(f"Start Time: {rows[0]['start_str']}")
        if details:
            det_para = doc.add_paragraph(" | ".join(details))
            det_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            det_para.paragraph_format.space_after = Pt(6)
            for run in det_para.runs:
                run.font.size = Pt(10)
                run.font.color.rgb = rgb("555555")
                run.font.name = "Calibri"

        table = doc.add_table(rows=1, cols=3)
        table.autofit = False
        col_widths = [Cm(3.8), Cm(0.4), Cm(13.2)]
        hdr = table.rows[0]
        for j, (cell, text) in enumerate(zip(hdr.cells, ["Timings", "", "Programme Details"])):
            cell.width = col_widths[j]
            set_cell_bg(cell, "F5F5F5")
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j != 2 else WD_ALIGN_PARAGRAPH.LEFT
            if text:
                run = p.add_run(text)
                run.bold = True
                run.font.size = Pt(12)
                run.font.color.rgb = rgb("2C2C2C")
                run.font.name = "Calibri"

        import re as _re_word

        def _bold_name_spans(item_text):
            """Find name spans (after 'by ' or with titles like Dr./Shri/Smt.) to bold in the item text."""
            spans = []
            by_matches = list(_re_word.finditer(r'\bby\b', item_text, flags=_re_word.IGNORECASE))
            if by_matches:
                start = by_matches[-1].end()
                rest = item_text[start:]
                m = _re_word.match(r'\s*([^,]+)', rest)
                if m:
                    name_start = start + m.start(1)
                    name_end = start + m.end(1)
                    spans.append((name_start, name_end))
            title_pattern = _re_word.compile(
                r"(?:Shri|Smt\.?|Dr\.?|Mr\.?|Ms\.?|Mrs\.?|Prof\.?|H\.E\.?|Hon['']?ble|Sri)\s+[A-Z][A-Za-z.]*(?:\s+[A-Z][A-Za-z.]*)*"
            )
            for m in title_pattern.finditer(item_text):
                spans.append((m.start(), m.end()))
            spans.sort()
            merged = []
            for s, e in spans:
                if merged and s <= merged[-1][1]:
                    merged[-1] = (merged[-1][0], max(merged[-1][1], e))
                else:
                    merged.append((s, e))
            return merged

        for row in rows:
            tr_row = table.add_row()
            values = [row["slot"], ":", row["item"]]
            for j, (cell, text) in enumerate(zip(tr_row.cells, values)):
                cell.width = col_widths[j]
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                set_cell_bg(cell, "FFFFFF")
                p = cell.paragraphs[0]
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.0

                if j == 2:
                    spans = _bold_name_spans(text)
                    if spans:
                        pos = 0
                        for s, e in spans:
                            if s > pos:
                                r_ = p.add_run(text[pos:s])
                                r_.font.size = Pt(12); r_.font.name = "Calibri"
                                r_.font.color.rgb = rgb("2C2C2C")
                            r_bold = p.add_run(text[s:e])
                            r_bold.font.size = Pt(12); r_bold.font.name = "Calibri"
                            r_bold.font.color.rgb = rgb("2C2C2C")
                            r_bold.bold = True
                            pos = e
                        if pos < len(text):
                            r_ = p.add_run(text[pos:])
                            r_.font.size = Pt(12); r_.font.name = "Calibri"
                            r_.font.color.rgb = rgb("2C2C2C")
                    else:
                        run = p.add_run(text)
                        run.font.size = Pt(12); run.font.name = "Calibri"
                        run.font.color.rgb = rgb("2C2C2C")
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    run = p.add_run(text)
                    run.font.size = Pt(12)
                    run.font.name = "Calibri"
                    if j == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        run.font.color.rgb = rgb("2C2C2C")
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run.font.color.rgb = rgb("888888")

        section = doc.sections[0]
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_run = footer_para.add_run(f"For internal use only | {event_name or 'Event'}")
        footer_run.font.size = Pt(8)
        footer_run.font.color.rgb = rgb("999999")
        footer_run.font.name = "Calibri"
        footer_run.italic = True

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf, None

    except Exception as e:
        import traceback
        return None, f"{type(e).__name__}: {e}\\n{traceback.format_exc()[-500:]}"


def is_activity_item(item):
    head = (item or '').strip().lower()
    prefixes = [
        'welcome address', 'film', 'industry perspective', 'special address',
        'biotech sector perspective', 'it & deeptech sector perspective',
        'response', 'keynote', 'inaugural address', 'opening remarks',
        'vote of thanks', 'panel discussion', 'interaction', 'avgc sector perspective',
        'startup ceos interaction', 'vc interaction', 'open house conversation',
        'meet and greet', 'breakfast', 'lunch', 'tea', 'plenary', 'address'
    ]
    return any(head.startswith(p) for p in prefixes)

def extract_all_names(rows, dais_text=""):
    import re as _re2
    seen = set()
    results = []

    def add_name(name_part, title='', company='', source=''):
        name_part = (name_part or '').strip(' ,.-')
        title = (title or '').strip(' ,.-')
        company = (company or '').strip(' ,.-')
        if not name_part:
            return
        key = _re2.sub(r'\s+', ' ', name_part).lower()
        if key in seen:
            return
        seen.add(key)
        results.append((name_part, title, company, source))

    if dais_text and str(dais_text).strip():
        for line in str(dais_text).splitlines():
            entry = line.strip(' •	-')
            if not entry:
                continue
            parts = [p.strip() for p in _re2.split(r'\s*[,:\-–]\s*', entry) if p.strip()]
            name = parts[0] if parts else ''
            title = parts[1] if len(parts) > 1 else ''
            company = ', '.join(parts[2:]) if len(parts) > 2 else ''
            add_name(name, title, company, 'custom dais list')
        if results:
            return results

    title_pattern = _re2.compile(r"(?:Shri|Smt\.?|Dr\.?|Mr\.?|Ms\.?|Mrs\.?|Prof\.?|H\.E\.?|Hon['']?ble|Sri|Padma(?:shri|bhushan|vibhushan)|Padma Shri)\s+[A-Z][A-Za-z]", _re2.IGNORECASE)
    by_pattern = _re2.compile(r"by", _re2.IGNORECASE)

    for row in rows:
        item = row.get('item', '')
        if not item:
            continue
        by_matches = list(by_pattern.finditer(item))
        if by_matches:
            rest = item[by_matches[-1].end():].strip().rstrip('.')
            if rest:
                parts = [p.strip() for p in rest.split(',') if p.strip()]
                name = parts[0] if parts else ''
                title = parts[1] if len(parts) > 1 else ''
                company = ', '.join(parts[2:]) if len(parts) > 2 else ''
                add_name(name, title, company, item)
        for m in title_pattern.finditer(item):
            fragment = item[m.start():]
            stop = _re2.search(r"(?:and|followed by|who|will|to|for)", fragment, _re2.I)
            if stop:
                fragment = fragment[:stop.start()]
            fragment = fragment.strip().rstrip('.,')
            parts = [p.strip() for p in fragment.split(',') if p.strip()]
            name = parts[0] if parts else ''
            title = parts[1] if len(parts) > 1 else ''
            company = ', '.join(parts[2:]) if len(parts) > 2 else ''
            if name:
                add_name(name, title, company, item)
    return results


def build_mc_doc(event_name, event_date, venue, rows, logo_bytes=None, lang_code="kan", dais_text=""):
    try:
        from docx import Document as DocxDoc
        from docx.shared import Pt, RGBColor, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_ALIGN_VERTICAL
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError as e:
        return None, f"python-docx not available: {e}"

    try:
        doc = DocxDoc()
        for section in doc.sections:
            section.top_margin = Cm(1.6)
            section.bottom_margin = Cm(1.6)
            section.left_margin = Cm(1.8)
            section.right_margin = Cm(1.8)

        normal = doc.styles['Normal']
        normal.font.name = 'Calibri'
        normal.font.size = Pt(11)
        NAVY = "1A2B4C"

        def rgb(h):
            h = h.lstrip('#')
            return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

        def set_cell_bg(cell, hex_color):
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto')
            shd.set(qn('w:fill'), hex_color); tcPr.append(shd)

        def visible_borders(cell):
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            for side in ['top','left','bottom','right']:
                el = OxmlElement(f'w:{side}')
                el.set(qn('w:val'),'single')
                el.set(qn('w:sz'),'4')
                el.set(qn('w:color'),'D9D9D9')
                tcBorders.append(el)
            tcPr.append(tcBorders)

        footer = doc.sections[0].footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr = fp.add_run(f"For internal use only | {event_name or 'Event'}")
        fr.font.size = Pt(8); fr.font.color.rgb = rgb("999999"); fr.italic = True

        heading = doc.add_paragraph()
        hr = heading.add_run("Dignitaries on the Dais")
        hr.bold = True; hr.font.size = Pt(20); hr.font.color.rgb = rgb(NAVY); hr.font.name = 'Calibri'

        subh = doc.add_paragraph()
        sr = subh.add_run(f"{event_name or 'Event'} | {event_date or ''} | {venue or ''}")
        sr.font.size = Pt(10); sr.font.color.rgb = rgb("555555"); sr.italic = True; sr.font.name = 'Calibri'

        all_names = extract_all_names(rows, dais_text)
        if all_names:
            dais_table = doc.add_table(rows=1, cols=1)
            dais_table.autofit = False
            hdr = dais_table.rows[0].cells[0]
            hdr.width = Cm(17.4)
            set_cell_bg(hdr, "F5F5F5"); visible_borders(hdr)
            p = hdr.paragraphs[0]
            r = p.add_run("Dignitaries on the Dais")
            r.bold = True; r.font.size = Pt(11); r.font.name = 'Calibri'
            for i, (name, title, company, source) in enumerate(all_names):
                tr = dais_table.add_row()
                cell = tr.cells[0]
                cell.width = Cm(17.4)
                set_cell_bg(cell, "FFFFFF" if i % 2 == 0 else "FAFAFA"); visible_borders(cell)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                cell.text = ""
                p1 = cell.paragraphs[0]
                r1 = p1.add_run(name)
                r1.font.size = Pt(12)
                r1.font.name = 'Calibri'
                r1.bold = True
                if title:
                    p2 = cell.add_paragraph()
                    r2 = p2.add_run(title)
                    r2.font.size = Pt(11.5)
                    r2.font.name = 'Calibri'
                    r2.bold = False
                if company:
                    p3 = cell.add_paragraph()
                    r3 = p3.add_run(company)
                    r3.font.size = Pt(11.5)
                    r3.font.name = 'Calibri'
                    r3.bold = False
        else:
            doc.add_paragraph().add_run("No names detected in the programme.")

        doc.add_page_break()
        _mc_lang_label = "हिन्दी" if lang_code == "hin" else "ಕನ್ನಡ"
        mc_heading = doc.add_paragraph()
        mhr = mc_heading.add_run(f"MC Script — Bilingual (English + {_mc_lang_label})")
        mhr.bold = True; mhr.font.size = Pt(16); mhr.font.color.rgb = rgb(NAVY); mhr.font.name = 'Calibri'

        note_p = doc.add_paragraph()
        nr = note_p.add_run("Replace all text in [brackets] with actual names/designations before the event.")
        nr.italic = True; nr.font.size = Pt(9); nr.font.color.rgb = rgb("888888"); nr.font.name = 'Calibri'

        for i, row in enumerate(rows):
            eng, kan = get_script(row['item'], lang_code)
            item_para = doc.add_paragraph()
            ir = item_para.add_run(f"{i+1}. {row['item']}")
            if is_activity_item(row['item']):
                ir.bold = True
            ir.font.size = Pt(11)
            ir.font.color.rgb = rgb("1A5276") if is_address(row['item']) else rgb(NAVY)
            ir.font.name = 'Calibri'
            sr2 = item_para.add_run(f" — {row['slot']}")
            sr2.font.size = Pt(10); sr2.font.color.rgb = rgb("555555"); sr2.font.name = 'Calibri'

            sc_table = doc.add_table(rows=1, cols=2)
            sc_table.autofit = False
            cells = sc_table.rows[0].cells
            cells[0].width = Cm(8.5); cells[1].width = Cm(8.9)
            for cell in cells:
                visible_borders(cell); set_cell_bg(cell, "FFFFFF")

            p_eng = cells[0].paragraphs[0]
            lbl_e = p_eng.add_run("English\n")
            lbl_e.bold = True; lbl_e.font.size = Pt(9); lbl_e.font.color.rgb = rgb("1A5276"); lbl_e.font.name = 'Calibri'
            body_e = p_eng.add_run(eng); body_e.font.size = Pt(9); body_e.font.name = 'Calibri'

            p_kan = cells[1].paragraphs[0]
            lbl_k = p_kan.add_run(f"{_mc_lang_label}\n")
            lbl_k.bold = True; lbl_k.font.size = Pt(9); lbl_k.font.color.rgb = rgb("7B5200"); lbl_k.font.name = 'Nirmala UI'
            body_k = p_kan.add_run(kan); body_k.font.size = Pt(9); body_k.font.name = 'Nirmala UI'

        notes_rows = [(r['item'], r.get('remarks','')) for r in rows if r.get('remarks','').strip()]
        if notes_rows:
            doc.add_page_break()
            notes_h = doc.add_paragraph()
            nhr = notes_h.add_run("Notes")
            nhr.bold = True; nhr.font.size = Pt(20); nhr.font.color.rgb = rgb(NAVY); nhr.font.name = 'Calibri'
            notes_table = doc.add_table(rows=1, cols=2)
            notes_table.autofit = False
            nhdr = notes_table.rows[0].cells
            nhdr[0].width = Cm(8.0); nhdr[1].width = Cm(9.4)
            for cell, text in zip(nhdr, ["Programme Item", "Notes / Instructions"]):
                set_cell_bg(cell, "F5F5F5"); visible_borders(cell)
                p = cell.paragraphs[0]
                run = p.add_run(text); run.bold = True; run.font.size = Pt(11); run.font.name = 'Calibri'
            for i, (item, note) in enumerate(notes_rows):
                tr = notes_table.add_row()
                bg = "FFFFFF" if i % 2 == 0 else "FAFAFA"
                tr.cells[0].width = Cm(8.0); tr.cells[1].width = Cm(9.4)
                for cell, text in zip(tr.cells, [item, note]):
                    set_cell_bg(cell, bg); visible_borders(cell)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                    p = cell.paragraphs[0]
                    run = p.add_run(text); run.font.size = Pt(11); run.font.name = 'Calibri'

        buf = io.BytesIO()
        doc.save(buf); buf.seek(0)
        return buf, None
    except Exception as e:
        import traceback
        return None, f"{type(e).__name__}: {e}\n{traceback.format_exc()[-500:]}"

def build_excel(event_name, event_date, venue, rows, logo_bytes=None, lang_code="kan"):
    MAROON="7B1B1B"; GOLD="C9A84C"; WHITE="FFFFFF"; DARK="2C2C2C"
    LT_GOLD="F5E6C8"

    def sd(s="thin",c=GOLD): return Side(style=s,color=c)
    def bdr(s="thin"): b=sd(s); return Border(left=b,right=b,top=b,bottom=b)
    def fill(h): return PatternFill("solid",fgColor=h)
    def fnt(bold=False,color=DARK,size=10,italic=False):
        return Font(name="Arial",bold=bold,color=color,size=size,italic=italic)
    def aln(h="center",v="center",wrap=False):
        return Alignment(horizontal=h,vertical=v,wrap_text=wrap)
    def mbdr():
        m=Side(style="medium",color=MAROON)
        return Border(left=m,right=m,top=m,bottom=m)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Programme Planner"
    for col,w in {"A":22,"B":52,"C":18}.items():
        ws1.column_dimensions[col].width = w

    ws1.row_dimensions[1].height = 26
    ws1.row_dimensions[2].height = 18
    ws1.row_dimensions[3].height = 18
    ws1.row_dimensions[4].height = 18

    c = ws1["A1"]
    c.value = f"M2M Programme for Inaugural of {event_name or 'Event'}"
    c.font = Font(name="Arial",bold=True,color=WHITE,size=14)
    c.fill = fill(MAROON); c.alignment = aln(h="left")
    ws1.merge_cells("A1:B1")
    c2 = ws1["C1"]; c2.fill = fill(MAROON)

    details_rows = [
        ("Event :", event_name or "—"),
        ("Date :", event_date or "—"),
        ("Venue :", venue or "—"),
        ("Start Time :", rows[0]["start_str"] if rows else "—"),
    ]

    for di, (lbl, val) in enumerate(details_rows):
        r = 2 + di
        ws1.row_dimensions[r].height = 18
        cl = ws1.cell(row=r, column=1)
        cl.value = lbl
        cl.font = Font(name="Arial",bold=True,color=MAROON,size=9)
        cl.fill = fill(LT_GOLD); cl.alignment = aln(h="right")
        cv = ws1.cell(row=r, column=2)
        cv.value = val
        cv.font = Font(name="Arial",color=DARK,size=9)
        cv.fill = fill(WHITE); cv.border = bdr()
        cv.alignment = aln(h="left")

    if logo_bytes:
        from openpyxl.drawing.image import Image as XLImage
        try:
            img = XLImage(io.BytesIO(logo_bytes))
            img.height = 95; img.width = 170; img.anchor = "C1"
            ws1.add_image(img)
            ws1.column_dimensions["C"].width = 24
        except Exception:
            pass

    ws1.row_dimensions[6].height = 24
    for ci, hdr in enumerate(["Time Slot","Programme Item / Activity"], 1):
        c = ws1.cell(row=6, column=ci)
        c.value = hdr; c.font = Font(name="Arial",bold=True,color=WHITE,size=10)
        c.fill = fill(MAROON); c.alignment = aln(wrap=True); c.border = bdr()

    ws1.column_dimensions["D"].width = 12
    ws1.column_dimensions["E"].width = 2
    ws1.column_dimensions["F"].width = 2

    start_time_str = rows[0]['start_str'] if rows else "10:00 AM"
    seed_cell = ws1["F1"]
    seed_cell.value = f'=TIMEVALUE("{start_time_str}")'
    seed_cell.number_format = "h:mm AM/PM"
    ws1.column_dimensions["F"].width = 0

    dur_hdr = ws1.cell(row=6, column=4)
    dur_hdr.value = "Duration\n(mins)"
    dur_hdr.font = Font(name="Arial",bold=True,color=WHITE,size=10)
    dur_hdr.fill = fill(MAROON); dur_hdr.alignment = aln(wrap=True); dur_hdr.border = bdr()

    for i, row in enumerate(rows):
        r = 7+i
        ws1.row_dimensions[r].height = 22
        item = row.get('item','')
        addr = is_address(item)
        duration = row.get('duration', 0)

        dcell = ws1.cell(row=r, column=4)
        dcell.value = duration
        dcell.font = fnt(color=MAROON, bold=True)
        dcell.fill = PatternFill("solid", fgColor="FFF8EE")
        dcell.alignment = aln(h="center")
        dcell.border = bdr()

        ecell = ws1.cell(row=r, column=5)
        if i == 0:
            ecell.value = "=$F$1"
        else:
            ecell.value = f"=F{r-1}"
        ecell.number_format = "h:mm AM/PM"
        ecell.font = Font(size=1, color="FFFFFF")

        fcell = ws1.cell(row=r, column=6)
        fcell.value = f"=E{r}+D{r}/1440"
        fcell.number_format = "h:mm AM/PM"
        fcell.font = Font(size=1, color="FFFFFF")

        acell = ws1.cell(row=r, column=1)
        acell.value = f'=TEXT(E{r},"h:mm AM/PM")&" – "&TEXT(F{r},"h:mm AM/PM")'
        acell.fill = PatternFill("solid", fgColor="FFFFFF")
        acell.font = fnt(color=DARK)
        acell.alignment = aln(h="left", wrap=False)
        acell.border = bdr()

        bcell = ws1.cell(row=r, column=2)
        bcell.value = item
        bcell.fill = PatternFill("solid", fgColor="FFFFFF")
        bcell.font = fnt(color=DARK, bold=addr)
        bcell.alignment = aln(h="left", wrap=True)
        bcell.border = bdr()

    tr = 7 + len(rows)
    ws1.row_dimensions[tr].height = 22
    last_row = 7 + len(rows) - 1
    c = ws1[f"A{tr}"]
    c.value = f'="Programme ends at "&TEXT(F{last_row},"h:mm AM/PM")&" (auto-calculated)"'
    c.font = Font(name="Arial",bold=True,color=GOLD,size=10)
    c.fill = fill(DARK); c.alignment = aln(h="left"); c.border = mbdr()
    ws1.merge_cells(f"A{tr}:B{tr}")
    ws1.freeze_panes = "A7"

    note_row = tr + 2
    ws1.merge_cells(f"A{note_row}:B{note_row}")
    nc = ws1[f"A{note_row}"]
    nc.value = ("ℹ️ Edit the Duration column (D) or Start Time (cell F1) — "
                "Time Slots above recalculate automatically. Use this sheet as an offline "
                "backup if the web app is unavailable.")
    nc.font = fnt(italic=True, color="888888", size=8)
    nc.alignment = aln(h="left", wrap=True)
    ws1.row_dimensions[note_row].height = 28

    ws2 = wb.create_sheet("Print View")
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 52
    ws2.column_dimensions["C"].width = 22
    PP = "'Programme Planner'"

    ws2.row_dimensions[1].height = 26
    c = ws2["A1"]
    c.value = f"Minute-to-Minute Programme for Inaugural of {event_name or 'Event'}"
    c.font = Font(name="Arial",bold=True,color=WHITE,size=14)
    c.fill = fill(MAROON); c.alignment = aln(h="left")
    ws2.merge_cells("A1:B1")

    detail_labels = ["Event :","Date :","Venue :","Start Time :"]
    for i, lbl in enumerate(detail_labels, start=2):
        ws2[f"A{i}"] = lbl
        ws2[f"A{i}"].font = Font(name="Arial",bold=True,color=MAROON,size=9)
        ws2[f"A{i}"].fill = fill(LT_GOLD)
        ws2[f"A{i}"].alignment = aln(h="right")

    ws2["B2"] = f'=IF({PP}!B2="","—",{PP}!B2)'
    ws2["B3"] = f'=IF({PP}!B3="","—",{PP}!B3)'
    ws2["B4"] = f'=IF({PP}!B4="","—",{PP}!B4)'
    ws2["B5"] = f'=IF({PP}!B5="","—",{PP}!B5)'
    for r in range(2,6):
        ws2[f"B{r}"].font = Font(name="Arial",color=DARK,size=9)
        ws2[f"B{r}"].fill = fill(WHITE)
        ws2[f"B{r}"].alignment = aln(h="left")

    ws2.row_dimensions[7].height = 24
    for ci, hdr in enumerate(["Time Slot","Programme Item / Activity"], 1):
        c = ws2.cell(row=7, column=ci)
        c.value = hdr; c.font = Font(name="Arial",bold=True,color=WHITE,size=10)
        c.fill = fill(MAROON); c.alignment = aln(wrap=True); c.border = bdr()

    for i, row in enumerate(rows):
        r = 8 + i
        ws2.row_dimensions[r].height = 22
        ac = ws2.cell(row=r, column=1)
        bc = ws2.cell(row=r, column=2)
        ac.value = f"={PP}!A{7+i}"
        bc.value = f"={PP}!B{7+i}"
        ac.font = fnt(color=DARK)
        bc.font = fnt(color=DARK, bold=is_address(row.get('item','')))
        ac.fill = fill(WHITE); bc.fill = fill(WHITE)
        ac.border = bdr(); bc.border = bdr()
        ac.alignment = aln(h="left")
        bc.alignment = aln(h="left", wrap=True)

    tr2 = 8 + len(rows)
    ws2[f"A{tr2}"] = f'=IF({PP}!A{7+len(rows)}<>"",{PP}!A{7+len(rows)},"")'
    ws2.merge_cells(start_row=tr2, start_column=1, end_row=tr2, end_column=2)
    ws2[f"A{tr2}"].font = Font(name="Arial",bold=True,color=GOLD,size=10)
    ws2[f"A{tr2}"].fill = fill(DARK)
    ws2[f"A{tr2}"].alignment = aln(h="left")

    ws3 = wb.create_sheet("MC Script")
    ws3.column_dimensions["A"].width = 52
    ws3.column_dimensions["B"].width = 52

    ws3["A1"] = "MC Script"
    ws3["A1"].font = Font(name="Arial",bold=True,color=WHITE,size=14)
    ws3["A1"].fill = fill(MAROON)
    ws3["A1"].alignment = aln(h="left")
    ws3.merge_cells("A1:B1")

    row_num = 3
    for i, row in enumerate(rows):
        eng, kan = get_script(row['item'], lang_code)
        ws3[f"A{row_num}"] = row['item']
        ws3[f"B{row_num}"] = row['slot']
        ws3[f"A{row_num}"].font = Font(name="Arial",bold=True,color=DARK,size=10)
        ws3[f"B{row_num}"].font = Font(name="Arial",color="666666",size=9)
        row_num += 1
        ws3[f"A{row_num}"] = eng
        ws3[f"B{row_num}"] = kan
        ws3[f"A{row_num}"].font = Font(name="Arial",size=9)
        ws3[f"B{row_num}"].font = Font(name="Nirmala UI",size=9)
        row_num += 2

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Simple title ───────────────────────────────────────────────────────────────
st.title("Minute-to-Minute Programme Planner")

# ── Sidebar ───────────────────────────────────────────────────────────────────
if st.session_state.get("last_entered"):
    top_cols = st.columns([4,1])
    with top_cols[0]:
        st.info("A previous session is available. Click to reload the last entered event details and programme items.")
    with top_cols[1]:
        if st.button("🔁 Reload Last Entered", use_container_width=True, type="primary"):
            le = st.session_state.last_entered
            st.session_state["_event_name_val"] = le.get("event_name","")
            st.session_state["_event_date_val"] = le.get("event_date","")
            st.session_state["_venue_val"] = le.get("venue","")
            st.session_state["_start_time_val"] = le.get("start_time","10:00 AM")
            if le.get("programme_items"):
                st.session_state.programme_items = [dict(x) for x in le["programme_items"]]
            st.rerun()

with st.sidebar:
    st.markdown("### ⚙️ Event Details")

    event_name = st.text_input("Event Name", value=st.session_state.get("_event_name_val",""), placeholder="e.g. BTS 2025 Inauguration", key="_event_name_val")
    event_date = st.text_input("Event Date", value=st.session_state.get("_event_date_val",""), placeholder="e.g. 15-Aug-2025", key="_event_date_val")
    venue = st.text_input("Venue", value=st.session_state.get("_venue_val",""), placeholder="e.g. Taj Vivanta, Bengaluru", key="_venue_val")
    start_time = st.text_input("Start Time", value=st.session_state.get("_start_time_val","10:00 AM"), key="_start_time_val")

    st.markdown("---")
    st.markdown("### 🗣️ MC Script Language")
    lang_choice = st.radio(
        "Bilingual pairing for MC Script",
        options=list(LANGUAGE_OPTIONS.keys()),
        index=0,
        help="Choose the regional language to pair with English for this event"
    )
    lang_code = LANGUAGE_OPTIONS[lang_choice]["code"]
    lang_label = LANGUAGE_OPTIONS[lang_choice]["label"]

    st.markdown("---")
    st.markdown("### 🖼️ Event Logo")
    logo_file = st.file_uploader(
        "Upload logo (PNG or JPG)",
        type=["png","jpg","jpeg"],
        help="Logo will appear in downloaded Excel and Word files"
    )
    logo_bytes = logo_file.read() if logo_file else None

    st.markdown("---")
    st.markdown("### ℹ️ How to use")
    st.markdown("""
1. Fill event details above
2. Upload logo (optional)
3. Add/edit programme items below
4. Download Excel or Word
""")
if logo_bytes:
    st.info("✅ Logo uploaded — it will appear in your downloaded Excel and Word files.")

# ── Event info panel ──────────────────────────────────────────────────────────
if event_name or event_date or venue:
    info_parts = []
    if event_name: info_parts.append(f"<b>{event_name}</b>")
    if event_date: info_parts.append(f"📅 {event_date}")
    if venue:      info_parts.append(f"📍 {venue}")
    if start_time: info_parts.append(f"🕐 {start_time}")
    st.markdown(
        f'<div class="event-info-panel">'
        f'{"  &nbsp;|&nbsp;  ".join(info_parts)}</div>',
        unsafe_allow_html=True)

# ── Parse start time ──────────────────────────────────────────────────────────
start_dt = parse_time(start_time) if start_time else parse_time("10:00 AM")
if not start_dt:
    st.error(f"⚠️ Can't parse '{start_time}'. Use format like '10:00 AM' or '09:30 AM'")
    st.stop()

# ── Blank default rows ─────────────────────────────────────────────────────────
DEFAULT_ROWS = 20
if "programme_items" not in st.session_state:
    st.session_state.programme_items = [
        {"item": "", "duration": 5, "remarks": ""} for _ in range(DEFAULT_ROWS)
    ]

# ── Programme table (editable) ────────────────────────────────────────────────
st.markdown('<div class="section-header">📝 Programme Items</div>', unsafe_allow_html=True)

df_input = pd.DataFrame(st.session_state.programme_items)
for col in ["item", "duration", "remarks"]:
    if col not in df_input.columns:
        df_input[col] = "" if col != "duration" else 5

df_input = df_input[["item", "duration", "remarks"]]

edited_df = st.data_editor(
    df_input,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "item": st.column_config.TextColumn("Programme Item", width="large"),
        "duration": st.column_config.NumberColumn("Duration (mins)", min_value=1, max_value=180, step=1),
        "remarks": st.column_config.TextColumn("Remarks / Notes", width="medium"),
    },
    key="programme_editor",
)

edited_df["item"] = edited_df["item"].apply(normalize_item_text)
edited_df["remarks"] = edited_df["remarks"].fillna("").astype(str)
st.session_state.programme_items = edited_df.to_dict("records")

rows = []
for r in st.session_state.programme_items:
    item = str(r.get("item", "")).strip()
    dur = r.get("duration", 0)
    if not item:
        continue
    try:
        dur = int(dur or 0)
    except Exception:
        dur = 0
    if dur <= 0:
        continue
    rows.append({
        "item": item,
        "duration": dur,
        "remarks": str(r.get("remarks","")).strip(),
    })

if not rows:
    st.info("Add one or more programme items to generate downloads. Blank rows are ignored.")
    rows = []
else:
    rows = compute_slots(rows, start_dt) if rows else []

# ── Programme preview ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">🗂️ Programme Preview</div>', unsafe_allow_html=True)
for row in rows:
    st.markdown(
        f"<div style='padding:6px 0;border-bottom:1px solid #eee;'>"
        f"<span class='time-pill'>{row['slot']}</span>&nbsp;&nbsp;"
        f"<b>{row['item']}</b></div>",
        unsafe_allow_html=True,
    )

# ── Downloads ──────────────────────────────────────────────────────────────────
st.divider()
st.markdown("### Dignitaries on the Dais")
dais_text = st.text_area(
    "Optional: paste dignitaries for MC Copy (one per line)",
    value=st.session_state.get("dais_text", ""),
    height=180,
    placeholder="""Example:
Shri S. Gopalakrishnan (Kris), Chairperson
Vision Group on IT, Govt. of Karnataka
Dr. Kiran Mazumdar-Shaw, Chairperson
Vision Group on Biotech, Govt. of Karnataka
Dr. G. Parameshwara, Hon'ble Deputy Chief Minister
Govt. of Karnataka""",
    key="dais_text_input",
)
st.caption("If this box is filled, the MC Document will use this list. If left blank, names will be pulled from the M2M programme.")
st.session_state["dais_text"] = dais_text


st.info("Paste one dignitary per line. The MC document will keep line breaks, place the name first (12 pt), then title/company below (11.5 pt), and bold activity titles like Welcome Address, Film, Industry Perspective, and Special Address.")
st.markdown("### 📥 Download")

fname_base = f"Minute-to-Minute Programme for {safe_filename(event_name)}" if event_name else "Minute-to-Minute Programme"

dcol1, dcol2, dcol3 = st.columns(3)

with dcol1:
    excel_buf = build_excel(event_name, event_date, venue, rows, logo_bytes, lang_code)
    st.download_button(
        label="⬇️ Download Excel (.xlsx)",
        data=excel_buf,
        file_name=f"{fname_base}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("3 sheets: Programme Planner · Print View · MC Script")

with dcol2:
    word_buf, word_err = build_word(event_name, event_date, venue, rows, logo_bytes, lang_code)
    if word_buf:
        st.download_button(
            label="⬇️ Download M2M (.docx)",
            data=word_buf,
            file_name=f"{fname_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.caption("Programme table only — clean M2M for sharing")
        save_generated_files(
            meta={
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "file_name": fname_base,
                "event_name": event_name,
                "event_date": event_date,
                "venue": venue,
                "start_time": start_time,
            },
            excel_bytes=excel_buf.getvalue() if hasattr(excel_buf, "getvalue") else excel_buf,
            word_bytes=word_buf.getvalue() if hasattr(word_buf, "getvalue") else word_buf,
            programme_items=st.session_state.get("programme_items", []),
        )
    else:
        st.error(f"Word export failed: {word_err}")

with dcol3:
    mc_buf, mc_err = build_mc_doc(event_name, event_date, venue, rows, logo_bytes, lang_code, dais_text)
    if mc_buf:
        st.download_button(
            label="⬇️ Download MC Document (.docx)",
            data=mc_buf,
            file_name=f"{fname_base} - MC Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.caption("Dignitaries list · Bilingual MC Script · Notes")
    else:
        st.error(f"MC Document export failed: {mc_err}")

render_history(limit=MAX_HISTORY)

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.markdown("""
<div style="text-align:center;color:#888;font-size:0.8rem">
    Minute-to-Minute Programme · Built for Event Management Teams
</div>""", unsafe_allow_html=True)
