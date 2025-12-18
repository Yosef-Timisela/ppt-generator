import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import matplotlib.pyplot as plt
import pandas as pd
import io
import uuid

# =============================================
# EXPERT PPT GENERATOR — PPT THEME UPLOAD & AUTO SUGGESTION
# =============================================

st.set_page_config(page_title="Expert PPT Generator", page_icon="🎨", layout="wide")

# ------------------------------
# SESSION STATE
# ------------------------------
if "slides" not in st.session_state:
    st.session_state.slides = []
if "themes" not in st.session_state:
    st.session_state.themes = {}

# ------------------------------
# HELPER FUNCTIONS
# ------------------------------
def load_theme_ppt(ppt_file):
    prs = Presentation(ppt_file)
    return prs.slide_width, prs.slide_height, prs.slide_layouts


def suggest_theme(slides):
    """Simple rule-based theme suggestion"""
    text_count = sum(1 for s in slides if s['type'] == 'Text')
    chart_count = sum(1 for s in slides if 'Chart' in s['type'])

    if chart_count > text_count:
        return "Business"
    if text_count > chart_count:
        return "Minimal"
    return "Creative"

# ------------------------------
# CUSTOM CSS
# ------------------------------
st.markdown("""
<style>
.main {background-color:#eef2f7;}
.card {background:white;padding:1.2rem;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.08);margin-bottom:1rem;}
.preview {background:#334155;border-radius:16px;padding:1.5rem;min-height:200px;color:white;}
.stButton>button {background:linear-gradient(135deg,#9333ea,#4f46e5);color:white;border-radius:12px;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ------------------------------
# HEADER
# ------------------------------
st.title("🎨 Expert PPT Generator with Theme Engine")
st.caption("Upload PPT themes • Auto theme suggestion • Smart generation")

# ------------------------------
# SIDEBAR — THEME ENGINE
# ------------------------------
st.sidebar.header("🎨 Theme Engine")

theme_ppt = st.sidebar.file_uploader("Upload Theme PPT (.pptx)", type=["pptx"])

if theme_ppt:
    theme_prs = Presentation(theme_ppt)
    st.session_state.themes[theme_ppt.name] = theme_prs
    st.sidebar.success(f"Theme '{theme_ppt.name}' loaded")

if st.session_state.themes:
    selected_theme = st.sidebar.selectbox(
        "Select Theme",
        list(st.session_state.themes.keys())
    )
else:
    selected_theme = None

# ------------------------------
# SLIDE CREATOR
# ------------------------------
st.subheader("➕ Create Slide")
slide_type = st.selectbox("Slide Type", ["Text","Bullet","Chart (CSV)"])
title = st.text_input("Slide Title")
content = st.text_area("Slide Content / Bullet")

if slide_type == "Chart (CSV)":
    csv_file = st.file_uploader("Upload CSV", type=["csv"])
else:
    csv_file = None

if st.button("Add Slide") and title:
    st.session_state.slides.append({
        "id": str(uuid.uuid4()),
        "type": slide_type,
        "title": title,
        "content": content,
        "csv": csv_file
    })

# ------------------------------
# THEME SUGGESTION
# ------------------------------
if st.session_state.slides:
    recommended = suggest_theme(st.session_state.slides)
    st.info(f"💡 Recommended theme type: **{recommended}**")

# ------------------------------
# PREVIEW
# ------------------------------
st.subheader("👀 Slide Preview")
for slide in st.session_state.slides:
    st.markdown(f"""
    <div class="card">
        <div class="preview">
            <h3>{slide['title']}</h3>
            <p>{slide['content'].replace(chr(10), '<br>')}</p>
            <small>{slide['type']}</small>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ------------------------------
# EXPORT PPT USING THEME
# ------------------------------
if st.button("🚀 Export PowerPoint"):
    if selected_theme:
        prs = Presentation()
        base_theme = st.session_state.themes[selected_theme]
        prs.slide_width = base_theme.slide_width
        prs.slide_height = base_theme.slide_height
    else:
        prs = Presentation()

    for slide_data in st.session_state.slides:
        if selected_theme:
            slide = prs.slides.add_slide(base_theme.slide_layouts[1])
        else:
            slide = prs.slides.add_slide(prs.slide_layouts[6])

        slide.shapes.title.text = slide_data['title']

        if slide_data['type'] == "Bullet":
            tf = slide.placeholders[1].text_frame
            for line in slide_data['content'].split("\n"):
                p = tf.add_paragraph()
                p.text = line
                p.level = 1
        elif slide_data['type'] == "Chart (CSV)" and slide_data['csv']:
            df = pd.read_csv(slide_data['csv'])
            fig, ax = plt.subplots()
            df.plot(ax=ax)
            img = io.BytesIO()
            plt.savefig(img)
            plt.close()
            img.seek(0)
            slide.shapes.add_picture(img, Inches(1), Inches(2.5), Inches(6))
        else:
            slide.placeholders[1].text = slide_data['content']

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    st.success("🎉 PPT generated using selected theme")
    st.download_button(
        "⬇️ Download PPT",
        buffer,
        file_name="themed_expert_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
