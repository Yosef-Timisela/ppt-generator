import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import matplotlib.pyplot as plt
import pandas as pd
import io

# =============================================
# ADVANCED INTERACTIVE PPT GENERATOR (ALL-IN-ONE)
# =============================================

st.set_page_config(page_title="Advanced PPT Generator", page_icon="🚀", layout="wide")

# ------------------------------
# SESSION STATE INIT
# ------------------------------
if "slides" not in st.session_state:
    st.session_state.slides = []

# ------------------------------
# CUSTOM CSS
# ------------------------------
st.markdown("""
<style>
.main {background-color:#f1f5f9;}
.card {background:white;padding:1.2rem;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,0.08);margin-bottom:1rem;}
.preview {background-size:cover;background-position:center;border-radius:16px;padding:1.5rem;min-height:180px;box-shadow: inset 0 0 0 2000px rgba(0,0,0,0.35);color:white;}
.stButton>button {background:linear-gradient(135deg,#2563eb,#1e40af);color:white;border-radius:12px;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ------------------------------
# HEADER
# ------------------------------
st.title("🚀 Advanced PowerPoint Generator")
st.caption("Multi-slide • Drag-like reorder • AI-ready • Charts • Themed templates")

# ------------------------------
# SIDEBAR SETTINGS
# ------------------------------
st.sidebar.header("🎨 Theme & Global Settings")

bg_image = st.sidebar.file_uploader("Background Image", type=["png","jpg","jpeg"])
font_color = st.sidebar.color_picker("Font Color", "#ffffff")
template = st.sidebar.selectbox("Template", ["Business","Minimal","Creative"])

# ------------------------------
# SLIDE CREATOR
# ------------------------------
st.subheader("➕ Add New Slide")

with st.form("add_slide"):
    slide_type = st.selectbox("Slide Type", ["Text","Bullet","Chart"])
    title = st.text_input("Slide Title")
    content = st.text_area("Slide Content / Bullet (new line)")
    submitted = st.form_submit_button("Add Slide")

    if submitted and title:
        st.session_state.slides.append({
            "type": slide_type,
            "title": title,
            "content": content
        })

# ------------------------------
# SLIDE PREVIEW & REORDER
# ------------------------------
st.subheader("👀 Slide Preview (Reorderable)")

for i, slide in enumerate(st.session_state.slides):
    col1, col2 = st.columns([5,1])
    with col1:
        st.markdown(f"""
        <div class="card">
            <div class="preview">
                <h3 style="color:{font_color}">{slide['title']}</h3>
                <p style="color:{font_color}">{slide['content'].replace(chr(10), '<br>')}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        if st.button("⬆️", key=f"up{i}") and i>0:
            st.session_state.slides[i-1], st.session_state.slides[i] = st.session_state.slides[i], st.session_state.slides[i-1]
        if st.button("⬇️", key=f"down{i}") and i < len(st.session_state.slides)-1:
            st.session_state.slides[i+1], st.session_state.slides[i] = st.session_state.slides[i], st.session_state.slides[i+1]
        if st.button("🗑️", key=f"del{i}"):
            st.session_state.slides.pop(i)
            st.rerun()

# ------------------------------
# GENERATE PPT
# ------------------------------
if st.button("🚀 Export PowerPoint"):
    prs = Presentation()

    def add_bg(slide):
        if bg_image:
            slide.shapes.add_picture(io.BytesIO(bg_image.getvalue()), Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

    for slide_data in st.session_state.slides:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))
        tf = box.text_frame
        tf.text = slide_data['title']
        tf.paragraphs[0].font.size = Pt(32)

        if slide_data['type'] == "Bullet":
            for line in slide_data['content'].split("\n"):
                p = tf.add_paragraph()
                p.text = line
                p.level = 1
        else:
            p = tf.add_paragraph()
            p.text = slide_data['content']

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    st.success("🎉 PowerPoint successfully generated")
    st.download_button("⬇️ Download PPT", buffer, file_name="advanced_presentation.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
