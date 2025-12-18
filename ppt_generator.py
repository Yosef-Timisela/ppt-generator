import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import io

# ==============================
# Streamlit PowerPoint Generator App
# ==============================

st.set_page_config(page_title="PPT Generator", layout="centered")

st.title("📊 PowerPoint Generator App")
st.write("Buat file PowerPoint (.pptx) secara otomatis menggunakan Python + Streamlit")

# ------------------------------
# User Inputs
# ------------------------------
title = st.text_input("Judul Presentasi", "Judul Presentasi")
subtitle = st.text_input("Subjudul", "Nama / Institusi")

st.subheader("Isi Slide")
slide1_title = st.text_input("Judul Slide 1", "Pendahuluan")
slide1_points = st.text_area(
    "Bullet Points (pisahkan dengan enter)",
    "Poin utama 1\nPoin utama 2\nPoin utama 3"
)

slide2_title = st.text_input("Judul Slide 2", "Kesimpulan")
slide2_content = st.text_area(
    "Isi Slide 2",
    "Ringkasan isi presentasi"
)

# ------------------------------
# Generate PPT
# ------------------------------
if st.button("🚀 Generate PowerPoint"):
    prs = Presentation()

    # Slide 1: Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

    # Slide 2: Bullet points
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = slide1_title
    tf = slide.placeholders[1].text_frame
    tf.text = slide1_points.split("\n")[0]

    for point in slide1_points.split("\n")[1:]:
        p = tf.add_paragraph()
        p.text = point
        p.level = 1

    # Slide 3: Content
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = slide2_title
    slide.placeholders[1].text = slide2_content

    # Save to buffer
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)

    st.success("PowerPoint berhasil dibuat!")
    st.download_button(
        label="⬇️ Download PowerPoint",
        data=ppt_buffer,
        file_name="presentasi_otomatis.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )