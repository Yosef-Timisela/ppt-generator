import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io

# ==============================
# Interactive PPT Generator with Theme Background
# ==============================

st.set_page_config(
    page_title="Interactive PPT Generator",
    page_icon="🎨",
    layout="wide"
)

# ------------------------------
# Custom CSS for Interactive UI
# ------------------------------
st.markdown(
    """
    <style>
    .main {background-color: #f1f5f9;}
    .card {
        background: white;
        padding: 1.2rem;
        border-radius: 16px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.08);
        margin-bottom: 1rem;
    }
    .preview-slide {
        background-size: cover;
        background-position: center;
        padding: 1.5rem;
        border-radius: 16px;
        min-height: 180px;
        color: white;
        box-shadow: inset 0 0 0 2000px rgba(0,0,0,0.35);
    }
    .stButton>button {
        background: linear-gradient(135deg, #2563eb, #1e40af);
        color: white;
        border-radius: 12px;
        font-weight: 600;
        padding: 0.6em 1.4em;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ------------------------------
# Header
# ------------------------------
st.title("🎨 Interactive PowerPoint Generator")
st.caption("Customize theme, preview slides, then export professional PPT")

# ------------------------------
# Layout
# ------------------------------
editor, preview = st.columns([2, 1])

with editor:
    st.subheader("📝 Slide Editor")

    title = st.text_input("Presentation Title", "AI Project Presentation")
    subtitle = st.text_input("Subtitle / Author", "Your Name")

    st.markdown("#### 🎯 Slide 1")
    slide1_title = st.text_input("Slide 1 Title", "Introduction")
    slide1_points = st.text_area(
        "Bullet Points",
        "Background problem\nObjective\nScope",
        height=120
    )

    st.markdown("#### 🧾 Slide 2")
    slide2_title = st.text_input("Slide 2 Title", "Conclusion")
    slide2_content = st.text_area(
        "Slide Content",
        "Summary and final remarks",
        height=100
    )

    st.subheader("🎨 Theme Settings")
    bg_image = st.file_uploader("Upload Background Image (PNG/JPG)", type=["png", "jpg", "jpeg"])
    font_color = st.color_picker("Font Color", "#ffffff")

with preview:
    st.subheader("👀 Live Preview")

    bg_style = ""
    if bg_image:
        bg_style = f"background-image: url('data:image/png;base64,{bg_image.getvalue().hex()}')"

    st.markdown(
        f"""
        <div class="card">
            <div class="preview-slide" style="{bg_style}">
                <h3 style="color:{font_color}">{title}</h3>
                <p style="color:{font_color}"><em>{subtitle}</em></p>
            </div>
        </div>

        <div class="card">
            <div class="preview-slide" style="{bg_style}">
                <h4 style="color:{font_color}">{slide1_title}</h4>
                <ul style="color:{font_color}">
                    {''.join([f'<li>{p}</li>' for p in slide1_points.split('\n') if p])}
                </ul>
            </div>
        </div>

        <div class="card">
            <div class="preview-slide" style="{bg_style}">
                <h4 style="color:{font_color}">{slide2_title}</h4>
                <p style="color:{font_color}">{slide2_content}</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------------------
# Generate PPT with Background
# ------------------------------
if st.button("🚀 Generate PPT"):
    prs = Presentation()

    def add_bg(slide, image_bytes):
        if image_bytes:
            slide.shapes.add_picture(
                io.BytesIO(image_bytes),
                Inches(0), Inches(0),
                width=prs.slide_width,
                height=prs.slide_height
            )

    # Cover Slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg_image.getvalue() if bg_image else None)
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    tf = title_box.text_frame
    tf.text = title
    tf.paragraphs[0].font.size = Pt(40)
    tf.paragraphs[0].font.color.rgb = None
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    # Bullet Slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg_image.getvalue() if bg_image else None)
    box = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))
    tf = box.text_frame
    tf.text = slide1_title

    for p in slide1_points.split("\n"):
        para = tf.add_paragraph()
        para.text = p
        para.level = 1

    # Content Slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg_image.getvalue() if bg_image else None)
    box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
    tf = box.text_frame
    tf.text = slide2_title
    tf.add_paragraph().text = slide2_content

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    st.success("🎉 PPT created with custom theme!")
    st.download_button(
        "⬇️ Download PowerPoint",
        buffer,
        file_name="interactive_theme_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
