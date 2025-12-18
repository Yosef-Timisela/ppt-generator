import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import io

# ==============================
# Professional Streamlit PPT Generator with Preview
# ==============================

st.set_page_config(
    page_title="Professional PPT Generator",
    page_icon="📊",
    layout="wide"
)

# ------------------------------
# Custom CSS
# ------------------------------
st.markdown(
    """
    <style>
    .main {background-color: #f7f9fc;}
    h1, h2, h3 {font-family: 'Segoe UI', sans-serif;}
    .slide-card {
        background: white;
        padding: 1rem;
        border-radius: 14px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.06);
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #2563eb;
        color: white;
        border-radius: 10px;
        padding: 0.6em 1.2em;
        font-weight: 600;
    }
    .stDownloadButton>button {
        background-color: #16a34a;
        color: white;
        border-radius: 10px;
        padding: 0.6em 1.2em;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ------------------------------
# Header
# ------------------------------
st.title("📊 Professional PowerPoint Generator")
st.caption("Preview your slides before downloading the presentation")

# ------------------------------
# Layout
# ------------------------------
editor, preview = st.columns([2, 1])

with editor:
    st.subheader("📝 Slide Editor")

    title = st.text_input("Presentation Title", "Your Presentation Title")
    subtitle = st.text_input("Subtitle / Author", "Your Name or Organization")

    slide1_title = st.text_input("Slide 1 Title", "Introduction")
    slide1_points = st.text_area(
        "Slide 1 Bullet Points",
        "Key point one\nKey point two\nKey point three",
        height=150
    )

    slide2_title = st.text_input("Slide 2 Title", "Conclusion")
    slide2_content = st.text_area(
        "Slide 2 Content",
        "Summary and final thoughts",
        height=120
    )

with preview:
    st.subheader("👀 Slide Preview")

    st.markdown(
        f"""
        <div class="slide-card">
            <h3>{title}</h3>
            <p><em>{subtitle}</em></p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="slide-card">
            <h3>{slide1_title}</h3>
            <ul>
                {''.join([f'<li>{p}</li>' for p in slide1_points.split('\n') if p])}
            </ul>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div class="slide-card">
            <h3>{slide2_title}</h3>
            <p>{slide2_content}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

# ------------------------------
# Generate PPT
# ------------------------------
if st.button("🚀 Generate & Download PPT"):
    prs = Presentation()

    # Cover Slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = subtitle

    # Bullet Slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = slide1_title
    tf = slide.placeholders[1].text_frame
    tf.text = slide1_points.split("\n")[0]

    for point in slide1_points.split("\n")[1:]:
        p = tf.add_paragraph()
        p.text = point
        p.level = 1

    # Content Slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = slide2_title
    slide.placeholders[1].text = slide2_content

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)

    st.success("🎉 Presentation ready!")
    st.download_button(
        "⬇️ Download PowerPoint",
        buffer,
        file_name="professional_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
