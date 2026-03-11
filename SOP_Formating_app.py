import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ------------------ PAGE CONFIG ------------------

st.set_page_config(
    page_title="SOP Formatter by S M Baqir",
    page_icon="📄",
    layout="centered"
)

# ------------------ PREMIUM STYLE ------------------

st.markdown("""
<style>
.main-title{
    font-size:38px;
    font-weight:700;
    text-align:center;
    color:#1f4e79;
}

.subtitle{
    text-align:center;
    color:gray;
    margin-bottom:30px;
}

.footer{
    position:fixed;
    bottom:10px;
    left:0;
    right:0;
    text-align:center;
    color:gray;
    font-size:14px;
}

.stButton>button{
    background-color:#1f4e79;
    color:white;
    border-radius:8px;
    padding:10px 20px;
}
</style>
""", unsafe_allow_html=True)

# ------------------ HEADER ------------------

st.markdown('<p class="main-title">SOP Formatter</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">by S M Baqir</p>', unsafe_allow_html=True)

st.write("Upload a Word SOP document to automatically format fonts, spacing, and alignment.")

# ------------------ INPUT OPTIONS ------------------

uploaded_file = st.file_uploader("Upload SOP Document (.docx)", type=["docx"])

col1, col2 = st.columns(2)

with col1:
    first_page_size = st.number_input(
        "First Page Font Size",
        min_value=8,
        max_value=40,
        value=16
    )

with col2:
    other_page_size = st.number_input(
        "Other Pages Font Size",
        min_value=8,
        max_value=30,
        value=11
    )

line_spacing = st.slider(
    "Line Spacing",
    1.0,
    3.0,
    1.5
)

# ------------------ PROCESS FUNCTION ------------------

def format_document(file):

    doc = Document(file)

    for i, para in enumerate(doc.paragraphs):

        # Justify alignment
        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Line spacing
        para.paragraph_format.line_spacing = line_spacing

        for run in para.runs:

            # Preserve bold automatically
            if i < 10:
                run.font.size = Pt(first_page_size)
            else:
                run.font.size = Pt(other_page_size)

    # Format tables as well
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    para.paragraph_format.line_spacing = line_spacing

                    for run in para.runs:
                        run.font.size = Pt(other_page_size)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# ------------------ PROCESS BUTTON ------------------

if uploaded_file:

    if st.button("✨ Format SOP Document"):

        output = format_document(uploaded_file)

        st.success("Document formatted successfully!")

        st.download_button(
            label="⬇ Download Formatted SOP",
            data=output,
            file_name="Formatted_SOP.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ------------------ FOOTER ------------------

st.markdown(
    '<div class="footer">OMAC Developer</div>',
    unsafe_allow_html=True
)
