import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ---------------- PAGE CONFIG ----------------

st.set_page_config(
    page_title="SOP Formatter by S M Baqir",
    page_icon="📄",
    layout="centered"
)

# ---------------- STYLE ----------------

st.markdown("""
<style>
.title{
    font-size:36px;
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
    width:100%;
    text-align:center;
    color:gray;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------

st.markdown('<p class="title">SOP Formatter</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">by S M Baqir</p>', unsafe_allow_html=True)

# ---------------- STEP 1 : UPLOAD ----------------

uploaded_file = st.file_uploader("📂 Upload SOP Document (.docx)", type=["docx"])

# ---------------- STEP 2 : SHOW OPTIONS AFTER UPLOAD ----------------

if uploaded_file:

    st.subheader("Formatting Settings")

    first_page_size = st.number_input(
        "First Page Font Size",
        min_value=8,
        max_value=40,
        value=16
    )

    other_page_size = st.number_input(
        "2nd Page Onwards Font Size",
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

    # ---------------- PROCESS FUNCTION ----------------

    def format_document(file):

        doc = Document(file)

        for i, para in enumerate(doc.paragraphs):

            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            para.paragraph_format.line_spacing = line_spacing

            for run in para.runs:

                if i < 10:
                    run.font.size = Pt(first_page_size)
                else:
                    run.font.size = Pt(other_page_size)

        # format tables
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

    # ---------------- STEP 3 : PROCESS ----------------

    if st.button("✨ Format SOP Document"):

        output = format_document(uploaded_file)

        st.success("Document formatted successfully!")

        st.download_button(
            "⬇ Download Formatted SOP",
            data=output,
            file_name="Formatted_SOP.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ---------------- FOOTER ----------------

st.markdown(
    '<div class="footer">OMAC Developer</div>',
    unsafe_allow_html=True
)
