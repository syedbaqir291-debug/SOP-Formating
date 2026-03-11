import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import tempfile

st.title("Word Document Formatter")

uploaded_file = st.file_uploader("Upload Word File (.docx)", type=["docx"])

first_page_size = st.number_input(
    "First Page Font Size",
    value=16
)

other_page_size = st.number_input(
    "2nd Page Onwards Font Size",
    value=11
)

line_spacing = st.number_input(
    "Line Spacing",
    value=1.5,
    step=0.1
)

if uploaded_file:

    if st.button("Process Document"):

        doc = Document(uploaded_file)

        for i, para in enumerate(doc.paragraphs):

            # Set alignment to justified
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Set line spacing
            para_format = para.paragraph_format
            para_format.line_spacing = line_spacing

            for run in para.runs:

                # preserve bold
                if run.bold:
                    run.bold = True

                # First page larger font
                if i < 10:   # approximation for first page
                    run.font.size = Pt(first_page_size)
                else:
                    run.font.size = Pt(other_page_size)

        # Save temporary file
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.download_button(
                "Download Formatted Document",
                f,
                file_name="formatted_document.docx"
            )

        st.success("Document formatted successfully!")
