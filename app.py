
import io
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(
    page_title="Word Document Formatter",
    page_icon="📝",
    layout="centered"
)

st.title("📝 Word Document Formatter")
st.write(
    "Upload a Word document, apply consistent formatting, "
    "and download the cleaned document."
)

uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

st.subheader("Formatting settings")

font_name = st.selectbox(
    "Body font",
    ["Times New Roman", "Arial", "Calibri", "Aptos", "Georgia"],
    index=0
)

body_size = st.number_input(
    "Body font size",
    min_value=8,
    max_value=20,
    value=12
)

line_spacing = st.selectbox(
    "Line spacing",
    [1.0, 1.15, 1.5, 2.0],
    index=2
)

alignment_label = st.selectbox(
    "Body alignment",
    ["Justified", "Left", "Center", "Right"],
    index=0
)

alignment_map = {
    "Justified": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "Left": WD_ALIGN_PARAGRAPH.LEFT,
    "Center": WD_ALIGN_PARAGRAPH.CENTER,
    "Right": WD_ALIGN_PARAGRAPH.RIGHT
}

space_after = st.number_input(
    "Space after paragraph (pt)",
    min_value=0,
    max_value=30,
    value=6
)

margin = st.number_input(
    "Margins (inches)",
    min_value=0.3,
    max_value=2.0,
    value=1.0,
    step=0.1
)

st.subheader("Heading settings")

heading1_size = st.number_input("Heading 1 size", 10, 30, 16)
heading2_size = st.number_input("Heading 2 size", 10, 26, 14)
heading3_size = st.number_input("Heading 3 size", 10, 24, 13)

add_page_numbers = st.checkbox("Add page numbers", value=True)

def set_run_font(run, name, size, bold=None):
    run.font.name = name
    run.font.size = Pt(size)

    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts

    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)

    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:eastAsia"), name)

    if bold is not None:
        run.bold = bold

def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run()

    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")

    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"

    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")

    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)

def format_paragraph(paragraph):
    style_name = paragraph.style.name if paragraph.style else ""

    if style_name == "Heading 1":
        size = heading1_size
        bold = True
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    elif style_name == "Heading 2":
        size = heading2_size
        bold = True
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    elif style_name == "Heading 3":
        size = heading3_size
        bold = True
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    else:
        size = body_size
        bold = None
        paragraph.alignment = alignment_map[alignment_label]

    paragraph.paragraph_format.line_spacing = line_spacing
    paragraph.paragraph_format.space_after = Pt(space_after)

    for run in paragraph.runs:
        set_run_font(run, font_name, size, bold)

def format_table(table):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.space_after = Pt(0)

                for run in paragraph.runs:
                    set_run_font(run, font_name, body_size)

def format_document(file_bytes):
    doc = Document(io.BytesIO(file_bytes))

    for section in doc.sections:
        section.top_margin = Inches(margin)
        section.bottom_margin = Inches(margin)
        section.left_margin = Inches(margin)
        section.right_margin = Inches(margin)

    for paragraph in doc.paragraphs:
        format_paragraph(paragraph)

    for table in doc.tables:
        format_table(table)

    if add_page_numbers:
        for section in doc.sections:
            footer = section.footer
            paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            paragraph.clear()
            add_page_number(paragraph)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

if uploaded_file:
    st.success(f"Loaded: {uploaded_file.name}")

    if st.button("Format document", type="primary"):
        try:
            result = format_document(uploaded_file.getvalue())

            st.success("Formatting complete.")

            output_name = uploaded_file.name.replace(
                ".docx", "_formatted.docx"
            )

            st.download_button(
                "Download formatted document",
                data=result,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as exc:
            st.error(f"Could not format the document: {exc}")

st.divider()
st.caption(
    "Tip: For best results, use Word Heading 1 / Heading 2 / Heading 3 styles "
    "in the source document."
)
