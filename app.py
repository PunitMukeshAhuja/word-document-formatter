
import io
import re
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(
    page_title="Word Document Formatter",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Word Document Formatter")
st.caption(
    "Upload a Word document, preview detected structure, correct it manually, "
    "apply formatting, and download the cleaned file."
)

CLASSIFICATION_OPTIONS = ["Title", "Heading 1", "Heading 2", "Heading 3", "Body"]

# -----------------------------
# Helpers
# -----------------------------

def paragraph_text(paragraph):
    return paragraph.text.strip()

def average_font_size(paragraph):
    sizes = []
    for run in paragraph.runs:
        if run.font.size:
            sizes.append(run.font.size.pt)
    return sum(sizes) / len(sizes) if sizes else None

def mostly_bold(paragraph):
    runs = [r for r in paragraph.runs if r.text.strip()]
    if not runs:
        return False
    bold_runs = sum(1 for r in runs if r.bold is True)
    return bold_runs / len(runs) >= 0.6

def is_all_caps(text):
    letters = [c for c in text if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters)

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

def detect_heading_level(paragraph, auto_detect, body_size):
    text = paragraph_text(paragraph)

    if not text:
        return None

    style_name = paragraph.style.name if paragraph.style else ""

    if style_name == "Title":
        return 0
    if style_name == "Heading 1":
        return 1
    if style_name == "Heading 2":
        return 2
    if style_name == "Heading 3":
        return 3

    if not auto_detect:
        return None

    if len(text) > 120:
        return None

    numbered = re.match(r"^\s*(\d+(?:\.\d+){0,2})[\.\)]?\s+\S+", text)
    if numbered:
        number_part = numbered.group(1)
        depth = number_part.count(".") + 1
        return min(depth, 3)

    avg_size = average_font_size(paragraph)
    bold = mostly_bold(paragraph)
    caps = is_all_caps(text)

    if len(text) <= 70 and caps and bold and avg_size and avg_size >= body_size + 4:
        return 0

    if len(text) <= 60 and caps and (bold or (avg_size and avg_size >= body_size + 2)):
        return 1

    if len(text) <= 70 and bold:
        if avg_size and avg_size >= body_size + 3:
            return 1
        if avg_size and avg_size >= body_size + 1:
            return 2
        return 2

    if len(text) <= 80 and avg_size:
        if avg_size >= body_size + 4:
            return 1
        if avg_size >= body_size + 2:
            return 2
        if avg_size >= body_size + 1:
            return 3

    return None

def classification_label(level):
    if level == 0:
        return "Title"
    if level == 1:
        return "Heading 1"
    if level == 2:
        return "Heading 2"
    if level == 3:
        return "Heading 3"
    return "Body"

def build_preview(file_bytes, auto_detect, body_size):
    doc = Document(io.BytesIO(file_bytes))
    rows = []

    for idx, paragraph in enumerate(doc.paragraphs, start=1):
        text = paragraph_text(paragraph)
        if not text:
            continue

        level = detect_heading_level(
            paragraph=paragraph,
            auto_detect=auto_detect,
            body_size=body_size
        )

        rows.append({
            "Paragraph": idx,
            "Detected as": classification_label(level),
            "Apply as": classification_label(level),
            "Text preview": text[:160] + ("..." if len(text) > 160 else ""),
            "Original style": paragraph.style.name if paragraph.style else "",
            "Avg font size": round(average_font_size(paragraph), 1)
            if average_font_size(paragraph) else None,
            "Mostly bold": "Yes" if mostly_bold(paragraph) else "No",
        })

    return pd.DataFrame(rows)

def format_table(table, font_name, body_size):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.line_spacing = 1.0
                paragraph.paragraph_format.space_after = Pt(0)

                for run in paragraph.runs:
                    set_run_font(run, font_name, body_size)

def format_document(
    file_bytes,
    font_name,
    body_size,
    line_spacing,
    alignment,
    space_after,
    margin,
    title_size,
    heading1_size,
    heading2_size,
    heading3_size,
    manual_map,
    add_page_numbers
):
    doc = Document(io.BytesIO(file_bytes))

    for section in doc.sections:
        section.top_margin = Inches(margin)
        section.bottom_margin = Inches(margin)
        section.left_margin = Inches(margin)
        section.right_margin = Inches(margin)

    for idx, paragraph in enumerate(doc.paragraphs, start=1):
        if not paragraph_text(paragraph):
            continue

        classification = manual_map.get(idx, "Body")

        if classification == "Title":
            size = title_size
            bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(12)

        elif classification == "Heading 1":
            size = heading1_size
            bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_before = Pt(10)
            paragraph.paragraph_format.space_after = Pt(6)

        elif classification == "Heading 2":
            size = heading2_size
            bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_before = Pt(8)
            paragraph.paragraph_format.space_after = Pt(4)

        elif classification == "Heading 3":
            size = heading3_size
            bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_before = Pt(6)
            paragraph.paragraph_format.space_after = Pt(3)

        else:
            size = body_size
            bold = None
            paragraph.alignment = alignment
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(space_after)

        paragraph.paragraph_format.line_spacing = line_spacing

        for run in paragraph.runs:
            set_run_font(run, font_name, size, bold)

    for table in doc.tables:
        format_table(table, font_name, body_size)

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

# -----------------------------
# UI
# -----------------------------

uploaded_file = st.file_uploader(
    "Upload a Word document",
    type=["docx"],
    help="Only .docx files are supported in this version."
)

if uploaded_file:
    file_bytes = uploaded_file.getvalue()

    if st.session_state.get("current_file_name") != uploaded_file.name:
        st.session_state["current_file_name"] = uploaded_file.name
        st.session_state.pop("edited_preview", None)
        st.session_state.pop("formatted_doc", None)
        st.session_state.pop("output_name", None)

    st.success(f"Loaded: {uploaded_file.name}")

    left, right = st.columns([1, 1])

    with left:
        st.subheader("1. Formatting")

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

    with right:
        st.subheader("2. Heading detection")

        auto_detect_headings = st.checkbox(
            "Auto-detect headings",
            value=True,
            help=(
                "Tries to detect headings using Word styles, numbering, "
                "bold text, capitalization, font size and paragraph length."
            )
        )

        title_size = st.number_input("Title size", 12, 36, 20)
        heading1_size = st.number_input("Heading 1 size", 10, 30, 16)
        heading2_size = st.number_input("Heading 2 size", 10, 26, 14)
        heading3_size = st.number_input("Heading 3 size", 10, 24, 13)

        add_page_numbers = st.checkbox("Add page numbers", value=True)

    st.divider()

    st.subheader("3. Review & correct structure")

    auto_preview_df = build_preview(
        file_bytes=file_bytes,
        auto_detect=auto_detect_headings,
        body_size=body_size
    )

    # Reset edited preview if structure/settings change materially.
    preview_signature = (
        uploaded_file.name,
        uploaded_file.size,
        auto_detect_headings,
        body_size
    )

    if st.session_state.get("preview_signature") != preview_signature:
        st.session_state["edited_preview"] = auto_preview_df.copy()
        st.session_state["preview_signature"] = preview_signature

    st.write(
        "You can change **Apply as** for any paragraph before formatting. "
        "The original automatic detection remains visible for comparison."
    )

    edited_df = st.data_editor(
        st.session_state["edited_preview"],
        use_container_width=True,
        hide_index=True,
        disabled=[
            "Paragraph",
            "Detected as",
            "Text preview",
            "Original style",
            "Avg font size",
            "Mostly bold"
        ],
        column_config={
            "Apply as": st.column_config.SelectboxColumn(
                "Apply as",
                help="Manually choose how this paragraph should be formatted.",
                options=CLASSIFICATION_OPTIONS,
                required=True
            ),
            "Text preview": st.column_config.TextColumn(
                "Text preview",
                width="large"
            )
        },
        key="structure_editor"
    )

    st.session_state["edited_preview"] = edited_df.copy()

    if not edited_df.empty:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Title", int((edited_df["Apply as"] == "Title").sum()))
        c2.metric("Heading 1", int((edited_df["Apply as"] == "Heading 1").sum()))
        c3.metric("Heading 2", int((edited_df["Apply as"] == "Heading 2").sum()))
        c4.metric("Heading 3", int((edited_df["Apply as"] == "Heading 3").sum()))
        c5.metric("Body", int((edited_df["Apply as"] == "Body").sum()))

    b1, b2 = st.columns([1, 1])

    with b1:
        if st.button("↩️ Reset to auto-detection", use_container_width=True):
            st.session_state["edited_preview"] = auto_preview_df.copy()
            st.rerun()

    with b2:
        st.caption(
            "Tip: correct only the paragraphs that look wrong; the rest can stay as detected."
        )

    st.divider()
    st.subheader("4. Format & download")

    if st.button("✨ Format document", type="primary", use_container_width=True):
        try:
            manual_map = {
                int(row["Paragraph"]): row["Apply as"]
                for _, row in st.session_state["edited_preview"].iterrows()
            }

            result = format_document(
                file_bytes=file_bytes,
                font_name=font_name,
                body_size=body_size,
                line_spacing=line_spacing,
                alignment=alignment_map[alignment_label],
                space_after=space_after,
                margin=margin,
                title_size=title_size,
                heading1_size=heading1_size,
                heading2_size=heading2_size,
                heading3_size=heading3_size,
                manual_map=manual_map,
                add_page_numbers=add_page_numbers
            )

            st.session_state["formatted_doc"] = result.getvalue()
            st.session_state["output_name"] = uploaded_file.name.replace(
                ".docx", "_formatted.docx"
            )

            st.success("Formatting complete.")

        except Exception as exc:
            st.error(f"Could not format the document: {exc}")

    if "formatted_doc" in st.session_state:
        st.download_button(
            "⬇️ Download formatted document",
            data=st.session_state["formatted_doc"],
            file_name=st.session_state["output_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

else:
    st.info(
        "Upload a .docx file to begin. You can preview and manually correct "
        "the detected structure before formatting."
    )

st.divider()
st.caption(
    "Automatic heading detection is heuristic-based. Manual corrections in the preview are applied to the downloaded document."
)
