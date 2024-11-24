import streamlit as st
from pptx import Presentation
from PyPDF2 import PdfReader
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import io

import spacy
# Load spaCy language model
nlp_spacy = spacy.load("en_core_web_sm")


# Streamlit App
st.title("RFP to PowerPoint Generator with NLP")
st.write("Upload an RFP PDF, and we'll extract key sections, summarize content, and generate a professional PowerPoint.")

# Step 1: Upload PDF
uploaded_file = st.file_uploader("Upload an RFP PDF", type=["pdf"])

if uploaded_file:
    st.write("**Uploaded RFP Content:**")
    rfp_text = extract_text_from_pdf(uploaded_file)
    st.text_area("Extracted RFP Text", value=rfp_text, height=300)

    # Step 2: Use NLP to Extract Key Details
    st.write("**Extracted Key Details Using NLP:**")
    key_details = extract_key_details(rfp_text)
    for section, content in key_details.items():
        st.write(f"**{section}:** {content}")

    # Step 3: Summarize Long Sections
    st.write("**Summarized Sections Using NLP:**")
    summarized_goals = summarize_text(key_details["Goals"])
    summarized_deliverables = summarize_text(key_details["Deliverables"])
    summarized_timeline = summarize_text(key_details["Timeline"])
    summarized_criteria = summarize_text(key_details["Evaluation Criteria"])

    st.write("**Summarized Goals:**", summarized_goals)
    st.write("**Summarized Deliverables:**", summarized_deliverables)
    st.write("**Summarized Timeline:**", summarized_timeline)
    st.write("**Summarized Evaluation Criteria:**", summarized_criteria)

    # Step 4: Generate PowerPoint Presentation
    if st.button("Generate PowerPoint"):
        prs = Presentation()
        prs.slide_width = Inches(8.5)
        prs.slide_height = Inches(11)

        # Title slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(7.5), Inches(2))
        title_frame = title_box.text_frame
        p = title_frame.add_paragraph()
        p.text = "RFP Response Presentation"
        p.font.name = "Helvetica"
        p.font.size = Pt(36)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)

        # Goals slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(7.5), Inches(1))
        title_frame = title_box.text_frame
        p = title_frame.add_paragraph()
        p.text = "Goals"
        p.font.name = "Helvetica"
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)

        content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(7.5), Inches(3))
        content_frame = content_box.text_frame
        p = content_frame.add_paragraph()
        p.text = summarized_goals
        p.font.name = "Helvetica"
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(50, 50, 50)

        # Save presentation to buffer
        buffer = io.BytesIO()
        prs.save(buffer)
        buffer.seek(0)

        # Step 5: Download PowerPoint
        st.download_button(
            label="Download PowerPoint",
            data=buffer,
            file_name="Generated_RFP_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
