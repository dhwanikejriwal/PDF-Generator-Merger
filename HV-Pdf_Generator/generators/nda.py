import streamlit as st
from datetime import datetime
import os
from docx import Document

from pdf_utils import convert_to_pdf
from session_manager import clear_session_keys

# ========== Helper Functions ==========

def replace_text_in_paragraph(paragraph, placeholders):
    """Replace placeholders in a paragraph."""
    full_text = ''.join(run.text for run in paragraph.runs)

    for key, value in placeholders.items():
        if key in full_text:
            before, sep, after = full_text.partition(key)
            for run in list(paragraph.runs)[::-1]:
                paragraph._p.remove(run._r)
            if before:
                paragraph.add_run(before)
            new_run = paragraph.add_run(value)
            if after:
                paragraph.add_run(after)
            return  # Done replacing this paragraph

def replace_text_in_table(table, placeholders):
    """Replace placeholders in all cells of a table."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, placeholders)

def edit_nda_template(template_path, output_path, placeholders):
    """Edit the NDA template and save the filled version."""
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, placeholders)

    for table in doc.tables:
        replace_text_in_table(table, placeholders)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path

# ========== Main Generator Function ==========

def generate_nda():
    st.title("NDA Generator")

    client_name = st.text_input("Enter Client Name:")
    company_name = st.text_input("Enter Company Name:")
    date_input = st.date_input("Enter Date:", datetime.today())
    address = st.text_area("Enter Address:")

    placeholders = {
        "<<ClientName>>": client_name,
        "<<CompanyName>>": company_name,
        "<<Date>>": date_input.strftime("%d-%m-%Y"),
        "<<Address>>": address,
    }

    template_name = "NDA Template.docx"
    output_dir = os.path.join("app", "generated_files", "nda")
    os.makedirs(output_dir, exist_ok=True)

    if st.button("Generate NDA"):
        try:
            clear_session_keys(["nda_docx", "nda_pdf", "nda_docx_name", "nda_pdf_name"])

            template_path = os.path.join(os.getcwd(), template_name)

            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            safe_name = ''.join(c if c.isalnum() else '_' for c in client_name)

            docx_output_path = os.path.join(output_dir, f"NDA_{safe_name}.docx")
            pdf_output_path = os.path.join(output_dir, f"NDA_{safe_name}.pdf")

            # Generate DOCX
            edit_nda_template(template_path, docx_output_path, placeholders)

            # Save DOCX to session
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.nda_docx = docx_file.read()
                st.session_state.nda_docx_name = f"NDA_{safe_name}.docx"

            # Convert to PDF
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)

                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.nda_pdf = pdf_file.read()
                        st.session_state.nda_pdf_name = f"NDA_{safe_name}.pdf"
                else:
                    st.warning("PDF not found after conversion.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                st.warning("PDF conversion failed, but DOCX is available.")

            # Download buttons
            col1, col2 = st.columns(2)

            with col1:
                if st.session_state.nda_docx:
                    st.download_button(
                        label="📥 Download NDA (Word)",
                        data=st.session_state.nda_docx,
                        file_name=st.session_state.nda_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            with col2:
                if st.session_state.nda_pdf:
                    st.download_button(
                        label="📥 Download NDA (PDF)",
                        data=st.session_state.nda_pdf,
                        file_name=st.session_state.nda_pdf_name,
                        mime="application/pdf"
                    )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())
