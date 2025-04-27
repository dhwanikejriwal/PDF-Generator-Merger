import streamlit as st
from datetime import datetime
import os
from docx import Document

from pdf_utils import convert_to_pdf
from session_manager import clear_session_keys

# ========== Helper Functions ==========

def replace_text_in_paragraph(paragraph, placeholders):
    """Replace placeholders in a paragraph."""
    if not paragraph.runs:
        return

    full_text = ''.join(run.text for run in paragraph.runs)

    for key, value in placeholders.items():
        if key in full_text:
            full_text = full_text.replace(key, value)

    # Remove all runs except first
    for i in range(len(paragraph.runs) - 1, 0, -1):
        paragraph._p.remove(paragraph.runs[i]._r)

    if paragraph.runs:
        paragraph.runs[0].text = full_text

def replace_text_in_table(table, placeholders):
    """Replace placeholders inside all cells of a table."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, placeholders)

def edit_hiring_template(template_path, output_path, placeholders):
    """Edit hiring contract template and save filled version."""
    doc = Document(template_path)

    for para in doc.paragraphs:
        replace_text_in_paragraph(para, placeholders)

    for table in doc.tables:
        replace_text_in_table(table, placeholders)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path

# ========== Main Generator Function ==========

def generate_hiring_contract():
    st.title("Hiring Contract Generator")

    employee_name = st.text_input("Enter Employee Name:")
    role = st.text_input("Enter Role:")
    date = st.date_input("Enter Today's Date:", datetime.today())
    starting_date = st.date_input("Enter Starting Date:")
    stipend = st.number_input("Enter the Stipend Amount:")
    working_hours = st.number_input("Enter Total Working Hours:")
    internship_duration = st.number_input("Enter Internship Duration (in months):")
    first_pay_date = st.date_input("Enter First Pay Cheque Date:")

    placeholders = {
        "<<Date>>": date.strftime("%d %B, %Y"),
        "<<Name>>": employee_name,
        "<<Role>>": role,
        "<<Starting Date>>": starting_date.strftime("%d %B, %Y"),
        "<<Internship Duration>>": str(int(internship_duration)),
        "<<First Pay>>": first_pay_date.strftime("%d %B, %Y"),
        "<<Stipend>>": str(stipend),
        "<<Working Hours>>": str(int(working_hours)),
    }

    template_name = "Hiring Contract.docx"
    output_dir = os.path.join("app", "generated_files", "hiring")
    os.makedirs(output_dir, exist_ok=True)

    if st.button("Generate Hiring Contract"):
        try:
            clear_session_keys(["hiring_docx", "hiring_pdf", "hiring_docx_name", "hiring_pdf_name"])

            template_path = os.path.join(os.getcwd(), template_name)

            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            safe_name = ''.join(c if c.isalnum() else '_' for c in employee_name)

            docx_output_path = os.path.join(output_dir, f"Hiring_{safe_name}.docx")
            pdf_output_path = os.path.join(output_dir, f"Hiring_{safe_name}.pdf")

            # Generate DOCX
            edit_hiring_template(template_path, docx_output_path, placeholders)

            # Save DOCX to session
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.hiring_docx = docx_file.read()
                st.session_state.hiring_docx_name = f"Hiring_{safe_name}.docx"

            # Convert to PDF
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)

                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.hiring_pdf = pdf_file.read()
                        st.session_state.hiring_pdf_name = f"Hiring_{safe_name}.pdf"
                else:
                    st.warning("PDF not found after conversion.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                st.warning("PDF conversion failed, but DOCX is available.")

            # Download buttons
            col1, col2 = st.columns(2)

            with col1:
                if st.session_state.hiring_docx:
                    st.download_button(
                        label="📥 Download Hiring Contract (Word)",
                        data=st.session_state.hiring_docx,
                        file_name=st.session_state.hiring_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            with col2:
                if st.session_state.hiring_pdf:
                    st.download_button(
                        label="📥 Download Hiring Contract (PDF)",
                        data=st.session_state.hiring_pdf,
                        file_name=st.session_state.hiring_pdf_name,
                        mime="application/pdf"
                    )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())
