import streamlit as st
from docx import Document
from datetime import datetime
import os
import platform
import subprocess
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
import fitz
from docx.oxml.ns import qn
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import tempfile
import uuid
import logging
import sys
from num2words import num2words


port = int(os.environ.get("PORT", 8080))  # Default to 8080
st.set_page_config(page_title="PDF Generator")
st.title("PDF Generator App")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger("pdf_generator")

def convert_to_pdf(doc_path, pdf_path):
    """Convert Word document to PDF with better error handling."""
    doc_path = os.path.abspath(doc_path)
    pdf_path = os.path.abspath(pdf_path)

    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Word document not found at {doc_path}")

    logger.info(f"Converting document: {doc_path} to PDF: {pdf_path}")

    # Use a temporary directory for the intermediate PDF file
    with tempfile.TemporaryDirectory() as temp_dir:
        # Get just the filename without path
        doc_filename = os.path.basename(doc_path)
        expected_pdf_name = os.path.splitext(doc_filename)[0] + ".pdf"
        temp_pdf_path = os.path.join(temp_dir, expected_pdf_name)

        # Step 1: Convert Word to PDF
        if platform.system() == "Windows":
            try:
                import comtypes.client
                import pythoncom
                pythoncom.CoInitialize()
                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(doc_path)
                doc.SaveAs(temp_pdf_path, FileFormat=17)  # FileFormat=17 is for PDF
                doc.Close()
                word.Quit()
                logger.info("Windows COM conversion completed")
            except Exception as e:
                logger.error(f"Error using COM on Windows: {e}", exc_info=True)
                raise Exception(f"Error using COM on Windows: {e}")
        else:
            try:
                # Run LibreOffice conversion
                logger.info(f"Running LibreOffice conversion to {temp_dir}")
                result = subprocess.run(
                    ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', temp_dir, doc_path],
                    check=True,
                    capture_output=True,
                    text=True
                )
                logger.info(f"LibreOffice stdout: {result.stdout}")
                
                # Check if the expected PDF file exists in the temp directory
                if not os.path.exists(temp_pdf_path):
                    logger.warning(f"Expected PDF not found at {temp_pdf_path}, searching for alternatives...")
                    # Try to find any PDF file that was created
                    pdf_files = [f for f in os.listdir(temp_dir) if f.endswith('.pdf')]
                    if pdf_files:
                        temp_pdf_path = os.path.join(temp_dir, pdf_files[0])
                        logger.info(f"Found alternative PDF: {temp_pdf_path}")
                    else:
                        # If no PDF was found, try using direct path instead
                        direct_pdf_path = os.path.splitext(doc_path)[0] + '.pdf'
                        if os.path.exists(direct_pdf_path):
                            temp_pdf_path = direct_pdf_path
                            logger.info(f"Using direct PDF path: {temp_pdf_path}")
                        else:
                            # Try alternative conversion
                            try:
                                logger.info("Attempting alternative PDF conversion using docx2pdf...")
                                alternative_convert_to_pdf(doc_path, temp_pdf_path)
                                if os.path.exists(temp_pdf_path):
                                    logger.info(f"Alternative conversion successful: {temp_pdf_path}")
                                else:
                                    raise FileNotFoundError("Alternative conversion did not produce a PDF file")
                            except Exception as alt_err:
                                logger.error(f"Alternative conversion failed: {alt_err}", exc_info=True)
                                raise FileNotFoundError(f"PDF conversion failed. No PDF file was created. LibreOffice output: {result.stdout}\nError: {result.stderr}")
            except subprocess.CalledProcessError as e:
                logger.error(f"LibreOffice conversion failed: {e}", exc_info=True)
                # Try alternative method - use unoconv if available
                try:
                    logger.info("Attempting unoconv conversion...")
                    subprocess.run(['unoconv', '-f', 'pdf', '-o', temp_dir, doc_path], check=True)
                    # Check if PDF was created
                    pdf_files = [f for f in os.listdir(temp_dir) if f.endswith('.pdf')]
                    if pdf_files:
                        temp_pdf_path = os.path.join(temp_dir, pdf_files[0])
                        logger.info(f"Unoconv created PDF: {temp_pdf_path}")
                    else:
                        # Try alternative conversion
                        logger.info("Unoconv didn't create a PDF, trying alternative method...")
                        alternative_convert_to_pdf(doc_path, temp_pdf_path)
                        if not os.path.exists(temp_pdf_path):
                            raise Exception("PDF conversion failed with all methods")
                except Exception as uno_err:
                    logger.error(f"Unoconv conversion failed: {uno_err}", exc_info=True)
                    # Last resort - try direct alternative conversion
                    try:
                        logger.info("Last resort - direct alternative conversion...")
                        alternative_convert_to_pdf(doc_path, pdf_path)
                        # If successful, early return
                        if os.path.exists(pdf_path):
                            logger.info(f"Direct alternative conversion successful: {pdf_path}")
                            return
                        else:
                            raise Exception(f"Error using LibreOffice: {e}. Unoconv fallback also failed: {uno_err}")
                    except Exception as final_err:
                        logger.error(f"All conversion methods failed: {final_err}", exc_info=True)
                        raise Exception(f"All PDF conversion methods failed: {final_err}")

        # Verify the PDF exists before proceeding
        if not os.path.exists(temp_pdf_path):
            raise FileNotFoundError(f"Expected PDF file not found at {temp_pdf_path}")
            
        # Step 2: Flatten the PDF or just copy it if flattening is not needed
        try:
            logger.info(f"Flattening PDF from {temp_pdf_path} to {pdf_path}")
            flatten_pdf(temp_pdf_path, pdf_path)
        except Exception as flatten_err:
            logger.warning(f"PDF flattening failed: {flatten_err}. Using direct copy instead.", exc_info=True)
            # If flattening fails, just copy the PDF
            import shutil
            shutil.copy2(temp_pdf_path, pdf_path)
            logger.info(f"Direct copy completed to {pdf_path}")

def alternative_convert_to_pdf(doc_path, pdf_path):
    """Alternative PDF conversion using reportlab and python-docx."""
    try:
        from docx2pdf import convert
        convert(doc_path, pdf_path)
    except ImportError:
        try:
            # Try a different approach using reportlab and python-docx
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
            from docx import Document
            
            # Extract text from DOCX
            doc = Document(doc_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            
            # Generate simple PDF
            pdf = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            content = [Paragraph(text, styles["Normal"]) for text in full_text if text.strip()]
            pdf.build(content)
        except Exception as e:
            raise Exception(f"Failed to convert DOCX to PDF using alternative methods: {e}")

def flatten_pdf(input_pdf_path, output_pdf_path):
    """
    Converts each page of a PDF into an image and re-embeds it to create a flattened, non-editable PDF.
    """
    if not os.path.exists(input_pdf_path):
        raise FileNotFoundError(f"Input PDF file not found: {input_pdf_path}")

    try:
        doc = fitz.open(input_pdf_path)  # Open the original PDF
        writer = PdfWriter()

        with tempfile.TemporaryDirectory() as temp_dir:
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(dpi=300)  # Render page to an image with 300 DPI
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Save the image as a temporary PDF
                temp_page_path = os.path.join(temp_dir, f"temp_page_{page_num}.pdf")
                image.save(temp_page_path, "PDF")

                # Read the temporary PDF and add it to the writer
                reader = PdfReader(temp_page_path)
                writer.add_page(reader.pages[0])

        # Save the flattened PDF
        with open(output_pdf_path, "wb") as f:
            writer.write(f)

        logger.info(f"Flattened PDF saved at: {output_pdf_path}")
    except Exception as e:
        logger.error(f"Error in PDF flattening: {e}", exc_info=True)
        # If flattening fails, allow the original function to handle the exception
        raise

# Common Functions (unchanged)
def apply_formatting(run, font_name, font_size, bold=False):
    """Apply specific formatting to a run."""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(font_size)
    run.bold = bold

# def replace_and_format(doc, placeholders, font_name, font_size, option):
#     """Replace placeholders and apply formatting."""
#     for para in doc.paragraphs:
#         if para.text:
#             for key, value in placeholders.items():
#                 if key in para.text:
#                     runs = para.runs
#                     for run in runs:
#                         if key in run.text:
#                             run.text = run.text.replace(key, value)
#                             if para == doc.paragraphs[0]:
#                                 apply_formatting(run, font_name, font_size, bold=True)
#                         else:
#                             run.text = run.text.replace(key, value)

#     for table in doc.tables:
#         for row in table.rows:
#             for cell in row.cells:
#                 if cell.text.strip():
#                     for key, value in placeholders.items():
#                         if key in cell.text:
#                             cell.text = cell.text.replace(key, value)
#                             for paragraph in cell.paragraphs:
#                                 paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if key == "<<Address>>" else WD_ALIGN_PARAGRAPH.CENTER
#                                 for run in paragraph.runs:
#                                     apply_formatting(run, "Times New Roman", 11 if option == "NDA" else 12)
#                             cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# def edit_word_template(template_path, output_path, placeholders, font_name, font_size, option):
#     """Edit Word document and apply formatting."""
#     try:
#         doc = Document(template_path)
#         replace_and_format(doc, placeholders, font_name, font_size, option)
#         doc.save(output_path)
#         return output_path
#     except Exception as e:
#         logger.error(f"Error editing Word template: {e}", exc_info=True)
#         raise Exception(f"Error editing Word template: {e}")

def apply_image_placeholder(doc, placeholder_key, image_file):
    """Replace a placeholder with an image in the Word document."""
    try:
        placeholder_found = False

        # Check inside tables first
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if placeholder_key in para.text:
                            para.clear()  # Clears text while preserving formatting
                            run = para.add_run()
                            run.add_picture(image_file, width=Inches(1.5), height=Inches(0.75))
                            placeholder_found = True

        # Check paragraphs outside tables
        for para in doc.paragraphs:
            if placeholder_key in para.text:
                para.clear()
                run = para.add_run()
                run.add_picture(image_file, width=Inches(1.2), height=Inches(0.75))
                placeholder_found = True

        if not placeholder_found:
            logger.warning(f"Placeholder '{placeholder_key}' not found in the document.")
        
        return doc

    except Exception as e:
        logger.error(f"Error inserting image: {e}", exc_info=True)
        return None  # Returning None to indicate failure

# Contract Generator
def replace_placeholders(doc, placeholders):
    """Replace placeholders in paragraphs and tables."""
    for para in doc.paragraphs:
        for key, value in placeholders.items():
            if key in para.text:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)
                        # Apply bold formatting for specific placeholders
                        if key == "<<EndDate>>":
                            run.bold = True  # Apply bold formatting

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)
                                    # Apply bold formatting for specific placeholders
                                    if key == "<<EndDate>>":
                                        run.bold = True  # Apply bold formatting
    return doc


def replace_text_in_table(table, placeholders):
    """Iterate over all cells and paragraphs in a table."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_placeholders(paragraph, placeholders)


def edit_contract_template(template_path, output_path, placeholders):
    """Load a Word template, replace placeholders, and save to output."""
    doc = Document(template_path)
    # Replace in document body
    for para in doc.paragraphs:
        replace_placeholders(para, placeholders)
    # Replace in all tables
    for table in doc.tables:
        replace_text_in_table(table, placeholders)
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path


def generate_contract():
    """Streamlit app to gather inputs and generate the contract document."""
    st.title("Contract Generator")

    # User inputs
    client_name = st.text_input("Enter Client Name:")
    company_name = st.text_input("Enter Company Name:")
    date_input = st.date_input("Enter Effective Date:", datetime.today())
    end_date_input = st.date_input("Enter Contract End Date:", datetime.today())
    address = st.text_area("Enter Address:")

    placeholders = {
        "<<ClientName>>": client_name,
        "<<CompanyName>>": company_name,
        "<<Date>>": date_input.strftime("%d-%m-%Y"),
        "<<StartDate>>": date_input.strftime("%d %B %Y"),
        "<<EndDate>>": end_date_input.strftime("%d %B %Y"),
        "<<Address>>": address,
    }

    template_name = "Contract Template.docx"
    
    if st.button("Generate Contract"):
        try:
            # Clear previous session state data
            for key in ["contract_docx", "contract_pdf", "contract_docx_name", "contract_pdf_name"]:
                if key in st.session_state:
                    st.session_state[key] = None if "name" not in key else ""

            # Define the hiring template file path
            template_path = os.path.join(os.getcwd(), template_name)
            
            # Verify the template exists
            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            
            safe_name = ''.join(c if c.isalnum() else '_' for c in client_name)
            
            # Save the contract to a temporary directory
            temp_dir = tempfile.gettempdir()
            docx_output_path = os.path.join(temp_dir, f"Contract_{safe_name}.docx")
            pdf_output_path = os.path.join(temp_dir, f"Contract_{safe_name}.pdf")

            # Edit the template and save the contract
            edit_hiring_template(template_path, docx_output_path, placeholders)
            # st.info("DOCX file created successfully. Converting to PDF...")

            # Load the generated DOCX file into session state for download
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.contract_docx = docx_file.read()
                st.session_state.contract_docx_name = f"Contract_{safe_name}.docx"

            # Convert DOCX to PDF with better error handling
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)
                # st.info(f"PDF conversion completed. Checking result...")
                
                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.contract_pdf = pdf_file.read()
                        st.session_state.contract_pdf_name = f"Contract_{safe_name}.pdf"
                    # st.success("PDF created successfully!")
                else:
                    st.warning("PDF file not found after conversion attempt.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                # Still allow DOCX download even if PDF fails
                st.warning("PDF conversion failed, but DOCX is available for download.")

            # Display download buttons based on what's available
            col1, col2 = st.columns(2)
            
            with col1:
                if st.session_state.contract_docx:
                    st.download_button(
                        label="📥 Download Contract (Word)",
                        data=st.session_state.contract_docx,
                        file_name=st.session_state.contract_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("DOCX file not available for download.")
                    
            with col2:
                if st.session_state.contract_pdf:
                    st.download_button(
                        label="📥 Download Contract (PDF)",
                        data=st.session_state.contract_pdf,
                        file_name=st.session_state.contract_pdf_name,
                        mime="application/pdf"
                    )
                else:
                    st.warning("PDF file not available for download.")
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())


def replace_text_in_paragraph(paragraph, placeholders):
    """Replace placeholders in a paragraph, preserving formatting and optionally bolding specific runs."""
    # Combine all run texts
    full_text = ''.join(run.text for run in paragraph.runs)

    for key, value in placeholders.items():
        if key in full_text:
            before, sep, after = full_text.partition(key)
            # Remove existing runs
            for run in list(paragraph.runs)[::-1]:
                paragraph._p.remove(run._r)
            # Add text before placeholder
            if before:
                paragraph.add_run(before)
            # Add replaced text
            new_run = paragraph.add_run(value)
            # Uncomment the next lines if you want to bold any specific placeholders:
            # if key in ('<<ClientName>>', '<<Date>>'):
            #     new_run.font.bold = True
            # Add text after placeholder
            if after:
                paragraph.add_run(after)
            # Stop after first replacement in this paragraph
            return


def replace_text_in_table(table, placeholders):
    """Iterate through table cells and replace placeholders."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, placeholders)


def edit_nda_template(template_path, output_path, placeholders):
    """Load template, replace placeholders in body and tables, then save."""
    doc = Document(template_path)

    # Replace in document body
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, placeholders)

    # Replace in all tables
    for table in doc.tables:
        replace_text_in_table(table, placeholders)

    # Ensure target directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path


def generate_nda():
    """Streamlit UI to collect inputs and generate the NDA document."""
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
    temp_dir = tempfile.gettempdir()


    if st.button("Generate NDA"):
        try:
            # Clear previous session state data
            for key in ["nda_docx", "nda_pdf", "nda_docx_name", "nda_pdf_name"]:
                if key in st.session_state:
                    st.session_state[key] = None if "name" not in key else ""

            # Define the hiring template file path
            template_path = os.path.join(os.getcwd(), template_name)
            
            # Verify the template exists
            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            
            safe_name = ''.join(c if c.isalnum() else '_' for c in client_name)
            
            # Save the hiring contract to a temporary directory
            temp_dir = tempfile.gettempdir()
            docx_output_path = os.path.join(temp_dir, f"NDA_{safe_name}.docx")
            pdf_output_path = os.path.join(temp_dir, f"NDA_{safe_name}.pdf")

            # Edit the hiring template and save the contract
            edit_nda_template(template_path, docx_output_path, placeholders)
            # st.info("DOCX file created successfully. Converting to PDF...")

            # Load the generated DOCX file into session state for download
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.nda_docx = docx_file.read()
                st.session_state.nda_docx_name = f"NDA_{safe_name}.docx"

            # Convert DOCX to PDF with better error handling
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)
                # st.info(f"PDF conversion completed. Checking result...")
                
                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.nda_pdf = pdf_file.read()
                        st.session_state.nda_pdf_name =f"NDA_{safe_name}.pdf"
                    # st.success("PDF created successfully!")
                else:
                    st.warning("PDF file not found after conversion attempt.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                # Still allow DOCX download even if PDF fails
                st.warning("PDF conversion failed, but DOCX is available for download.")

            # Display download buttons based on what's available
            col1, col2 = st.columns(2)
            
            with col1:
                if st.session_state.nda_docx:
                    st.download_button(
                        label="📥 Download NDA(Word)",
                        data=st.session_state.nda_docx,
                        file_name=st.session_state.nda_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("DOCX file not available for download.")
                    
            with col2:
                if st.session_state.nda_pdf:
                    st.download_button(
                        label="📥 Download NDA(PDF)",
                        data=st.session_state.nda_pdf,
                        file_name=st.session_state.nda_pdf_name,
                        mime="application/pdf"
                    )
                else:
                    st.warning("PDF file not available for download.")
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())

# def generate_document(option):
#     """Streamlit UI for generating NDA or Contract documents."""
#     if option== "NDA":
#         st.title("NDA Generator")
#     elif option == "Contract":
#         st.title("Contract Generator")
    

#     base_dir = os.path.abspath(os.path.dirname(__file__))
#     template_paths = {
#         "NDA": os.path.join(base_dir, "NDA Template.docx"),
#         "Contract": os.path.join(base_dir, "Contract Template.docx"),
#     }

#     # Check if template exists
#     if not os.path.exists(template_paths[option]):
#         st.error(f"Template file missing: {template_paths[option]}")
#         return

#     client_name = st.text_input("Enter Client Name:")
#     company_name = st.text_input("Enter Company Name:")
#     address = st.text_area("Enter Address:")
#     date_field = st.date_input("Enter Date:", datetime.today())

#     placeholders = {
#         "ClientName": client_name,
#         "CompanyName": company_name,
#         "Address": address,
#         "Date": date_field.strftime("%d-%m-%Y"),
#         "Date,": date_field.strftime("%d-%m-%Y"),
#         "ContractEndDate": date_field.replace(year=date_field.year + 1).strftime("%d-%m-%Y"),
#     }

#     if st.button(f"Generate Document"):
#         try:
#             # Clear previous session state data
#             for key in ['doc_data', 'pdf_data', 'filenames']:
#                 if key in st.session_state:
#                     del st.session_state[key]

#             formatted_date = date_field.strftime("%d %b %Y")
#             doc_filename = f"{option} - {client_name} {formatted_date}.docx"
#             pdf_filename = doc_filename.replace(".docx", ".pdf")

#             # Create temporary files
#             with tempfile.TemporaryDirectory() as temp_dir:
#                 doc_path = os.path.join(temp_dir, doc_filename)
#                 pdf_path = os.path.join(temp_dir, pdf_filename)

#                 # Generate DOCX
#                 font_size = 11 if option == "NDA" else 12
#                 doc = Document(template_paths[option])
#                 replace_and_format(doc, placeholders, "Times New Roman", font_size, option)

#                 doc.save(doc_path)
#                 logger.info(f"Document saved to {doc_path}")

#                 # Save DOCX to session state
#                 with open(doc_path, "rb") as doc_file:
#                     st.session_state.doc_data = doc_file.read()
                
#                 # Try to convert to PDF
#                 try:
#                     st.info("Converting document to PDF...")
#                     convert_to_pdf(doc_path, pdf_path)
                    
#                     if os.path.exists(pdf_path):
#                         with open(pdf_path, "rb") as pdf_file:
#                             st.session_state.pdf_data = pdf_file.read()
#                         logger.info(f"PDF conversion successful: {pdf_path}")
#                     else:
#                         st.warning("PDF conversion failed. Only DOCX will be available.")
#                         logger.warning(f"PDF file not found after conversion: {pdf_path}")
#                 except Exception as pdf_err:
#                     st.warning(f"Error during PDF conversion: {pdf_err}")
#                     logger.error(f"PDF conversion error: {pdf_err}", exc_info=True)
                
#                 st.session_state.filenames = {
#                     "doc": doc_filename,
#                     "pdf": pdf_filename
#                 }
            
#             st.success(f"{option} Document Generated Successfully!")

#         except Exception as e:
#             st.error(f"An error occurred: {e}")
#             logger.error(f"Error in generate_document: {e}", exc_info=True)
    
#     # Display download buttons
#     if 'doc_data' in st.session_state:
#         col1, col2 = st.columns(2)
#         with col1:
#             st.download_button(
#                 label="Download Document (Word)",
#                 data=st.session_state.doc_data,
#                 file_name=st.session_state.filenames["doc"],
#                 mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
#                 key="download_word"
#             )
        
#         if 'pdf_data' in st.session_state:
#             with col2:
#                 st.download_button(
#                     label="Download Document (PDF)",
#                     data=st.session_state.pdf_data,
#                     file_name=st.session_state.filenames["pdf"],
#                     mime="application/pdf",
#                     key="download_pdf"
#                 )
            

# Hiring Contract
def replace_text_in_paragraph(paragraph, placeholders):
    """Replace placeholders in paragraph text."""
    if not paragraph.runs:
        return
        
    full_text = "".join(run.text for run in paragraph.runs)
    
    for key, value in placeholders.items():
        if key in full_text:
            full_text = full_text.replace(key, value)
    
    # Clear all runs
    for i in range(len(paragraph.runs) - 1, 0, -1):
        p = paragraph._p
        p.remove(paragraph.runs[i]._r)
    
    # Assign replaced full text to the first run
    if paragraph.runs:
        paragraph.runs[0].text = full_text

# Function to edit the Hiring template and replace placeholders
def edit_hiring_template(template_path, output_path, placeholders):
    """Edit hiring contract template replacing placeholders."""
    doc = Document(template_path)

    def replace_text_in_table(table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, placeholders)

    for para in doc.paragraphs:
        replace_text_in_paragraph(para, placeholders)

    for table in doc.tables:
        replace_text_in_table(table)

    doc.save(output_path)
    return output_path

def generate_hiring_contract():
    """Streamlit UI for generating hiring contracts."""
    # Initialize session state for DOCX and PDF
    for key in ["hiring_docx", "hiring_pdf", "hiring_docx_name", "hiring_pdf_name"]:
        if key not in st.session_state:
            st.session_state[key] = None if "name" not in key else ""

    st.title("Hiring Contract Generator")

    # Collect inputs
    Employee_name = st.text_input("Enter Employee Name:")
    Role = st.text_input("Enter Role:")
    date = st.date_input("Enter Date:", datetime.today())
    Starting_Date = st.date_input("Enter the starting date: ")
    Stipend = st.number_input("Enter the Stipend:")
    Working_hours = st.number_input("Enter the total working hours:")
    Internship_duration = st.number_input("Enter the internship duration:")
    First_Pay_Cheque_Date = st.date_input("Enter the First Pay Cheque Date:")

    placeholders = {
        "<<Date>>": date.strftime("%d-%m-%Y"),
        "<<Name>>": Employee_name,
        "<<Role>>": Role,
        "<<Starting Date>>": Starting_Date.strftime("%d %B %Y"),
        "<<Internship Duration>>": str(int(Internship_duration)),
        "<<First Pay>>": First_Pay_Cheque_Date.strftime("%d %B %Y"),
        "<<Stipend>>": str(Stipend),
        "<<Working Hours>>": str(int(Working_hours))
    }

    template_name = "Hiring Contract.docx"
    
    if st.button("Generate Hiring Contract"):
        try:
            # Clear previous session state data
            for key in ["hiring_docx", "hiring_pdf", "hiring_docx_name", "hiring_pdf_name"]:
                if key in st.session_state:
                    st.session_state[key] = None if "name" not in key else ""

            # Define the hiring template file path
            template_path = os.path.join(os.getcwd(), template_name)
            
            # Verify the template exists
            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            
            safe_name = ''.join(c if c.isalnum() else '_' for c in Employee_name)
            
            # Save the hiring contract to a temporary directory
            temp_dir = tempfile.gettempdir()
            docx_output_path = os.path.join(temp_dir, f"Hiring_{safe_name}.docx")
            pdf_output_path = os.path.join(temp_dir, f"Hiring_{safe_name}.pdf")

            # Edit the hiring template and save the contract
            edit_hiring_template(template_path, docx_output_path, placeholders)
            # st.info("DOCX file created successfully. Converting to PDF...")

            # Load the generated DOCX file into session state for download
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.hiring_docx = docx_file.read()
                st.session_state.hiring_docx_name = f"Hiring_{safe_name}.docx"

            # Convert DOCX to PDF with better error handling
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)
                # st.info(f"PDF conversion completed. Checking result...")
                
                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.hiring_pdf = pdf_file.read()
                        st.session_state.hiring_pdf_name = f"Hiring_{safe_name}.pdf"
                    # st.success("PDF created successfully!")
                else:
                    st.warning("PDF file not found after conversion attempt.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                # Still allow DOCX download even if PDF fails
                st.warning("PDF conversion failed, but DOCX is available for download.")

            # Display download buttons based on what's available
            col1, col2 = st.columns(2)
            
            with col1:
                if st.session_state.hiring_docx:
                    st.download_button(
                        label="📥 Download Hiring Contract (Word)",
                        data=st.session_state.hiring_docx,
                        file_name=st.session_state.hiring_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("DOCX file not available for download.")
                    
            with col2:
                if st.session_state.hiring_pdf:
                    st.download_button(
                        label="📥 Download Hiring Contract (PDF)",
                        data=st.session_state.hiring_pdf,
                        file_name=st.session_state.hiring_pdf_name,
                        mime="application/pdf"
                    )
                else:
                    st.warning("PDF file not available for download.")
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())


def format_price(amount, currency):
    """Format price based on currency."""
    formatted_price = f"{amount:,.2f}"
    return f"{currency} {formatted_price}" if currency == "USD" else f"Rs. {formatted_price}"

# Function to format percentages
def format_percentage(value):
    """Format percentage without decimals."""
    return f"{int(value)}%"

# Function to get the next invoice number
def get_next_invoice_number():
    """Fetch and increment the invoice number."""
    invoice_file = "invoice_counter.txt"
    if not os.path.exists(invoice_file):
        with open(invoice_file, "w") as file:
            file.write("1000")  # Starting invoice number
    try:
        with open(invoice_file, "r") as file:
            current_invoice = int(file.read().strip())
    except ValueError:
        current_invoice = 1000
    next_invoice = current_invoice + 1
    with open(invoice_file, "w") as file:
        file.write(str(next_invoice))
    return current_invoice

# Function to convert amount to words
def amount_to_words(amount):
    """Convert amount to words without currency formatting."""
    try:
        words = num2words(amount, lang='en').replace(',', '').title()
        return words
    except Exception as e:
        logger.error(f"Error converting amount to words: {e}", exc_info=True)
        return f"[Error: Unable to convert {amount} to words]"

# Function to replace placeholders in DOCX
def replace_placeholders(doc, placeholders):
    """Replace placeholders in paragraphs and tables."""
    for para in doc.paragraphs:
        for key, value in placeholders.items():
            if key in para.text:
                for run in para.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)
                        # Apply bold formatting for specific placeholders
                        if key.startswith("<<Price") or key.startswith("<<Total") or key == "<<Amt to Word>>":
                            run.bold = True  # Apply bold formatting

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        for para in cell.paragraphs:
                            for run in para.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, value)
                                    # Apply bold formatting for specific placeholders
                                    if key.startswith("<<Price") or key.startswith("<<Total") or key == "<<Amt to Word>>":
                                        run.bold = True  # Apply bold formatting
    return doc

# Function to edit invoice template and save the result
def edit_invoice_template(template_name, output_path, placeholders):
    """Edit an invoice template and save the result."""
    try:
        doc = Document(template_name)
        replace_placeholders(doc, placeholders)
        doc.save(output_path)
        return output_path
    except Exception as e:
        logger.error(f"Error editing invoice template: {e}", exc_info=True)
        raise Exception(f"Error editing invoice template: {e}")


def generate_invoice():
    """Streamlit app for generating invoices."""
    st.title("Invoice Generator")
    # Select region
    region = st.selectbox("Select Region", ["INR", "USD"])

    # Set payment options dynamically
    payment_options = ["1 Payment", "3 EMI", "5 EMI"] if region == "INR" else ["3 EMI", "5 EMI"]

    # Input Fields
    client_name = st.text_input("Client Name")
    client_address = st.text_input("Client Address")
    project_name = st.text_input("Project Name")
    phone_number = st.text_input("Phone Number")
    gst_number = st.text_input("GST Number")
    base_amount = st.number_input("Base Amount (excluding GST)", min_value=0.0, format="%.2f")
    payment_option = st.selectbox("Payment Option", payment_options)
    invoice_date = st.date_input("Invoice Date", value=datetime.today())
    formatted_date = invoice_date.strftime("%d-%m-%Y")

    # Calculate GST and total amount
    gst_amount = round(base_amount * 0.18)
    total_amount = base_amount + gst_amount

    # Prepare placeholders for template
    placeholders = {
        "<<Client Name>>": client_name,
        "<<Client Address>>": client_address,
        "<<GST Number>>": gst_number,
        "<<Client Email>>": client_address,
        "<<Project Name>>": project_name,
        "<<Mobile Number>>": phone_number,
        "<<Date>>": formatted_date,
        "<<Amt to word>>": amount_to_words(int(total_amount)),
    }
# Select the correct template based on payment option
    if payment_option == "1 Payment":
        template_name = f"Invoice Template - {region} - 1 Payment 1.docx"
        placeholders.update({
            "<<Price 1>>": format_price(base_amount, region),
            "<<Price 2>>": format_price(gst_amount, region),
            "<<Price 3>>": format_price(total_amount, region),
            "<<Total 1>>": format_price(total_amount, region),
        })

    elif payment_option == "3 EMI":
        template_name = f"Invoice Template - {region} - 3 EMI Payment Schedule 1.docx"
        p1 = round(total_amount * 0.30)
        p2 = round(total_amount * 0.40)
        p3 = total_amount - (p1 + p2)
        placeholders.update({
            "<<Price 1>>": format_price(p1, region),
            "<<Price 2>>": format_price(p2, region),
            "<<Price 3>>": format_price(p3, region),
            "<<Price 4>>": format_price(gst_amount, region),
            "<<Price 5>>": format_price(total_amount, region),
            "<<Price 6>>": format_price(p1, region),
            "<<Price 7>>": format_price(p2, region),
            "<<Price 8>>": format_price(p3, region),
        })

    elif payment_option == "5 EMI":
        template_name = f"Invoice Template - {region} - 5 EMI Payment Schedule 1.docx"
        p1 = round(total_amount * 0.20)
        p2 = round(total_amount * 0.20)
        p3 = round(total_amount * 0.20)
        p4 = round(total_amount * 0.20)
        p5 = total_amount - (p1 + p2 + p3 + p4)
        placeholders.update({
            "<<Price 1>>": format_price(p1, region),
            "<<Price 2>>": format_price(p2, region),
            "<<Price 3>>": format_price(p3, region),
            "<<Price 4>>": format_price(p4, region),
            "<<Price 5>>": format_price(p5, region),
            "<<Price 6>>": format_price(p1, region),
            "<<Price 7>>": format_price(p2, region),
            "<<Price 8>>": format_price(p3, region),
            "<<Price 9>>": format_price(p4, region),
            "<<Price 10>>": format_price(p5, region),
            "<<Total 1>>": format_price(total_amount, region),
        })

    # Generate Invoice Button
    if st.button("Generate Invoice"):
        try:
            for key in ["invoice_docx","invoice_pdf","invoice_docx_name","invoice_pdf_name"]:
                if key in st.session_state:
                    st.session_state[key] =None if "name" not in key else ""

            # Generate the next invoice number
            invoice_number = get_next_invoice_number()
            placeholders["<<Invoice>>"] = str(invoice_number)
            placeholders["<<Invoice No>>"] = str(invoice_number)

            # Define the invoice template file path
            template_path = os.path.join(os.getcwd(), template_name)

            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return

            # Save the invoice to a temporary directory
            temp_dir = tempfile.gettempdir()
            sanitized_client_name = "".join([c if c.isalnum() or c.isspace() else "_" for c in client_name])
            docx_output_path = os.path.join(temp_dir, f"Invoice_{sanitized_client_name}_{invoice_number}.docx")
            pdf_output_path = os.path.join(temp_dir, f"Invoice_{sanitized_client_name}_{invoice_number}.pdf")

            # Edit the template and save the invoice
            edit_invoice_template(template_path, docx_output_path, placeholders)
            
           #Save the file to session
            with open(docx_output_path, "rb") as docx_file:
                st.session_state.invoice_docx = docx_file.read()
                st.session_state.invoice_docx_name = f"Invoice_{sanitized_client_name}_{invoice_number}.docx"
            # Convert DOCX to PDF with better error handling
            try:
                convert_to_pdf(docx_output_path, pdf_output_path)
                st.info(f"PDF conversion completed. Checking result...")
                
                if os.path.exists(pdf_output_path):
                    with open(pdf_output_path, "rb") as pdf_file:
                        st.session_state.invoice_pdf = pdf_file.read()
                        st.session_state.invoice_pdf_name = f"Invoice_{sanitized_client_name}_{invoice_number}.pdf"
                    st.success("PDF created successfully!")
                else:
                    st.warning("PDF file not found after conversion attempt.")
            except Exception as pdf_err:
                st.error(f"PDF Conversion Error: {pdf_err}")
                # Still allow DOCX download even if PDF fails
                st.warning("PDF conversion failed, but DOCX is available for download.")

            # Display download buttons based on what's available
            col1, col2 = st.columns(2)
            
            with col1:
                if st.session_state.invoice_docx:
                    st.download_button(
                        label="📥 Download Invoice (Word)",
                        data=st.session_state.invoice_docx,
                        file_name=st.session_state.invoice_docx_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("DOCX file not available for download.")
                    
            with col2:
                if st.session_state.invoice_pdf:
                    st.download_button(
                        label="📥 Download Invoice (PDF)",
                        data=st.session_state.invoice_pdf,
                        file_name=st.session_state.invoice_pdf_name,
                        mime="application/pdf"
                    )
                else:
                    st.warning("PDF file not available for download.")
                    
        except Exception as e:
            st.error(f"An error occurred: {e}")
            import traceback
            st.code(traceback.format_exc())


def main():
    st.sidebar.title("Select Application")
    app_choice = st.sidebar.radio("Choose an application:", ["NDA", "Contract", "Invoice", "Pricing List", "Proposal","Hiring Contract"])
    if app_choice == "Invoice":
        generate_invoice()
    
    elif app_choice=="Contract":
        generate_contract()

    elif app_choice == "NDA":
        generate_nda()

    elif app_choice == "Hiring Contract":
        generate_hiring_contract()
    


if __name__ == "__main__":
    main()