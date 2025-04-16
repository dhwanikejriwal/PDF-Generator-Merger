import streamlit as st
import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import io
import base64

# Set folders
First_Page_Folder = "pdfs/First_page"
Last_Page_Folder = "pdfs/Last_page"

# Create folders if running locally (optional)
os.makedirs(First_Page_Folder, exist_ok=True)
os.makedirs(Last_Page_Folder, exist_ok=True)

st.set_page_config(page_title="PDF Merger", layout="centered")
st.title("📄 Merge PDFs with Preview")

# Get all PDF files from folders
first_page_files = [f for f in os.listdir(First_Page_Folder) if f.endswith(".pdf")]
last_page_files = [f for f in os.listdir(Last_Page_Folder) if f.endswith(".pdf")]

selected_first_page = st.selectbox("📄 Select First Page", first_page_files)
selected_last_page = st.selectbox("📄 Select Last Page", last_page_files)

# Regenerate a PDF in memory (to fix minor issues)
def regenerate_pdf(file_path):
    reader = PdfReader(file_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()

# Preview function with scrollable iframe and fallback link
def preview_pdf(file_path, label):
    try:
        with open(file_path, "rb") as f:
            pdf_bytes = f.read()
            if not pdf_bytes or len(pdf_bytes) < 100:
                st.warning(f"{label} PDF might be empty or corrupted.")
                return

            # Regenerate the PDF (optional fix for corruption)
            pdf_bytes = regenerate_pdf(file_path)

            base64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

            st.markdown(f"#### Preview: {label}")
            iframe_html = f'''
                <iframe 
                    src="data:application/pdf;base64,{base64_pdf}#page=1" 
                    width="100%" height="600px" 
                    style="border: none;"
                ></iframe>
            '''
            st.markdown(iframe_html, unsafe_allow_html=True)

            # Fallback view link
            view_link = f'<a href="data:application/pdf;base64,{base64_pdf}" target="_blank">🔍 Open full {label} in new tab</a>'
            st.markdown(view_link, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error loading {label} PDF: {str(e)}")

# Preview the selected PDFs
if selected_first_page:
    preview_pdf(os.path.join(First_Page_Folder, selected_first_page), "First Page")

if selected_last_page:
    preview_pdf(os.path.join(Last_Page_Folder, selected_last_page), "Last Page")

# Merge and download
if st.button("🔗 Merge Selected PDFs"):
    path_a = os.path.join(First_Page_Folder, selected_first_page)
    path_b = os.path.join(Last_Page_Folder, selected_last_page)

    merger = PdfMerger()
    merger.append(path_a)
    merger.append(path_b)

    output = io.BytesIO()
    merger.write(output)
    merger.close()
    output.seek(0)

    st.success("✅ PDFs merged successfully!")
    st.download_button(
        label="📥 Download Merged PDF",
        data=output,
        file_name=f"{selected_first_page.split('.')[0]}_{selected_last_page.split('.')[0]}_merged.pdf",
        mime="application/pdf"
    )
