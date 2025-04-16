import streamlit as st
import os
from PyPDF2 import PdfMerger
import io
import base64

import uuid


# Set up folder for different types
First_Page_Folder = "pdfs\First_page"
Last_Page_Folder = "pdfs\Last_page"

os.makedirs(First_Page_Folder, exist_ok=True)
os.makedirs(Last_Page_Folder, exist_ok=True)

st.set_page_config(page_title = "PDF Merger", layout = "centered")

st.title("Merge PDFs")

first_page_files = [f for f in os.listdir(First_Page_Folder) if f.endswith(".pdf")]
last_page_files = [f for f in os.listdir(Last_Page_Folder) if f.endswith(".pdf")]

selected_first_page = st.selectbox("Select First Page",first_page_files)
selected_last_page = st.selectbox("Select Last Page",last_page_files)

if st.button("Merge Selected PDFs"):
    path_a = os.path.join(First_Page_Folder , selected_first_page)
    path_b = os.path.join(Last_Page_Folder,selected_last_page)

    merger = PdfMerger()
    merger.append(path_a)
    merger.append(path_b)

    output = io.BytesIO()
    merger.write(output)
    merger.close()
    output.seek(0)

    st.success("PDFs merged successfully!")
    st.download_button(
        label="📥 Download Merged PDF",
        data=output,
        file_name=f"{selected_first_page.split('.')[0]}_{selected_last_page.split('.')[0]}_merged.pdf",
        mime="application/pdf"
    )