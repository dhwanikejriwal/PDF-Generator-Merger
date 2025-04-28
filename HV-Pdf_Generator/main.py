import streamlit as st
from session_manager import initialize_session_state
from firebase_utils import upload_to_firebase , show_documents , manage_documents
from firebase_config import initialize_firebase

from generators.nda import generate_nda
from generators.hiring import generate_hiring
from generators.invoice import generate_invoice
from generators.contract import generate_contract


initialize_firebase()
initialize_session_state()

def main():

    document_type = {
        "NDA",
        "Contract",
        "Hiring Contract",
        "Invoice"
    }

    operations = {
        "Upload Documents",
        "View Documents",
        "Update/Delete Documents"
        }

    st.set_page_config(page_title="Documnet Generator and firebase Manager" , layout="wide")
    st.sidebar.title("Application menu")

    section = st.sidebar.radio("Choose Section",["Document Generator" , "Firebase Crud Operations"])

    if section == "Document Generator":
        doc_choice = st.sidebar.radio("Select Document type" , document_type)

        if doc_choice == "NDA":
            generate_nda()
        
        elif doc_choice == "Invoice":
            generate_invoice()

        elif doc_choice == "Hiring Contract":
            generate_hiring()

        elif doc_choice == "Contract":
            generate_contract()

    elif section == "Firebase Crud Operations":
        crud_choice = st.sidebar.radio("Choose Operation" , operations)

        if crud_choice == "Upload Documents":
            st.subheader("Upload Document to Firebase")

            with st.form("upload_form"):
                name = st.text_input("Enter the doucment name")

                uploaded_file = st.file_uploader("Choose a file" , type = ["pdf"])

                submit_btm = st.form_submit_button("Upload")

                if submit_btm and uploaded_file and name:
                    upload_to_firebase(uploaded_file , name)

        
        elif crud_choice == "View Documents":
            st.subheader("View Uploaded Documents")
            show_documents()

        elif crud_choice == "Update/Delete Documents":
            st.subheader("Update/Delete Documents")
            manage_documents()


if __name__ == "__main__":
    main()
