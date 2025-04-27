import streamlit as st

def initialize_session_state():
    """Initialize all session state variables used in the app."""
    keys = [
        "nda_docx", "nda_pdf", "nda_docx_name", "nda_pdf_name",
        "contract_docx", "contract_pdf", "contract_docx_name", "contract_pdf_name",
        "hiring_docx", "hiring_pdf", "hiring_docx_name", "hiring_pdf_name",
        "invoice_docx", "invoice_pdf", "invoice_docx_name", "invoice_pdf_name"
    ]
    for key in keys:
        if key not in st.session_state:
            st.session_state[key] = None if "name" not in key else ""

def clear_session_keys(keys):
    """Clear specific keys from Streamlit session_state."""
    for key in keys:
        if key in st.session_state:
            st.session_state[key] = None if "name" not in key else ""
