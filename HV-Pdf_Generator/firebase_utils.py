import streamlit as st
from google.cloud import firestore
from firebase_config import initialize_firebase

bucket, db = initialize_firebase()

def upload_to_firebase(uploaded_file, name):
    """Upload file to Firebase Storage and store details in Firestore."""
    blob = bucket.blob(f"uploaded_docs/{uploaded_file.name}")
    blob.upload_from_file(uploaded_file, content_type=uploaded_file.type)
    blob.make_public()

    link = blob.public_url
    # Store the document details (name and link) in Firestore
    db.collection("ProposalPDFPage2").document().set({
        "name": name,
        "link": link
    })
    st.success("File uploaded successfully!")
    st.markdown(f"[Click to View]({link})")


def show_documents():
    """Fetch and display uploaded documents from Firestore."""
    st.markdown("### Uploaded Documents")
    # Fetch documents from Firestore
    docs = list(db.collection("ProposalPDFPage2").stream())

    # Loop through the documents and display their details
    for idx, doc in enumerate(docs, 1):
        doc_data = doc.to_dict()
        name = doc_data.get("name", "No Name")
        link = doc_data.get("link", "").strip()

        st.markdown(f"**{idx}. {name}**")
        
        # Display the link as a clickable link for all document types
        if link:
            st.markdown(f"🔗 [Click to View Document]({link})")
        else:
            st.warning("_No link available_")


def update_document(doc_id, new_name, new_link):
    db.collection("ProposalPDFPage2").document(doc_id).update({
        "name": new_name,
        "link": new_link.strip()
    })
    st.success("Document updated successfully!")


def delete_document(doc_id):
    db.collection("ProposalPDFPage2").document(doc_id).delete()
    st.success("Document deleted successfully!")


def manage_documents():
    """Allow users to manage (update or delete) uploaded documents."""
    docs = list(db.collection("ProposalPDFPage2").stream())

    if not docs:
        st.info("No documents found.")
        return

    # Create a dropdown to select which document to update or delete
    doc_options = {f"{doc.to_dict().get('name', 'Unnamed')} ({doc.id})": doc.id for doc in docs}
    selected_label = st.selectbox("Select a document to update or delete", list(doc_options.keys()))
    selected_id = doc_options[selected_label]
    selected_data = db.collection("ProposalPDFPage2").document(selected_id).get().to_dict()

    st.markdown("### Update Document")
    with st.form("update_form"):
        updated_name = st.text_input("Updated Name", selected_data.get("name", ""))
        updated_link = st.text_input("Updated Link", selected_data.get("link", ""))
        update_btn = st.form_submit_button("Update")
        if update_btn:
            update_document(selected_id, updated_name, updated_link)

    st.markdown("### Delete Document")
    if st.button("Delete This Document"):
        delete_document(selected_id)
