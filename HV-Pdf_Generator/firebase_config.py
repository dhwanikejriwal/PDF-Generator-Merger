import streamlit as st
import firebase_admin
from firebase_admin import credentials, storage, firestore

bucket = None
db = None

def initialize_firebase():
    global bucket, db

    firebase_info = dict(st.secrets["FIREBASE"])
    bucket_name = "hv-technologies.appspot.com"  # Replace with your real bucket

    if not firebase_admin._apps:
        cred = credentials.Certificate(firebase_info)
        firebase_admin.initialize_app(cred, {
            'storageBucket': bucket_name
        })

    bucket = storage.bucket()
    db = firestore.client()   # ✅ use firebase_admin's firestore, not google's firestore.Client()

    return bucket, db
