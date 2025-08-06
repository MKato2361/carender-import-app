import streamlit as st
import firebase_admin
from firebase_admin import auth, credentials, firestore
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
import os

# --- Firebase Initialization ---
def init_firebase():
    if not firebase_admin._apps:
        cred = credentials.Certificate("firebase_credentials.json")
        firebase_admin.initialize_app(cred)

# --- Google OAuth 2.0 Flow ---
def get_google_auth_url():
    """Generates the Google OAuth 2.0 authorization URL."""
    init_firebase()
    
    # We need to get the redirect URI, which is the current app URL.
    if st.secrets.get("google_oauth", {}).get("redirect_uri"):
        redirect_uri = st.secrets.google_oauth.redirect_uri
    else:
        st.error("Please configure `redirect_uri` in your Streamlit secrets.")
        return "#"

    flow = Flow.from_client_config(
        client_config={
            "web": {
                "client_id": st.secrets.google_oauth.client_id,
                "project_id": st.secrets.firebase.project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets.google_oauth.client_secret,
                "redirect_uris": [redirect_uri],
            }
        },
        scopes=st.secrets.google_oauth.scopes,
        redirect_uri=redirect_uri,
    )

    authorization_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
    )
    return authorization_url

def handle_google_auth_callback(auth_code):
    """Exchanges the authorization code for tokens and handles user creation/login."""
    init_firebase()
    
    # We need to get the redirect URI, which is the current app URL.
    redirect_uri = st.secrets.google_oauth.redirect_uri

    flow = Flow.from_client_config(
        client_config={
            "web": {
                "client_id": st.secrets.google_oauth.client_id,
                "project_id": st.secrets.firebase.project_id,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_secret": st.secrets.google_oauth.client_secret,
                "redirect_uris": [redirect_uri],
            }
        },
        scopes=st.secrets.google_oauth.scopes,
        redirect_uri=redirect_uri,
    )
    
    flow.fetch_token(code=auth_code)
    
    creds = flow.credentials
    user_info = auth.get_user(creds.id_token) # Get user info from id_token

    db = firestore.client()
    doc_ref = db.collection("google_tokens").document(user_info.uid)
    doc_ref.set({
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
        "expires_at": creds.expiry.isoformat(),
    })
    
    google_auth_info = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
        "expiry": creds.expiry,
    }
    
    return user_info, google_auth_info
