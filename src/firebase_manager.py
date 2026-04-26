import os
import json
import time
import requests
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

class FirebaseAuthManager:
    """Handles Google OAuth and exchanges it for Firebase Auth token."""
    def __init__(self, client_secrets_path: str):
        self.client_secrets_path = client_secrets_path
        
        # Read API key and Project ID from client_secrets.json or env vars
        try:
            with open(self.client_secrets_path, "r") as f:
                secrets = json.load(f)
                web_or_installed = secrets.get("installed", secrets.get("web", {}))
                self.api_key = web_or_installed.get("firebase_api_key") or os.environ.get("FIREBASE_API_KEY")
                self.project_id = web_or_installed.get("project_id") or os.environ.get("FIREBASE_PROJECT_ID")
        except FileNotFoundError:
            self.api_key = os.environ.get("FIREBASE_API_KEY")
            self.project_id = os.environ.get("FIREBASE_PROJECT_ID")
            
        if not self.api_key or not self.project_id:
            raise ValueError("Firebase API Key or Project ID missing from client_secrets.json and environment variables.")
        
        self.id_token = None
        self.refresh_token_val = None
        self._token_expiry = 0
        self.uid = None
        self.email = None
        
        self.db_url = f"https://firestore.googleapis.com/v1/projects/{self.project_id}/databases/(default)/documents"
        
        self.token_cache_file = os.path.join(os.path.dirname(client_secrets_path), "token.json")

    def login(self) -> None:
        """Runs the Google OAuth local server and signs into Firebase."""
        creds = None
        
        if os.path.exists(self.token_cache_file):
            try:
                creds = Credentials.from_authorized_user_file(self.token_cache_file, ["openid", "https://www.googleapis.com/auth/userinfo.email"])
            except Exception:
                pass
            
        if not creds or not creds.valid or not creds.id_token:
            if creds and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.client_secrets_path,
                        scopes=["openid", "https://www.googleapis.com/auth/userinfo.email"]
                    )
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.client_secrets_path,
                    scopes=["openid", "https://www.googleapis.com/auth/userinfo.email"]
                )
                creds = flow.run_local_server(port=0)
            
            with open(self.token_cache_file, "w") as token:
                token.write(creds.to_json())
        
        # Now exchange Google ID token for Firebase token
        url = f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithIdp?key={self.api_key}"
        payload = {
            "postBody": f"id_token={creds.id_token}&providerId=google.com",
            "requestUri": "http://localhost",
            "returnIdpCredential": True,
            "returnSecureToken": True
        }
        res = requests.post(url, json=payload)
        
        if res.status_code != 200:
            print("Failed to authenticate with Firebase:")
            print(res.text)
            if "INVALID_IDP_RESPONSE" in res.text:
                print("Make sure Google Sign-In is enabled in your Firebase Console > Authentication > Sign-in method!")
            res.raise_for_status()
        
        data = res.json()
        self.id_token = data.get("idToken")
        self.uid = data.get("localId")
        self.email = data.get("email")
        self.refresh_token_val = data.get("refreshToken")
        self._token_expiry = time.time() + int(data.get("expiresIn", 3600)) - 60

    def _refresh_id_token(self):
        url = f"https://securetoken.googleapis.com/v1/token?key={self.api_key}"
        payload = {"grant_type": "refresh_token", "refresh_token": self.refresh_token_val}
        res = requests.post(url, json=payload)
        res.raise_for_status()
        data = res.json()
        self.id_token = data.get("id_token")
        self.refresh_token_val = data.get("refresh_token")
        self._token_expiry = time.time() + int(data.get("expires_in", 3600)) - 60

    @property
    def auth_headers(self) -> dict:
        if not self.id_token:
            raise ValueError("User not logged in. Call login() first.")
        if time.time() >= self._token_expiry:
            self._refresh_id_token()
        return {"Authorization": f"Bearer {self.id_token}"}

    def get_bets_url(self, doc_id: str = "") -> str:
        """Returns the Firestore URI for the user's bet collection or a specific document."""
        base = f"{self.db_url}/users/{self.uid}/matchbets"
        return f"{base}/{doc_id}" if doc_id else base
