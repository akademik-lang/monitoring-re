"""
auth_setup.py
=============
Jalankan SEKALI untuk mendapatkan/refresh token Google Drive.
Akan membuka browser untuk otorisasi.

Usage:
  python auth_setup.py
"""

import os
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
CRED_FILE  = os.path.join(BASE_DIR, "auth", "credentials.json")
TOKEN_FILE = os.path.join(BASE_DIR, "auth", "token.json")

# Scope: baca metadata file Drive (list files di folder)
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

def main():
    print("=" * 50)
    print("  Google Drive Auth Setup")
    print("=" * 50)

    creds = None

    # Cek token lama
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        if creds and creds.valid:
            print(f"\n✅ Token sudah valid!")
            print(f"   Scope: {creds.scopes}")
            print(f"   Tidak perlu re-auth.")
            return
        elif creds and creds.expired and creds.refresh_token:
            print("\nToken expired, mencoba refresh...")
            try:
                creds.refresh(Request())
                print("✅ Token berhasil di-refresh!")
                with open(TOKEN_FILE, 'w') as f:
                    f.write(creds.to_json())
                print(f"   Token disimpan ke: {TOKEN_FILE}")
                return
            except Exception as e:
                print(f"  Refresh gagal: {e}. Akan re-auth ulang.")
                creds = None

    # Token tidak ada atau tidak valid → buka browser
    print("\nMembuka browser untuk otorisasi Google Drive...")
    print("Silakan login dan klik 'Allow' di browser.\n")

    flow = InstalledAppFlow.from_client_secrets_file(CRED_FILE, SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)

    # Simpan token baru
    with open(TOKEN_FILE, 'w') as f:
        f.write(creds.to_json())

    print(f"\n✅ Otorisasi berhasil!")
    print(f"   Token disimpan ke: {TOKEN_FILE}")
    print(f"   Scope: {creds.scopes}")
    print(f"\nSekarang jalankan: python generate_excel.py")

if __name__ == "__main__":
    main()
