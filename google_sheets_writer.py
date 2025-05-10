from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os

# Konfiguracja
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
FOLDER_ID = '17gKONL0gLBx7Wvd4Cx3cBVEhHVdLQP2_'  # <- Twój folder w Google Drive

def get_google_creds():
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            print("⚠️ Błędny token. Usuwam...")
            os.remove(TOKEN_FILE)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # ✅ Zapis jako poprawny JSON
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def create_spreadsheet(creds, title, folder_id=None):
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet_body = {
            'properties': {'title': title}
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet_body, fields='spreadsheetId').execute()
        sheet_id = spreadsheet.get('spreadsheetId')

        # 🟡 Umieść w folderze jeśli podano
        if folder_id:
            drive_service = build('drive', 'v3', credentials=creds)
            drive_service.files().update(fileId=sheet_id, addParents=folder_id, removeParents='root').execute()

        return sheet_id
    except HttpError as err:
        print(f"Błąd HTTP: {err}")
        return None

def write_to_spreadsheet(creds, spreadsheet_id, values):
    try:
        service = build('sheets', 'v4', credentials=creds)
        body = {
            'values': values
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range='A1',
            valueInputOption='RAW',
            body=body
        ).execute()
        print(f"Zapisano {result.get('updatedCells')} komórek.")
    except HttpError as err:
        print(f"Błąd zapisu: {err}")

if __name__ == '__main__':
    creds = get_google_creds()

    # 🔹 Przykład: utwórz arkusz na podstawie tekstu z programu
    title = "Testowy_plik_z_Python"
    spreadsheet_id = create_spreadsheet(creds, title, folder_id=FOLDER_ID)

    if spreadsheet_id:
        print(f"✅ Utworzono arkusz: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")

        # Przykładowe dane
        sample_data = [
            ["X", "Detal", "PR", "Ilość", "Uwagi"],
            ["[X]", "ABC123", "PR-01", "2", ""],
            ["[]", "XYZ999", "PR-02", "5", "Brak"]
        ]
        write_to_spreadsheet(creds, spreadsheet_id, sample_data)
