from flask import Flask, request, jsonify
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
TOKEN_FILE = 'token.json'
FOLDER_ID = '17gKONL0gLBx7Wvd4Cx3cBVEhHVdLQP2_'  # Folder zamówień

def get_credentials():
    return Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

def find_sheet_by_name(drive_service, name):
    results = drive_service.files().list(
        q=f"name contains '{name}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    return files[0] if files else None

@app.route('/save', methods=['POST'])
def save_data():
    try:
        data = request.json
        filename = data.get('filename')
        rows = data.get('rows')

        if not filename or not rows:
            return jsonify({'error': 'Brak danych'}), 400

        creds = get_credentials()
        sheets_service = build('sheets', 'v4', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        existing = find_sheet_by_name(drive_service, filename)
        if existing:
            file_id = existing['id']
            sheets_service.spreadsheets().values().update(
                spreadsheetId=file_id,
                range='A1',
                valueInputOption='RAW',
                body={'values': rows}
            ).execute()
            return jsonify({'message': 'Zaktualizowano', 'id': file_id})
        else:
            sheet = sheets_service.spreadsheets().create(
                body={'properties': {'title': filename}},
                fields='spreadsheetId'
            ).execute()
            file_id = sheet['spreadsheetId']

            # Przenieś do folderu
            drive_service.files().update(
                fileId=file_id,
                addParents=FOLDER_ID,
                removeParents='root',
                fields='id, parents'
            ).execute()

            sheets_service.spreadsheets().values().update(
                spreadsheetId=file_id,
                range='A1',
                valueInputOption='RAW',
                body={'values': rows}
            ).execute()

            return jsonify({'message': 'Utworzono', 'id': file_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/load', methods=['GET'])
def load_data():
    try:
        filename = request.args.get('name')
        if not filename:
            return jsonify({'error': 'Brak nazwy pliku'}), 400

        creds = get_credentials()
        sheets_service = build('sheets', 'v4', credentials=creds)
        drive_service = build('drive', 'v3', credentials=creds)

        file = find_sheet_by_name(drive_service, filename)
        if not file:
            return jsonify({'error': 'Nie znaleziono pliku'}), 404

        file_id = file['id']
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=file_id,
            range='A1:E'
        ).execute()
        values = result.get('values', [])

        return jsonify({'data': values, 'filename': file['name']})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/list', methods=['GET'])
def list_sheets():
    try:
        creds = get_credentials()
        drive_service = build('drive', 'v3', credentials=creds)

        results = drive_service.files().list(
            q=f"'{FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'",
            fields="files(name)"
        ).execute()

        files = results.get('files', [])
        names = [f['name'] for f in files]

        return jsonify(names)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/')
def index():
    return 'LIFTRA API działa!'
