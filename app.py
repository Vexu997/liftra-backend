from flask import Flask, request, jsonify
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import os

app = Flask(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
TOKEN_FILE = 'token.json'

def get_google_creds():
    return Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

@app.route('/list-files', methods=['GET'])
def list_files():
    try:
        creds = get_google_creds()
        service = build('drive', 'v3', credentials=creds)

        response = service.files().list(
            q="'17gKONL0gLBx7Wvd4Cx3cBVEhHVdLQP2_' in parents and mimeType='application/vnd.google-apps.spreadsheet'",
            fields="files(id, name)"
        ).execute()

        files = response.get('files', [])
        return jsonify(files)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/get-sheet', methods=['GET'])
def get_sheet():
    sheet_id = request.args.get('id')
    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)

        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id, range="A1:E"
        ).execute()

        values = result.get('values', [])
        return jsonify(values)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/save-sheet', methods=['POST'])
def save_sheet():
    data = request.get_json()
    sheet_id = data.get('id')
    values = data.get('values')

    try:
        creds = get_google_creds()
        service = build('sheets', 'v4', credentials=creds)

        body = {'values': values}
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range='A1', valueInputOption='RAW', body=body
        ).execute()

        return jsonify({'status': 'success'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)
