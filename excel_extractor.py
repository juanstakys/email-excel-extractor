import os
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Alcance que tiene la app sobre la cuenta de Gmail. Si se modifica, eliminar token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


# Cargar credenciales
def loadCredentials():

    creds = None

    # El archivo token.json guarda el acceso del usuario y se refresca para no tener que pasar
    # por el proceso de autorización al ejecutar nuevamente el programa. El archivo se crea
    # automáticamente cuando el flujo de autorización completa por primera vez.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # Si no hay credenciales o no hay credenciales válidas, pide al usuario que inicie sesión.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Guarda las credenciales en el archivo token.json para la siguiente ejecución.
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def main():

    store_dir = 'downloaded_attachments/'
    subject_to_detect = 'email reto'

    # Cargar credenciales
    creds = loadCredentials()

    
    try:
        # Llama a la API de Gmail.
        service = build('gmail', 'v1', credentials=creds)
        results = service.users().messages().list(userId='me').execute()
        messages = results.get('messages', [])

        def getSubject(message):
            # Devuelve el asunto del mensaje.
            headers = service.users().messages().get(
                userId='me', id=message['id']).execute()['payload']['headers']
            for header in headers:
                if header['name'] == 'Subject':
                    return header['value']
            return None

        def getAttachments(message):
            # Devuelve los adjuntos del mensaje como generator iterator.
            messageData = service.users().messages().get(
                userId='me', id=message['id']).execute()
            if('parts' in messageData['payload'].keys()):
                for part in messageData['payload']['parts']:
                    if part['filename']:
                        attachment = service.users().messages().attachments().get(
                            userId='me', messageId=message['id'], id=part['body']['attachmentId']).execute()
                        yield {"filename": part["filename"], "data": attachment['data']}

        print(
            f"Searching for messages with subject: '{subject_to_detect}' and extracting excel attachments")

        if not messages:
            print('No messages found.')
            return

        for message in messages:
            subject = getSubject(message)
            if subject.lower() == subject_to_detect.lower():
                print(f"Subject: {subject} found!")
                attachments = list(getAttachments(message))
                if not attachments:
                    print('No attachments found.')
                for attachment in attachments:
                    if attachment['filename'].endswith('.xlsx'):
                        print(f"Found attachment: {attachment['filename']}")
                        path = os.path.join(store_dir, attachment['filename'])
                        os.mkdir(store_dir) if not os.path.exists(
                            store_dir) else None
                        with open(path, 'wb') as f:
                            f.write(base64.urlsafe_b64decode(
                                attachment['data']))
                    else:
                        print(f"Skipping attachment: {attachment['filename']}")
                print(f"{'-'*20}")

    except HttpError as error:
        print(f'An error occurred: {error}')


if __name__ == '__main__':
    main()
