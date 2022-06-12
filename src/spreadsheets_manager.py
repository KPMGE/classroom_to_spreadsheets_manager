from __future__ import print_function

import os.path
import json
import numpy as np
from pprint import pprint
import requests

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

#====spreadsheet ID (You can find this ID on spreadsheet link)
SAMPLE_SPREADSHEET_ID = '1si98UZlpeTOpPvuqHYaYocnOvOnHQGonG8k9jgrWm6k'


def main():
    
    #========================================Autentication=================================#

    #initializing credentials as NULL
    creds = None
    
    #Verifying if token.json already exists
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    #Verifying credentials to get the token or create a new token
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    #========================================Data management=========================================#

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='Página1!A1:B5').execute()
                                    
        values = result.get('values', [])

        print(values)

    except HttpError as err:
        print(err)

    # =============================================================================== #

    # Another try to post data
    try:

        import get_course_works
        list = get_course_works.req2

        # declaring arrays
        line = []
        lines =[] # matrix

        for works in list:

            line = []
            line.append(works['title'])


            for students in works['submissions']:
                line.append(students['student']['name'])

            if len(line) < get_course_works.size:
                for i in range (get_course_works.size - len(line)):
                    line.append(' ')

            lines.append(line)
            

        matrix = np.array(lines, dtype=object)
        matrix = matrix.transpose()
        matrix = matrix.tolist()


        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range='Página1!A1', valueInputOption="USER_ENTERED", 
            body = {"values": matrix})

        response = result.execute()

        pprint(response)

    except HttpError as err:
        print(err)


#===========================================================================================    

if __name__ == '__main__':
    main()


# Se o codigo nao funcionar, voltar a gerar arquivo e ler os dados do arquivo 
# O codigo nao está transpondo a matriz quando é passado direto 