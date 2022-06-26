from __future__ import print_function

import os.path
import numpy as np
from pprint import pprint

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
import paintCell

class SpreadSheet: 
    def __init__(self, credentials_file, token_file):
        self.credentials_file = credentials_file
        self.token_file = token_file

    def __update_cell(self, line, column, sheet_id):
        paintCell.body["requests"][0]["updateCells"]["start"]["rowIndex"] = line
        paintCell.body["requests"][0]["updateCells"]["start"]["columnIndex"] = column

        # request to updateCells
        request = self.sheet.batchUpdate(spreadsheetId=sheet_id, body=paintCell.body)
        response = request.execute()
        pprint(response)

    def save_course_works(self, sheet_id, course_works, all_students): 
        # declaring arrays
        line  = [] # single line
        lines = [] # matrix

        for j, works in enumerate(course_works):
            # creating new line to new work
            line = []
            line.append(works['title'])

            for i, student in enumerate(works['submissions']):
                line.append(student['student']['name'])             #appending student name to line
                
                if(student["late"]):
                    self.__update_cell((i+1),j,sheet_id)

            # filling void fields to transpose matrix
            if len(line) < len(all_students):
                for i in range (len(all_students) - len(line)):
                    line.append(' ')  

            # appending new line to matrix
            lines.append(line)
            
        # using numpy to transpose matrix
        matrix = np.array(lines, dtype=object)
        matrix = matrix.transpose()
        matrix = matrix.tolist()

        # putting students names in to sheet
        result = self.sheet.values().update(spreadsheetId=sheet_id,
            range='Página1!A1', valueInputOption="USER_ENTERED", 
            body = {"values": matrix})

        response = result.execute()

        pprint(response)

    def authorize(self): 
        creds = None

        #Verifying if token.json already exists
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file(self.token_file, SCOPES)

        #Verifying credentials to get the token or create a new token
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())

        try:    
            # connect to google sheets
            service = build('sheets', 'v4', credentials=creds)
            self.sheet  = service.spreadsheets()

        except HttpError as err:
            print(err)
