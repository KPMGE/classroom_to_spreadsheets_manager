from __future__ import print_function

import os.path
import numpy as np

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

#### DEPRACATED ####
# import paintCell
####################
import addSheet
import adjustColumns
import formatHeader

class SpreadSheet: 
    def __init__(self, credentials_file, token_file, spreadsheetId):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.spreadsheetId = spreadsheetId

    ############################## DEPRACATED ####################################################
    # def __update_cell(self, line, column):
    #     paintCell.body["requests"][0]["updateCells"]["start"]["rowIndex"] = line
    #     paintCell.body["requests"][0]["updateCells"]["start"]["columnIndex"] = column

    #     # request to updateCells
    #     self.sheet.batchUpdate(spreadsheetId=self.spreadsheetId, body=paintCell.body).execute()
    ##############################################################################################

    def __create_sheet(self, page_name):
        addSheet.body['requests'][0]['addSheet']['properties']['title'] = page_name
        response = self.sheet.batchUpdate(spreadsheetId=self.spreadsheetId, body=addSheet.body).execute()
        page_id = response['replies'][0]['addSheet']['properties']['sheetId']
        return  page_id

    def __adjust_columns(self, page_id): 
        adjustColumns.body['requests'][0]['autoResizeDimensions']['dimensions']['sheetId'] = page_id
        self.sheet.batchUpdate(spreadsheetId=self.spreadsheetId, body=adjustColumns.body).execute()

    def __format_header(self, size_header, page_id):
        formatHeader.body['requests'][0]['repeatCell']['range']['endColumnIndex'] = size_header
        formatHeader.body['requests'][1]['updateSheetProperties']['properties']['sheetId'] = page_id
        formatHeader.body['requests'][0]['repeatCell']['range']['sheetId'] = page_id
        self.sheet.batchUpdate(spreadsheetId=self.spreadsheetId, body=formatHeader.body).execute()

    def save_course_works(self, course_works, all_students): 
        # declaring arrays
        line  = [] # single line
        lines = [] # matrix

        for j, works in enumerate(course_works):
            # creating new line to new work
            line = []
            line.append(works['title'])

            for i, student in enumerate(works['submissions']):
                # appends student name to the matrix
                line.append('  ' + student['student']['name'] + '  ')
                
                ################################## DEPRACATED ####################################
                # if(student["late"]):
                #     self.__update_cell((i+1),j)
                ##################################################################################

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
        result = self.sheet.values().update(
            spreadsheetId=self.spreadsheetId,
            range='Página1!A1',
            valueInputOption="USER_ENTERED", 
            body = {"values": matrix})
        result.execute()

        self.__adjust_columns('0')
        self.__format_header(len(course_works), '0')


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
    

    def __get_works_amount(self, student_id, course_works):
        count = 0
        for works in course_works:
            for submission in works['submissions']:
                if student_id == submission['student']['id']:
                    count+=1
        return count


    def __calculate_percentage(self, total_exercises, returned_exercises):
        return (returned_exercises/total_exercises) * 100


    def list_all_students(self, all_students, course_works):
        matrix = []
        title = ['ALUNOS', '   EXERCÍCIOS CONCLUÍDOS   ', '    PORCENTAGEM   ']

        for student in all_students:
            name = '  ' + student['name'] + '  '
            amount = self.__get_works_amount(student['id'], course_works)
            percentage = self.__calculate_percentage(len(course_works), amount)
            percentage = round(percentage,2)    
            matrix.append([name, amount, f'{percentage}%'])
            matrix.sort()
           
        matrix.insert(0, title)
        page_name = 'RESULTADOS'
        page_id = self.__create_sheet(page_name)

        result = self.sheet.values().update(
            spreadsheetId=self.spreadsheetId,
            range=f'{page_name}!A1',
            valueInputOption="USER_ENTERED", 
            body = {"values": matrix})
        result.execute()

        self.__adjust_columns(page_id)
        self.__format_header(3, page_id)
