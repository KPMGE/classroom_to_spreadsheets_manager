import requests
import json
from spreadsheet import SpreadSheet

response = requests.get('http://127.0.0.1:3333/api/courses').text
courses = json.loads(response)


print('Which course do you wanna save? ')
for i, course in enumerate(courses):
    print(f'{i} - {course["title"]}')


choice = int(input('> '))
spreadsheet_id = str(input('Enter the id of your spreadsheet: '))


chosen_course_id = courses[choice]['id']
print('getting course works...')
response = requests.get(f'http://127.0.0.1:3333/api/course-works/{chosen_course_id}').text
course_works = json.loads(response)


print('getting all students...')
response = requests.get(f'http://127.0.0.1:3333/api/students/{chosen_course_id}').text
all_students = json.loads(response)


spreadsheet = SpreadSheet('client_secret.json', 'token.json', spreadsheet_id)
spreadsheet.authorize()
print('saving course works into your spreadsheet...')
spreadsheet.save_course_works(course_works, all_students)

spreadsheet.list_all_students(all_students, course_works)


print('all done!')
