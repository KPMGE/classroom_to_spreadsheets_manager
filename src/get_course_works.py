import requests
import json

# Route to list courses IDs
req = requests.get('http://127.0.0.1:3333/api/courses')

# Getting course ID to insert on course works request (in this case, I have only one class)
id = req.json()[0]['id']

# Route to list course works
req2 = requests.get(f'http://127.0.0.1:3333/api/course-works/{id}').text
req2 = json.loads(req2)

# Route to get all students registered
req3 = requests.get(f'http://127.0.0.1:3333/api/students/{id}').text
size = len(req3)