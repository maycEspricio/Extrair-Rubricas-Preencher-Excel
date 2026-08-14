import urllib.request
import json

req = urllib.request.Request('http://localhost:5000/api/preencher/check')
with urllib.request.urlopen(req) as response:
    print(response.read().decode())
