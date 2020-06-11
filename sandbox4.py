import requests

url = "https://kronosweb.health.unm.edu/wfc/portal"
headers = {'Content-Type': 'text/xml'}
data = """<Kronos_WFC version = "1.0">
              <Request Object="System" Action="Logon" Username="ejmooney" Password="python#1"/>
          </Kronos_WFC>"""

# Login to Kronos and print response
session = requests.Session()  # preserve login cookies across requests
response = session.post(url, data=data, headers=headers)
print(response.text)