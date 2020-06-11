import requests
from xml.etree import ElementTree


url = 'http://uh-krontest1.health.unm.edu/wfc/XmlService'


response = requests.get(url)

url = "https://kronosweb.health.unm.edu/wfc/XmlService"
headers = {'Content-Type': 'text/xml'}
data = """<Kronos_WFC version = "1.0">
              <Request Object="System" Action="Logon" Username="ejmooney" Password="python#1"/>
          </Kronos_WFC>"""

          https://kronosweb.health.unm.edu/wfc/applications/navigator/Navigator.do
# Login to Kronos and print response
session = requests.Session()  # preserve login cookies across requests
response = session.post(url, data=data, headers=headers)
print(response.text)





test = '''<Kronos_WFC version="1.0">
  <Request Action="Load">
    <Timesheet>
      <Employee>
        <PersonIdentity PersonNumber="100031775"/>
      </Employee>
      <Period>
        <TimeFramePeriod PeriodDateSpan="5/1/2020 - 5/30/2020"/>
      </Period>
    </Timesheet>
  </Request>
</Kronos_WFC>'''

test2 = '''<?xml version="1.0" ?>
    <Kronos_WFC version="1.0">
    <Request Action="RetrieveAllNames" >
    <VolumeDriver/>
    </Request>
    </Kronos_WFC>'''


# Login to Kronos and print response
session = requests.Session()  # preserve login cookies across requests
response = session.post(url, data=data, headers=headers)

response = session.post(url, data=test, headers=headers)


print(response.text)
response._content
tree = ElementTree.fromstring(response._content)
print(response.text)



session.request(url=url, data=test2, method='xml')
print(session)
response = session.post(url, data=data, headers=headers)