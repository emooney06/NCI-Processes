from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google import auth
from oauth2client.client import GoogleCredentials
import pandas as pd
import time

#authorize the google drive access
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

timestr = time.strftime("%Y%m%d-%H%M")
file1 = drive.CreateFile({'id': '1U362h3YgTplBN6uNIWXQV9dq4i0Z7VrY'})
file1.SetContentString(timestr)
file1.Upload() # Files.insert()

downloaded = drive.CreateFile({'id':"1U362h3YgTplBN6uNIWXQV9dq4i0Z7VrY"})   # replace the id with id of file you want to access
content = downloaded.GetContentFile('Hello.csv')        # replace the file name with your file

content = pd.read_csv('Hello.csv')
content


#timestr = time.strftime("%Y%m%d-%H%M")
#file1 = drive.CreateFile({'title': 'Hello.csv'})
#file1.SetContentString(timestr)
#file1.Upload() # Files.insert()