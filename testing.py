import win32com.client
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "I AM SUBJECT!!"
newMail.Body = "I AM IN THE BODY\nSO AM I!!!"
newMail.To = "ejmooney@salud.unm.edu"
#newMail.CC = "moreaddresses here"
#newMail.BCC = "address"
#attachment1 = "Path to attachment no. 1"
#attachment2 = "Path to attachment no. 2"
#newMail.Attachments.Add(attachment1)
#newMail.Attachments.Add(attachment2)
newMail.display()
newMail.Send()

import pandas as pd

df = pd.DataFrame({'test1': 12340.0, 'test2':12345.0}, [0])
num = 12340.0
df['test1'] = df['test1'].astype(str)
df['test1'] = df['test1'].str[:5]