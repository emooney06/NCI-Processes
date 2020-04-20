
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
import re
from datetime import datetime
import sys

 
# you == recipient's email address
me = 'UNMH-Official-Communications@salud.unm.edu'
you = 'cdroberts@salud.unm.edu'

# Create message container - the correct MIME type is multipart/alternative.
msg = MIMEMultipart('alternative')
msg['Subject'] = 'test email'
msg['From'] = me
msg['To'] = you
msg['Cc'] = 'ejmooney@salud.unm.edu'

# Create the body of the message (a plain-text and an HTML version).
# text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttps://www.python.org\n"
greeting = 'where did this message go?'


html = greeting 


# Record the MIME types of both parts - text/plain and text/html.
# part1 = MIMEText(text, 'plain')
part2 = MIMEText(html, 'html')

# Attach parts into message container.
# According to RFC 2046, the last part of a multipart message, in this case
# the HTML message, is best and preferred.
#msg.attach(part1)
msg.attach(part2)

# Send the message via local SMTP server.
s = smtplib.SMTP('HSCLink.health.unm.edu')
# sendmail function takes 3 arguments: sender's address, recipient's address
# and message to send - here it is sent as one string.
s.sendmail(me, you, msg.as_string())
s.quit()
