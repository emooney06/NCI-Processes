from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody, Mailbox
from pathlib import Path
from datetime import timedelta
from exchangelib import UTC_NOW
import time 
import stdiomask
  
try: 
    ad_password = stdiomask.getpass(prompt= 'Enter Active Directory Password: ', mask='*') 
except Exception as error: 
    print('ERROR', error) 


#credentials are the domain name with username and password
creds = Credentials(username='health\\ejmooney', password=ad_password)
#account configuration
config = Configuration(server='HSCLink.health.unm.edu', credentials=creds, retry_policy=FaultTolerance(max_wait=3600))
#create the instance of account class object
a = Account('ejmooney@salud.unm.edu', credentials=creds, autodiscover=True)
#define and create the auto_rules folder
auto_folder = a.root / 'Top of Information Store' / 'auto_rules'

##list of file name text keywords in attachments to save
#attachment_to_save = ['#8940', 'random_rule_check']

#generate a time difference variable for the recency of hours that messages have arrived
#since = UTC_NOW() - timedelta(hours=4)

#define the function with inputs (name of attachment you are looking for, file path you want to save the attachment to, and name you want to call the file)
def save_attach(attach_name, path_to_save, name_to_save):
    #generate a time difference variable for the recency of hours that messages have arrived
    since = UTC_NOW() - timedelta(hours=4)
    #generate a timestamp for the print message  
    timestr = time.strftime("%Y%m%d-%H%M_")
    #filter the items in the general inbox by date received less than since variable (1 hour)
    for item in a.inbox.all().filter(datetime_received__gt=since).order_by('-datetime_received'):
        #look at each of the attachments
        for attachment in item.attachments:
            #check if the attachment is a FileAttachment (class type)
            if isinstance(attachment, FileAttachment):
                #check if the attach_name is in the file attachment
                if attach_name in attachment.name:
                    print('first print statement: ' + attachment.name)
                    #define the path and name of the file to save as
                    local_path = Path(path_to_save, name_to_save)
                    #open the path to save
                    with open(local_path, 'wb') as f:
                       #write the attachment 
                       f.write(attachment.content)
                    print('saved attachment to', local_path)
                    #move the message to the auto_folder
                    item.move(auto_folder)            
            #elif isinstance(attachment, ItemAttachment):
            #    if isinstance(attachment.item, Message):
            #        print('last print statement: ' + attachment.item, attachment.item.body)

while True:
    timestr = time.strftime("%Y%m%d-%H%M_")
    try:
        save_attach('#8940', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen','#8940 Covid Screen.xlsx')
        print(timestr) 
        print('from mail_rule - executed #8940 rule with no issues')
    except Exception as e: 
        print(timestr)
        print(e)
        print('trouble executing #8940')
    try:
        save_attach('rule_check_timestamp', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/rule_check_folder','timestamp_from_message.csv')
        print(timestr)
        print('from mail_rule - executed rule_check_timestamp rule with no issues')
    except Exception as e: 
        print(timestr)
        print(e)
        print('from mail_rule - trouble executing rule_check_timestamp')    
    #except Exception as e:
    #try:
    #    #m = Message(account=a, subject='an exception was triggered with your mail_rule module',
    #    #body='please check the mail_rule module; executing one of the save_attach functions triggered an exception',
    #    #to_recipients=[
    #    #    Mailbox(email_address='ejmooney@salud.unm.edu'),
    #    #    Mailbox(email_address='mooney.ethan@gmail.com'),
    #    #],
    #    ##cc_recipients=['carl@example.com', 'denice@example.com'],  # Simple strings work, too
    #    ##bcc_recipients=[
    #    ##    Mailbox(email_address='erik@example.com'),
    #    ##    'felicity@example.com',
    #    ##],  # Or a mix of both
    #    #)
    #    #m.send()
    #    print(e)
    #    print('exception triggered while trying to move message: ' + attach_name + timestr)
    #except Exception as e:
    #    print(e)
    #    print('2nd level exception triggered; while trying to move message and while trying to send warning email: ' + timestr)
    print('from mail_rule - sleeping for 20 minutes')
    time.sleep(1200)
