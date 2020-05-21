from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody
from pathlib import Path
from datetime import timedelta
from exchangelib import UTC_NOW
import time 


creds = Credentials(username='health\\ejmooney', password='python#1')
config = Configuration(server='HSCLink.health.unm.edu', credentials=creds)

a = Account('ejmooney@salud.unm.edu', credentials=creds, autodiscover=True)

auto_folder = a.root / 'Top of Information Store' / 'auto_rules'

attachment_to_save = ['#8940', 'random_rule_check']
since = UTC_NOW() - timedelta(hours=1)

timestr = time.strftime("%Y%m%d-%H%M_")

def save_attach(attach_name, path_to_save, name_to_save):
    for item in a.inbox.all().filter(datetime_received__gt=since).order_by('-datetime_received'):
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                if attach_name in attachment.name:
                    print('first print statement: ' + attachment.name)
                    local_path = Path(path_to_save, name_to_save)
                    with open(local_path, 'wb') as f:
                        f.write(attachment.content)
                    print('saved attachment to', local_path)
                    item.move(auto_folder)
            elif isinstance(attachment, itemattachment):
                if isinstance(attachment.item, message):
                    print('last print statement: ' + attachment.item, attachment.item.body)

while true:
    save_attach('#8940', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen','#8940 Covid Screen.xlsx')

    save_attach('rule_check_timestamp', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/rule_check_folder','timestamp_from_message.csv')

    print(timestr + ' sleeping for 10 min')
    time.sleep(600)
