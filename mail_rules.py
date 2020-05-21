from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody
from pathlib import Path
from datetime import timedelta
from exchangelib import UTC_NOW


creds = Credentials(username='health\\ejmooney', password='python#1')
config = Configuration(server='HSCLink.health.unm.edu', credentials=creds)

a = Account('ejmooney@salud.unm.edu', credentials=creds, autodiscover=True)

auto_folder = a.root / 'Top of Information Store' / 'auto_rules'

attachment_to_save = ['#8940', 'random_rule_check']
since = UTC_NOW() - timedelta(hours=1)

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


save_attach('#8940', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen','#8940 Covid Screen.xlsx')

save_attach('rule_check_timestamp', '//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/rule_check_folder','timestamp_from_message.csv')

#for item in a.inbox.all():
#    for attachment in item.attachments:
#        if isinstance(attachment, FileAttachment):
#            local_path = os.path.join('/tmp', attachment.name)
#            with open(local_path, 'wb') as f, attachment.fp as fp:
#                buffer = fp.read(1024)
#                while buffer:
#                    f.write(buffer)
#                    buffer = fp.read(1024)
#            print('Saved attachment to', local_path)



##a = Account('nci@salud.unm.edu', credentials=creds, autodiscover=True, access_type=DELEGATE)

##account = Account(
##    primary_smtp_address='ejmooney@salud.unm.edu',
##    autodiscover=False, 
##    config=config,
##    access_type=DELEGATE
###)

##for item in account.inbox.all().order_by('-datetime_received')[:10]:
##    print(item.subject, item.sender, item.datetime_received)


#    # It's possible to create, delete and get attachments connected to any item type:
## Process attachments on existing items. FileAttachments have a 'content' attribute
## containing the binary content of the file, and ItemAttachments have an 'item' attribute
## containing the item. The item can be a Message, CalendarItem, Task etc.

##from exchangelib import Account, FileAttachment, ItemAttachment, Message, CalendarItem, HTMLBody

##a = Account
##x=0

## Streaming downloads of file attachment is supported. This reduces memory consumption since we
## never store the full content of the file in-memory:
#for item in a.inbox.all():
#    for attachment in item.attachments:
#        if isinstance(attachment, FileAttachment):
#            local_path = os.path.join('/tmp', attachment.name)
#            with open(local_path, 'wb') as f, attachment.fp as fp:
#                buffer = fp.read(1024)
#                while buffer:
#                    f.write(buffer)
#                    buffer = fp.read(1024)
#            print('Saved attachment to', local_path)

### Streaming downloads of file attachment is supported. This reduces memory consumption since we
### never store the full content of the file in-memory:
##for item in a.inbox.all():
##    for attachment in item.attachments:
##        if isinstance(attachment, FileAttachment):
##            local_path = os.path.join('/tmp', attachment.name)
##            with open(local_path, 'wb') as f, attachment.fp as fp:
##                buffer = fp.read(1024)
##                while buffer:
##                    f.write(buffer)
##                    buffer = fp.read(1024)
##            print('Saved attachment to', local_path)



## Create a new item with an attachment
#item = Message(...)
#binary_file_content = 'Hello from unicode æøå'.encode('utf-8')  # Or read from file, BytesIO etc.
#my_file = FileAttachment(name='my_file.txt', content=binary_file_content)
#item.attach(my_file)
#my_calendar_item = CalendarItem(...)
#my_appointment = ItemAttachment(name='my_appointment', item=my_calendar_item)
#item.attach(my_appointment)
#item.save()

## Add an attachment on an existing item
#my_other_file = FileAttachment(name='my_other_file.txt', content=binary_file_content)
#item.attach(my_other_file)

## Remove the attachment again
#item.detach(my_file)

## If you want to embed an image in the item body, you can link to the file in the HTML
#message = Message(...)
#logo_filename = 'logo.png'
#with open(logo_filename, 'rb') as f:
#    my_logo = FileAttachment(name=logo_filename, content=f.read(), is_inline=True, content_id=logo_filename)
#message.attach(my_logo)
#message.body = HTMLBody('<html><body>Hello logo: <img src="cid:%s"></body></html>' % logo_filename)

## Attachments cannot be updated via EWS. In this case, you must to detach the attachment, update
## the relevant fields, and attach the updated attachment.

## Be aware that adding and deleting attachments from items that are already created in Exchange
# (items that have an item_id) will update the changekey of the item.