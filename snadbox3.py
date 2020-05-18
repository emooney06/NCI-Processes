from exchangelib import DELEGATE, Account, Credentials, Configuration


creds = Credentials(username='health\ejmooney@salud.unm.edu', password='python#1')
config = Configuration(server='HSCLink.health.unm.edu', credentials=creds)


account = Account('ejmooney@salud.unm.edu', credentials=creds, autodiscover=True)

account = Account('ejmooney@salud.unm.edu', config=config, autodiscover=False)

account = Account(primary_smtp_address='ejmooney@salud.unm.edu', config=config, autodiscover=False, access_type=DELEGATE)



#account = Account(
#    primary_smtp_address='ejmooney@salud.unm.edu',
#    autodiscover=False, 
#    config=config,
#    access_type=DELEGATE
#)

for item in account.inbox.all().order_by('-datetime_received')[:100]:
    print(item.subject, item.sender, item.datetime_received)