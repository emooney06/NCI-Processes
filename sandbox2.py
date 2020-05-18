import smtplib, ssl


This message is sent from Python."""

#define the file paths and file names0
data_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen')
archive_path = Path('//uh-nas/Groupshare3/ClinicalAdvisoryTeam/data_folders/8940_covid_screen/8940_archive')
file_name = 'covid_screen.xlsx'

while True:
    try:
    #create a timestamp for the archive file name
        timestr = time.strftime("%Y%m%d-%H%M_")
        #add the time stamp to the file name to create the archive file name string
        archive_file =  timestr + file_name 
        print('attempting to read the file')
        #read the file from the PI report
        df = pd.read_excel(data_path / file_name)
        #save the dataframe as an archive
        df.to_excel(archive_path / archive_file, index=False)
        # drop duplicates
        df = df.drop_duplicates()
    except:
        print('executing the except statement')
        #sleep for 24 hours
        time.sleep(20)
    print("now i'm doing the stuff")
    time.sleep(30)
