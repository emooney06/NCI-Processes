import sqlalchemy
import os
import pandas as pd
import datetime as dt
from my_functions import max_pd_display

max_pd_display()

engine = sqlalchemy.create_engine('mssql+pymssql://unmmg-epide/PULSE')

sql_nicu = " \
    SELECT Distinct \
    Unit_Clinical_DESC \
    FROM ndnqi_raw_export "
    
sql_nicu = " \
    SELECT * \
    FROM ndnqi_raw_export \
    where Unit_Clinical_DESC = 'Newborn ICU (12455)'"
    

x_df = pd.read_sql(sql_nicu, con=engine)
x_df