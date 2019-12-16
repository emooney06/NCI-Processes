import pandas as pd
from pathlib import Path

#test master alias path: master_alias_path = Path('C:/Users/ejmooney/Desktop/testData/MasterAlias_test.xlsx')

# Real master alias path: 
master_alias_path = Path('J:/NDNQI/MasterAliasRecord.xlsx')
master_alias = pd.read_excel(master_alias_path, 'MainAlias')

mmm_dict = {'01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun', '07': 'Jul', \
        '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'}
