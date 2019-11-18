# %reset -f

import os
import pandas as pd
from glob import glob


# Set working direction:
os.chdir("F:\\Workshop\\01_Match8_feedback")

# Detect and import Excel:
file_name = glob('Source\\*Review*.xls')

if len(file_name) > 0:
# rename the unmodified files:    
    raw_df = [pd.read_excel(f) for f in file_name]
    
    def find_new_name(df):
        date_range = df.iloc[:, 6]
        wk_start = pd.to_datetime(date_range.min()).strftime("%Y_%b%d")
        wk_end = pd.to_datetime(date_range.max()).strftime("%b%d")
        new_name = 'RawData' + wk_start + '_to_' + wk_end
        return new_name
    
    new_names = [find_new_name(df) for df in raw_df]
    
    # Mapping new names to old names for double check:
    print('The following files will be renamed:')
    for a, b in zip(file_name, new_names):
        print(a, '-->', b)
    
    
    'Source\\' +new_names[0] + '.xls'
    
    rename_que = input('Rename all files? (y/n)')
    
    if rename_que == 'y':
        for a, b in zip(file_name, new_names):
            os.rename(a,  'Source\\' + b + '.xls')
        print('All files have been renamed.')
    else:
        print('Rename has been cancelled.')

else:
    file_name = glob('Source\\*.xls')
    print('There is no new imported raw data.')
    print('Current files list:')
    for f in file_name:
        print(f)