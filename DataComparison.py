#%%
from cgi import test
from doctest import DocFileCase
import pandas as pd
import os
import glob
import numpy as np
from datetime import datetime

from sqlalchemy import column

##! Make sure you add a variable called test_start_time so that we can ingest only the xlsx files that were downloaded after the testing was started

#%% 
# Getting the list of files
test_start_time = datetime.strptime('2022-04-19 09:00:00', '%Y-%m-%d %H:%M:%S').timestamp()
downloads_path = "../Downloads/"
list_of_xlsx_files = glob.glob(os.path.join(downloads_path, "*.xlsx"))
valid_files = []
for file in list_of_xlsx_files:
    file_time = os.path.getctime(file) # getting the file time
    #file_name = file.split("\\",)[1] # Cutting the Dowloads\\ from the file name
    if file_time > test_start_time: # only those files which were downloaded after the test was started
        valid_files.append([file, file_time])

# %%
header_names = ['REASON', 'Actual YTD', 'YEF', 'Target 2022', 'Actual YTD %', 'YEF %',
       'Target 2022', 'SIGN', 'Gap to 2022 Target']

# %%
df_valid_files = []
for file in valid_files:
    file_df = pd.read_excel(file[0])
    file_df.set_axis(header_names, axis=1, inplace=True)
    df_valid_files.append(file_df)

# %%
def _color_red_or_green(val):
    print(val)
    color = 'green' if val[2] == 0 else 'red'
    return 'color: %s' % color
# %%
differences = df_valid_files[0].compare(df_valid_files[1], keep_equal=True)
header_name_to_drop = []
new_df = pd.DataFrame()

for col in range(differences.shape[1]):
    if col % 2 == 0:
        print(header_name_to_drop)
        new_col_name = "_"+str(col)
        old_col_name = differences.columns[col][0]
        new_df[new_col_name] = np.zeros(differences.shape[0])
        if type(differences[old_col_name].self[1]) is str:
            print("Do nothing for ", old_col_name)
        else:
            header_name_to_drop.append(old_col_name)
            print("Do something for ", old_col_name)
            for cell in range(differences.shape[0]): 
                m = np.array([differences[old_col_name].self[cell], differences[old_col_name].other[cell], differences[old_col_name].other[cell]-differences[old_col_name].self[cell]], dtype=object)
                new_df[new_col_name][cell] = m
                print("Differences with ", new_col_name, " at ", cell, " should be = ", m)
#differences = differences.drop(columns=header_name_to_drop, axis=1)
#differences.columns = header_names 

#%%
new_df.style.applymap(_color_red_or_green)

# %%
