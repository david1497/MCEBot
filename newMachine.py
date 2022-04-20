#%%
import pyautogui as pg
import subprocess
import pandas as pd
import webbrowser
import time
from codecarbon import EmissionsTracker
#from helper_xlsx import *

tracker = EmissionsTracker()
tracker.start()

#%%
def filter_selections(tabs, filters):
    for y in range(0, len(filters)):
        if filters.Attribute[y] == "Scroll":
            print("-----------------------------------")
            a = filters.X[y]
            if filters.STEP[y] == 'vscroll':
                pg.scroll(int(a))
                print(f"Scrolled {filters.X[y]}")
            else:
                print("\n I am in that else statement \n")
                pg.dragTo(int(filters.X[y]), int(filters.Y[y]), button='left')  
        else:
            print(f"{filters.Name[y]}, {filters.X[y]}, {filters.Y[y]}, button={filters.STEP[y]}")
            pg.click(filters.X[y], filters.Y[y], button=filters.STEP[y])
            time.sleep(1.5)
            print('using updated version of this method 10')
        time.sleep(1)
    time.sleep(2)


#%%
def filter_selections_demo(tabs, filters):
    for i in range(0, len(filters)):
        if filters.Attribute[i] == "Scroll":
            a = filters.X[i]
            if filters.STEP[i] == 'vscroll':
                pg.scroll(int(a))
                print(f"Scrolled {filters.X[i]}")
            else:
                pg.dragTo(int(filters.X[i]), int(filters.Y[i]), button='left')  
        else:
            posX = filters.X[i]
            posY = filters.Y[i]
            if filters.Single[i] == "N" and filters.AlwaysOneSelected[i] == "N":
                for w in range(0, int(filters.Nr_Elements[i])):
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    time.sleep(2)
                    print(f"{filters.Name[i]}, {posX}, {posY}, button={filters.STEP[i]}")
                    pg.click(posX, posY, button=filters.STEP[i])
                    time.sleep(2)
                    print(f" \t {filters.Name[i]}, {filters.X[i]}, {filters.Y[i]}, button={filters.STEP[i]}")
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    time.sleep(2)
                    posX += filters.DistanceX[i]
                    posY += filters.DistanceY[i]
            elif filters.Single[i] == "N" and filters.AlwaysOneSelected[i] == "Y":
                for w in range(0, int(filters.Nr_Elements[i])):
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    print(f" \t {filters.Name[i]}, {filters.X[i]}, {filters.Y[i]}, button={filters.STEP[i]}")                   
                    time.sleep(2)
                    print(f"{filters.Name[i]}, {posX}, {posY}, button={filters.STEP[i]}")
                    pg.click(posX, posY, button=filters.STEP[i])
                    time.sleep(2)
                    posX += filters.DistanceX[i]
                    posY += filters.DistanceY[i]
            else:
                print(f"{filters.Name[i]} | Single item")
                pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
            time.sleep(1.5)
            print("-----------------------------")
        time.sleep(1)
    time.sleep(2)


# %% Reading Coordinates file C:/Users/A374410/Desktop/
file = 'UpdatedCoordinates.xlsx'
tabs = pd.read_excel(file, sheet_name='Tabs')

#%% Starting the browser
webbrowser.open('https://qlikview-qa.srv.volvo.com/QvAJAXZfc/opendoc.htm?document=gtt-apmcd%5Cmce.qvw&lang=en-US&host=QVS%40QA')
#start_recording()
time.sleep(15)
pg.moveTo(800, 800)
with pg.hold('ctrl'):
    pg.press('-')
    print("Zoomed out")

# going through each xlsx sheet with filters and calling filter_selections to imitate user's behavious
for y in range(0, len(tabs)):
    print(tabs.Name[y])
    filters = pd.read_excel(file, sheet_name=tabs.Name[y])
    print(f"reading the {tabs.Name[y]}")
    print("===============================")
    pg.click(tabs.X[y], tabs.Y[y])
    filter_selections_demo(tabs.Name[y], filters)

#end_recording()
# %%
emission: float = tracker.stop()
print(f"Emissions: {emission} kg")

#%% Starting and Ending the recordings
def start_recording():
    pg.hotkey('ctrl', 'shift', 'I')
    time.sleep(15)
    pg.hotkey('alt', 'tab')
    pg.hotkey('f2')
    print("Recording started.")

def end_recording():
    pg.hotkey('ctrl', 'shift', 'I')
    time.sleep(5)
    pg.hotkey('f1')
    print("Recording ended.")
    time.sleep(2)
    pg.hotkey('alt', 'f4')
    print("Closing the popup window and the recording app.")

# %%
vide_recordings_path = f"//Vcn.ds.volvo.net/cli-hm/hm1163/a374410/My Documents/My Videos/RecForth"
# %% Reading from XLSX

a = get_from_xlsx()
a = a.rename(columns={'REASON':'REASON', 'Actual YTD': 'ActualYTD', 'YEF':'YEF', 'Target 2021':'Target2021', 'Actual YTD % ':'ActualYTD%', 'YEF % ':'YEF%',
       'Target 2021.1':'Target2021%', ' ':'SIGN', 'Gap to 2021 Target ':'GapTo2021Target'})

main_table_df = pd.DataFrame()
main_table_df = main_table_df.append(a, ignore_index=True)

/** 
*! This is a warning comment
** This is a simple comments
*? This is a question comment
* TODO This is a TODO comment 
*/ 


#%%
import pandas as pd
file = '911cc4cc3a954da8ac6865ba054d1a2a.xlsx' # xlsx with the coordinates Add from Ses
commodity = pd.read(file)
# %%
