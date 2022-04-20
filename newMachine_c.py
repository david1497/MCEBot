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


#%% Reading Coordinates file
#/**  
#*! Reading the tabs and their coordinates from the xlsx, reading the header and footer
#*! as well, filtering only those tabs that are chosen by the user before the start of the test. 
#*/ C:/Users/A374410/Desktop/
file = 'UpdatedCoordinates.xlsx' # xlsx with the coordinates Add from Sessions_Variables (OS)
tabs = pd.read_excel(file, sheet_name='Tabs') # reading all tabs
header = pd.read_excel(file, sheet_name='Header') # reading the header's steps
footer = pd.read_excel(file, sheet_name='Footer') # reading the footer's steps
# Getting only those tabs that user selected (1 for TRUE; 0 for FALSE)
current_tabs = tabs.where(tabs.ToTest==1).dropna().reset_index(drop=True)


#%%
def header_footer(param):
    """
    This function is applying the header and footer for each tab(Scrolling up and down, and Clear All)
    it takes only one parameter called param, the param should be a DF with coordinates and steps to be executed
    """
    for i in range(len(param)):
        if param.Attribute[i] == "Scroll":
            a = param.X[i]
            if param.STEP[i] == 'vscroll':
                pg.scroll(int(a))
                print(f"Scrolled {param.X[i]}")
            else:
                pg.dragTo(int(param.X[i]), int(param.Y[i]), button='left')
                print(f"Dragged ")
        else:
            pg.click(param.X[i], param.Y[i], button=param.STEP[i])
            print(f"Clicked on {param.Name[i]}, {param.X[i]}, {param.Y[i]}, button={param.STEP[i]}")
        time.sleep(1)


#%% 
def body_filters(filters):
    """ 
    This functin is performing the click actions on the X/Y coordinates of the screen.
    It takes one parameter called filters, the param should be a DF with the coordinates and steps to be executed.
    """
    for i in range(len(filters)): 
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
            if filters.Attribute[i] == "Scroll":
                pg.scroll(int(filters.X[i]))
                print(f"Scrolled {filters.X[i]}")
            else:
                print(f"{filters.Name[i]} | Single item")
                pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
        time.sleep(1.5)
        print("-----------------------------")
        time.sleep(2)


#%%
def body_filters1(filters):
    """ 
    This functin is performing the click actions on the X/Y coordinates of the screen.
    It takes one parameter called filters, the param should be a DF with the coordinates and steps to be executed.
    """
    for i in range(len(filters)): 
        posX = filters.X[i]
        posY = filters.Y[i]
        if filters.ClickOnce[i] == "N": # =====================> It is a filter
            for w in range(0, int(filters.Nr_Elements[i])):
                if filters.AlwaysOneSelected[i] == "N": # If it is not always one selected, you need 3 clicks
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    time.sleep(2)
                    pg.click(posX, posY, button=filters.STEP[i])
                    time.sleep(2)
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    time.sleep(2)
                else: # If it is always one selected
                    pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
                    time.sleep(2)
                    pg.click(posX, posY, button=filters.STEP[i])
                    time.sleep(2)
                posX += filters.DistanceX[i]
                posY += filters.DistanceY[i]
        elif filters.ClickOnce[i] == "Y": # ==========================> It is a button
            for w in range(0, int(filters.Nr_Elements[i])):
                print(f"{filters.Name[i]}, {posX}, {posY}, button={filters.STEP[i]}")
                pg.click(posX, posY, button=filters.STEP[i])
                time.sleep(2)
                posX += filters.DistanceX[i]
                posY += filters.DistanceY[i]
        else: # =================================> It is a scroll
            if filters.Attribute[i] == "Scroll":
                pg.scroll(int(filters.X[i]))
                print(f"Scrolled {filters.X[i]}")
            else:
                print(f"{filters.Name[i]} | Single item")
                pg.click(filters.X[i], filters.Y[i], button=filters.STEP[i])
        time.sleep(1.5)
        print("-----------------------------")
        time.sleep(2)

#%%
def filter_selections_demo(steps, include_header_footer):
    """
    Calls the appropriate functions for each tab (headless/footless or not).
    It takes two parameters: steps - a DF with the coordinates and steps; include_header_footer [True/False], 
    in case True, it will call the header_footer function before the body_filters and after it.
    """
    if include_header_footer:
        header_footer(header)
        time.sleep(1)
        print("\n >>>>>>>>>>>>>>> Done with the header")
        body_filters1(steps)
        print("\n >>>>>>>>>>>>>>> Done with the body")
        header_footer(footer)
        print("\n >>>>>>>>>>>>>>> Process completed!")
    else:
        body_filters1(steps)

#%% Starting the browser
webbrowser.open('https://qlikview-qa.srv.volvo.com/QvAJAXZfc/opendoc.htm?document=gtt-apmcd%5Cmce.qvw&lang=en-US&host=QVS%40QA')
#start_recording()
time.sleep(15)
pg.moveTo(800, 800)
with pg.hold('ctrl'):
    pg.press('-')
    print("Zoomed out")
# going through each xlsx sheet with filters and calling filter_selections to imitate user's behavious
for y in range(len(current_tabs)):
    print(current_tabs.Name[y])
    steps = pd.read_excel(file, sheet_name=current_tabs.Name[y])
    print(f"reading the {current_tabs.Name[y]}")
    print("===============================")
    pg.click(current_tabs.X[y], current_tabs.Y[y])
    if current_tabs.Name[y] == "Commodity":
        print("Headless")
        filter_selections_demo(steps, False)
    else:
        print("FullBody")
        filter_selections_demo(steps, True)

#end_recording()
# %%
emission: float = tracker.stop()
print(f"Emissions: {emission} kg")


/** 
*! This is a warning comment
** This is a simple comments
*? This is a question comment
* TODO This is a TODO comment 
*/ 