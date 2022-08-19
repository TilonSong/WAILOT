# Project Name: WAILOT [Wish All In List On Time]
# Author      : Tilon Song
# Created Date: 6th August 2022
# Description : An automated system that automatically send designated greetings and wishes for specific person 
#               or group with the help of Windows' Task Scheduler function.

from openpyxl import load_workbook
from datetime import datetime
import pywhatkit 
import time

data_file = 'list.xlsx' #file name of the Excel file with the data recorded

workbook = load_workbook(data_file)

worksheet = workbook['Sheet1']
all_columns = list(worksheet.columns)

cell_data_list = []

e = 0

for i in range (1, 8):
    try:
        day = datetime.today()
        t = datetime.now()
        h = t.hour
        m = t.minute 
        date = day.strftime("%m/%d") #remember to change the cell format in Excel file to Text

        e = 0

        for cell in all_columns[0:5]:  #0 represents 1st cell in the specific row and 5 represents the 6th cell in the specific row, so 0:5 means get the data from the 1st to the 6th cells of the specific row
            cell_data = cell[i].value  #1 represents the 2nd row in the sheet, so it will only get the data from the 2nd row in the sheet
            cell_data_list.append(cell_data)
            
        if (cell_data_list[1] == date):
            if (cell_data_list[2] == "WhatsApp Group"):
                pywhatkit.sendwhatmsg_to_group(str(cell_data_list[3]), str(cell_data_list[4]), h, m+1, 15, True, 3) #have to change the data type in Excel file from General to Text in order to send the messages as string
                cell_data_list.clear() #need to clear the content of the list or else it the output will just continue to show the content of the first row of data only
                time.sleep(5) #need to delay because the WhatsApp web action will be completed before the mentioned time, causing the loop to be stucked until tomorrow
            elif (cell_data_list[2] == "WhatsApp Contact"):
                pywhatkit.sendwhatmsg(str(cell_data_list[3]), str(cell_data_list[4]), h, m+1, 15, True, 3)
                cell_data_list.clear() #need to clear the content of the list or else it the output will just continue to show the content of the first row of data only
                time.sleep(5)
    except: #error will occur if the loop have less than 15 seconds to run as pywhatkit need 15 seconds to open WhatsApp web, so use try except in Python to prevent the system from crashing
        e += 1
        time.sleep(5)
        if (e > 3): #if error occurs for 3 consecutive times, it means that the system had reached the end of the available data in the Excel file and hence the system can be close
            exit()

