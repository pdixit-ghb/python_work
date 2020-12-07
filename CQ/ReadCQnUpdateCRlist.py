#
# All packages import
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import os
import glob
import os.path
import win32com.client
import pandas as pd

import sys
import time
import math

from array import *
from datetime import datetime


#
# get latest file in folder
def latest_file(files):
    return max(files , key = os.path.getctime)


#
# Read from one xls and write to other xls
def ReadXLS1_WriteXLS2(xls1_name, xls1_sheet_name, xls2_name, xls2_sheet_name):
    #
    data_xls1_column_A = xlApp_xls1.Range("A2:A150")
    
    #
    xls2_ds_start_row = 2
    xls2_ds_end_row = 3000
    xls2_df = pd.read_excel(xls2_name, xls2_sheet_name, header=None, usecols="A,B,C,D,E", skiprows=xls2_ds_start_row, nrows=(xls2_ds_end_row-xls2_ds_start_row))
	
    #
    #search in datafram based on search key
    #xls2_df1 = xls2_df.loc[(xls2_df[0]==2109)]
    #print(xls2_df1)
	#print(f"PR task index = {ds_start_row + epr_df3.index[0]}")
    
    #
    #xls2_start_date_col = 11
    #xls2_end_date_col = 12
    j = 2
    
    #
    print("\n")
    print ("Started to search in Database...")
    print ("------------------------------------")
    
    for i in data_xls1_column_A:
        if (str(i)) != "None" :
            cr_id_key = int(str(i))
            xls2_df1 = xls2_df.loc[(xls2_df[0] == cr_id_key)]
            #print (f"Serch Keys = {cr_id_key}")
            #xlWorksheet_xls1.Cells(j, 2).Value = str(xls2_df1[1].values[0])
            column_C_value = xls2_df1[2].values[0]
            if (column_C_value != column_C_value) :
                column_C_value = ""

            xlWorksheet_xls1.Cells(j, 2).Value = xls2_df1[1].values[0]
            xlWorksheet_xls1.Cells(j, 3).Value = column_C_value
            xlWorksheet_xls1.Cells(j, 4).Value = xls2_df1[3].values[0]
            xlWorksheet_xls1.Cells(j, 5).Value = xls2_df1[4].values[0]
            print(str(int(xls2_df1[0].values[0])))
            #print(column_C_value)
        
        j = j+1
        
    
    return


#################################################################################
# Main Execution
#################################################################################


#
# get browser handle by driver and open URL with browser
#driver = webdriver.Firefox()
driver = webdriver.Chrome('chromedriver')
driver.get("https://clearquest.alstom.hub/cqweb/restapi/CQat/atvcm/QUERY/47363098?format=HTML&noframes=true")
print(driver.title)

#
# login to CQ to get access of the query (auto enter username and password)
time.sleep(1)
driver.switch_to.active_element
username = driver.find_element_by_name('loginId_Id')
username.send_keys("pdixit")
time.sleep(1)
password = driver.find_element_by_name('passwordId')
password.send_keys("passwd@JUN2018")

#
# auto click excel download
time.sleep(1)
connect_button = driver.find_element_by_id('loginButtonId')
connect_button.click()
time.sleep(5)
export_button = driver.find_element_by_id('dijit_form_ComboButton_1_arrow')
export_button.click()
time.sleep(5)
export_button = driver.find_element_by_id('dijit_MenuItem_35_text')
export_button.click()

time.sleep(10)

#
#close the driver
print(driver.current_url)
driver.close()

#
xlApp_xls1 	= win32com.client.Dispatch("Excel.Application")
xlApp_xls2	= win32com.client.Dispatch("Excel.Application")

#
# latest downloaded file xls
files_pattern = glob.glob('C:\\Users\\pdixit\\Downloads\\*.xls')
cq_query_output_xls = latest_file(files_pattern)
print(cq_query_output_xls)

#cr_list_xls = "C:\\Users\\pdixit\\OneDrive - Alstom\Software\\SwPM\\R13-MooN-Mw-TasksList-Test.xlsx"
cr_list_xls = "C:\\Users\\pdixit\\OneDrive - Alstom\Software\\SwPM\\R13-MooN-Mw-TasksList.xlsx"
xls1_name = cr_list_xls
xls2_name = cq_query_output_xls
xls1_sheet_name = "R13-Beta2-CRs-Live"
xls2_sheet_name = "IBM Rational ClearQuest Web"

#
xlWbook_xls2 = xlApp_xls2.Workbooks.Open(os.path.abspath(xls2_name))
xlWorksheet_xls2 = xlWbook_xls2.Worksheets(xls2_sheet_name).Select()
xlWorksheet_xls2 = xlWbook_xls2.Worksheets(xls2_sheet_name)

#
xlWbook_xls1 = xlApp_xls1.Workbooks.Open(os.path.abspath(xls1_name), ReadOnly=0)
xlWorksheet_xls1 = xlWbook_xls1.Worksheets(xls1_sheet_name).Select()
xlWorksheet_xls1 = xlWbook_xls1.Worksheets(xls1_sheet_name)

#
ReadXLS1_WriteXLS2(xls1_name, xls1_sheet_name, xls2_name, xls2_sheet_name)

#
xlWbook_xls1.Save()

#
xlWbook_xls1.Close()
xlWbook_xls2.Close()

#
xlApp_xls2.Quit()
xlApp_xls1.Quit()

#
# delet downloaded file
os.remove(cq_query_output_xls)
##################################################################################
