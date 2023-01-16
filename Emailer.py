# Importing Library
import win32
import win32com.client
import os
import pandas as pd
import getpass
import datetime
from datetime import datetime
# Datetime filtering for email shooting
month = datetime.now().strftime("%b")
year = datetime.now().strftime("%Y")
user = getpass.getuser()
# Calling The Outlook Object for creating email
outlook = win32com.client.Dispatch('Outlook.Application')
mapi = outlook.GetNamespace("MAPI")
#File Operations
path = "C:\\Users\\"+user+"\\Downloads\\Abhijeet_Gode1\\File_{}{}".format(month, year+".xlsx")
Data_sheet = pd.read_excel(path, sheet_name = "A, B")#, skiprows=1
Data_sheet.reset_index(drop=True, inplace=True)
# Data sorting for the email shooting condition
Week_days = int(input('Enter the Week days: '))
Data_sheet['Column_Name'] = (Data_sheet['Column_Name'].astype(str).replace({'NaT': ''}))
emailer = Data_sheet[(Data_sheet['Column_Name'] == Week_days) & (Data_sheet['Column_Name'] == '')]
ab = emailer[['Column_Name', 'Column_Name']].dropna()#.str.split(';')
ind = ab.reset_index(drop=True)

# Creating a email with Mapi object and Win32 function
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetLast()

for i in ind['Column_Name']:
        for sub in ind['Column_Name']:
                mail = outlook.CreateItem(0)
                mail.SentOnBehalfOfName = 'From_email@xy.co'
                mail.To = i
                mail.CC = "CC_email_list@xyz.com"
                mail.Subject="Re: Weekly budget report for {0} day for {1} {2}".format(sub, month, year)
                mail.HTMLBody = f"""
                    <b>Hi,</b><br><br>
                    Thank you for the email

                    """
                mail.Save()