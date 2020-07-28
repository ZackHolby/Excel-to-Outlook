import os
import sys
import win32com.client as win32
import pandas as pd

#populate and create email
def Emailer(text, subject, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    attachmentPath = r'import your absolute path here' #add absolute path to attachment file or remove as necessary
    mail.Attachments.Add(attachmentPath)
    mail.Display(True)

def importExcel(filename):
    df = pd.read_excel(filename)
    return df

df = importExcel('Network Sheet.xlsx') #import the excel file to read from

for ind in df.index:
    ind = ind +10 #change this number to however many rows you want to skip, this skips first 10 people. Remove this if you want to start on first person

    headers = ['First','Last Name','Company', 'University','email address','position'] #fill with headers in excel file appropriately
    
    firstName = df[headers[0]][ind]
    lastName = df[headers[1]][ind]
    comp = df[headers[2]][ind]
    university = df[headers[3]][ind]
    email = df[headers[4]][ind]
    positionAtComp = df[headers[4]][ind]

    subjectStr = "Networking Email" #string that will populate subject frame
    bodyStr = "Hello "+firstName #add any text to bodyStr that you want in the email body
    
    Emailer(bodyStr, subjectStr ,email)
    print("Building template for "+firstName+ " "+lastName)
    input("Press Enter to build next template...")