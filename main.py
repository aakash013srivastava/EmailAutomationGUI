from tkinter import *
from tkinter import messagebox
import os
import pandas as pd
import glob

from email.message import EmailMessage
import ssl
import smtplib

global count

root = Tk()
root.title('Send Email to excel file recipients')
root.geometry('500x200')
root.minsize(500,200)
root.maxsize(500,200)

def send_mails(receivers,messages):
    print("in sendMails")
    email_sender = 'aakashsrivastava013@gmail.com' # Put registered email with 2 factor authentication here
    email_password = 'aiedmjzfmbmvsigm' # Put your app password here

    for i in range(0,len(receivers)):

        email_receiver = receivers[i]
        subject = 'ForwardedEmail'
        body = messages[i]

        em = EmailMessage()
        em['From'] = email_sender
        em['To'] = email_receiver
        em['Subject'] = subject
        em.set_content(body)

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
            smtp.login(email_sender,email_password)
            smtp.sendmail(email_sender,email_receiver,em.as_string())

# -------------------------------------------------------------------------
#------------------------------------------------------------------------------
def get_latest_excel_file():
    username = str(os.getlogin())
    path1 = '/home/'+username+'/Downloads/*.xlsx'
    list_of_files = glob.glob(path1) #Get all excel python files with path
    files = os.listdir(f"/home/"+username+"/Downloads/")#Get all files
    excel_files = []
    for x in list_of_files:
        if os.path.splitext(x)[1] == '.xlsx':
            excel_files.append(x)
    
    latest_file = max(excel_files, key=os.path.getctime)
    
    latest_file_name = latest_file.split('/')[-1]
    f = open('db.txt','r+')
    lines = f.readlines()
    
    if len(lines)>0:
        for x in lines:
            file_name_in_db = x.split(':')[0]

            if  latest_file_name == file_name_in_db:
                flag=True
            else:
                flag=False
    else:
        flag = False
        f.close()
    if not flag:  
        try:
            recipient_list = list(pd.read_excel(latest_file)['Email'])
            message_text_list = list(pd.read_excel(latest_file)['Message'])
            send_mails(recipient_list,message_text_list)
            message_alert = str(pd.read_excel(latest_file)['Message'].count())+" mails sent"
            msg = messagebox.showinfo("Alert",message_alert)
            msg.pack()
            f = open('db.txt','a+')
            
            f.write(latest_file_name+":"+str(pd.read_excel(latest_file)['Message'].count())+" Mails Sent"+'\n')
            f.close()
        except Exception as e:
            print((e))
        
            
    else:
        print("executed")
        msg = messagebox.showinfo("AlertAlreadySent","Mails already sent to latest file").pack()
    
# -----------------------------------------------------------------------------




newFileSelectLabel = Label(text='Email messages to addresses from latest excel file in downloads folder').pack()
# mailBodyText = Label(text='Enter mail message:').pack()
# mailBody = Text(root,height=5,width=20).pack()
checkNewFileButton = Button(text='Send',command=get_latest_excel_file).pack()


root.mainloop() 


