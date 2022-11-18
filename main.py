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
root.geometry('400x200')
root.minsize(400,200)
root.maxsize(400,200)

def send_mails(receivers,mailBody):
    email_sender = 'aakashsrivastava013@gmail.com' # Put registered email with 2 factor authentication here
    email_receiver = receivers
    email_password = 'aiedmjzfmbmvsigm' # Put your app password here


    subject = 'TestEmail'
    body = mailBody

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
def get_latest_excel_file(mailBody):
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
            email_list = list(pd.read_excel(latest_file)['Email'])
            ls = [x for x in email_list if type(x) == str]
            send_mails(ls,mailBody)
            count = len(ls)
            msg = messagebox.showinfo("Alert",str(count)+" mails sent").pack()
        except Exception as e:
            print(e)
        else:
            f = open('db.txt','a+')
            
            f.write(latest_file_name+":"+str(count)+" Mails Sent"+'\n')
            f.close()
    else:
        msg = messagebox.showinfo("Alert","Mails already sent to latest file").pack()
    
# -----------------------------------------------------------------------------




newFileSelectLabel = Label(text='Latest file in directory:').pack()
mailBodyText = Label(text='Enter mail message:').pack()
mailBody = Entry(root).pack()
checkNewFileButton = Button(text='Send',command=get_latest_excel_file(mailBody)).pack()

root.mainloop() 


