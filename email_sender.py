import smtplib
from email.message import EmailMessage
import credentials
import pandas as pd
import imghdr

user_email=credentials.email()
user_password=credentials.password()
email_list=pd.read_excel('email_details.xlsx')
size=len(email_list['EMAIL'])

def welcome_email(subject,body,attachments):
    
    send_email=email_list['EMAIL'].values[size-1]
    print(send_email)
    msg=EmailMessage()
    msg['Subject']=subject
    msg['From']=user_email
    msg['To']=send_email
    msg.set_content(body)
    if attachments!='no':
        with open('welcome_attachment.jpg','rb') as file:
            file_data=file.read()
            filetype=imghdr.what(file.name)
            file_name=file.name
        msg.add_attachment(file_data,maintype='image',subtype=filetype,filename=file_name)
            
    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
        smtp.login(user_email,user_password)
        smtp.send_message(msg)

def send_mail(subject,body,attachments):
    i=0
    while i<size:
        send_email=email_list['EMAIL'].values[i]
        print(send_email,i,size)
        msg=EmailMessage()
        msg['Subject']=subject
        msg['From']=user_email
        msg['To']=send_email
        msg.set_content(body)
        if attachments!='no':
            with open('another attachment.jpg','rb') as file:
                file_data=file.read()
                filetype=imghdr.what(file.name)
                file_name=file.name
            msg.add_attachment(file_data,maintype='image',subtype=filetype,filename=file_name)

        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
            smtp.login(user_email,user_password)
            smtp.send_message(msg)
        i+=1
        
if __name__=="__main__":
    print("This is email_sender.py file")