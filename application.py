import email_sender
import pandas as pd
from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as tkMessageBox
from openpyxl import load_workbook


def add_destroy():
    window.destroy()
    
def send_destroy():
    send_mail_window.destroy()
    
def clear():
    age.set("")
    name.set("")
    email.set("")
    
def send_clear():
    subject.set("")
    body.set("")
    attachment.set("")
    
def send_welcome_mail(name,age):
    body_1=f'''Dear {name},\n Hope you are doing fine.We are happy to share that you are selected for further onboarding process.I know that you are {age} years and eager to join ACCENTURE. We have attached a document for your verification.SO, We will share our future updates with you.Keep in touch with us.\n\n\n\nRegards,Accenture India '''
    sub_1="Updates Regarding Recruitment"
    email_sender.welcome_email(sub_1,body_1,'Yes')
    
def writedata():
    
    email_user=email.get()
    age_user=age.get()
    name_user=name.get()
    
    add='email_details.xlsx'
    cust_data=pd.read_excel(add)
    
    f=1
    for i in cust_data.get('EMAIL'):
        if i==email_user:
            print("found the data")
            f=0
            break
        
    if f==1:
        new=[[name_user,age_user,email_user]]
        df=load_workbook(add)
        sheet=df.active
        
        for row in new:
            sheet.append(row)
        
        df.save(add)
        send_welcome_mail(name_user,age_user)
        
    
def submit():   
    if '@' not in email.get() or len(email.get())<9:
        result=tkMessageBox.showerror("Warning","Enter valid email")
        sub_lbl.config(command=lambda: email.set(''))
        
    elif name.get()=='' or age.get()=='' or email.get()=='':
        result=tkMessageBox.showerror("Warning","Enter all provided fields")
    
    else:
        succ_lbl=Label(window,text='Value added successfully to the database',font=('Arial',25,'bold'),foreground='Green',bg='grey')
        succ_lbl.place(x=80,y=550)
        clear_but=Button(window,text='CLEAR',font=('Arial',25,'bold'),command=clear,foreground='black',bg='white',border=5,borderwidth=5)
        clear_but.pack(side=TOP)
        clear_but.place(x=300,y=650)
        writedata()
        
def Addnew():
    global window,email_box,age_box,sub_lbl
    
    window=Toplevel()
    window.title('Adding new Member')
    window.config(bg='Grey')
    window.resizable(TRUE,TRUE)
    window.geometry('1000x750')
    
    clear()
    
    title=Label(window,text='Adding New Customers',font=("Comic Sans MS", 40, "bold"),background='Black',foreground='White')
    title.pack(fill=X)
    fill=Label(window,text='Fill all the details below to add a customer:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    fill.place(x=80,y=100)
    
    name_lbl=Label(window,text='Name:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    name_lbl.place(x=80,y=180)
    age_lbl=Label(window,text='Age:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    age_lbl.place(x=80,y=260)   
    email_lbl=Label(window,text='Email:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    email_lbl.place(x=80,y=340)
    
    name_box=Entry(window,textvariable=name,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    name_box.place(x=200,y=185)
    age_box=Entry(window,textvariable=age,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    age_box.place(x=200,y=265)
    email_box=Entry(window,textvariable=email,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    email_box.place(x=200,y=345)
      
    exit_lbl=Button(window,text='Exit',font=("Comic Sans MS", 25, "bold"),command=add_destroy,foreground='Black',bg='white',borderwidth=15,border=10)
    exit_lbl.pack(side=LEFT)
    exit_lbl.place(x=80,y=450)
    sub_lbl=Button(window,text='Submit',font=("Comic Sans MS", 25, "bold"),command=submit,foreground='Black',bg='white',borderwidth=15,border=10)
    sub_lbl.pack(side=RIGHT)
    sub_lbl.place(x=560,y=450)  
    
    
    
def send_final_email(subject,body,attachment):
    email_sender.send_mail(subject,body,attachment)
    
def check_and_send_final_mail():
    if subject.get()=='' or body.get()=='' or attachment.get()=='':
        result=tkMessageBox.showerror("Warning","Enter all provided fields")
    else:
        send_final_email(subject.get(),body.get(),attachment.get())
        succ_lbl=Label(send_mail_window,text='Email Sent Successfully',font=('Arial',25,'bold'),foreground='Green',bg='grey')
        succ_lbl.place(x=80,y=750)
        clear_but=Button(send_mail_window,text='CLEAR',font=('Arial',25,'bold'),command=send_clear,foreground='black',bg='white',border=5,borderwidth=5)
        clear_but.pack(side=TOP)
        clear_but.place(x=300,y=800)
        

def sendmail():
    global send_mail_window
    
    send_mail_window=Toplevel()
    send_mail_window.title('Send Mail')
    send_mail_window.config(bg='Grey')
    send_mail_window.resizable(True,True)
    send_mail_window.geometry('1000x900')
    
    send_clear()
    
    title=Label(send_mail_window,text='Sending Mail to Customers',font=("Comic Sans MS", 40, "bold"),background='Black',foreground='White')
    title.pack(fill=X)
    fill=Label(send_mail_window,text='Fill all the details below to Send a Mail to customer:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    fill.place(x=80,y=100)
    
    sub_lbl=Label(send_mail_window,text='Subject:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    sub_lbl.place(x=80,y=180)
    age_lbl=Label(send_mail_window,text='Body:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    age_lbl.place(x=80,y=270)   
    attachment_lbl=Label(send_mail_window,text='Attachment:',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    attachment_lbl.place(x=80,y=480)
    
    sub_box=Entry(send_mail_window,textvariable=subject,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    sub_box.place(x=280,y=180,height=60)
    body_box=Entry(send_mail_window,textvariable=body,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    body_box.place(x=280,y=265,height=200)
    attachment_box=Entry(send_mail_window,textvariable=attachment,font=('Comic Sans MS', 15,'bold'),foreground='black',background='White',width=40,border=5,borderwidth=5)
    attachment_box.place(x=280,y=480)

    exit_lbl=Button(send_mail_window,text='Exit',font=("Comic Sans MS", 25, "bold"),command=send_destroy,foreground='Black',bg='white',borderwidth=15,border=10)
    exit_lbl.pack(side=LEFT)
    exit_lbl.place(x=80,y=600)
    sub_lbl=Button(send_mail_window,text='Send Mail',font=("Comic Sans MS", 25, "bold"),command=check_and_send_final_mail,foreground='Black',bg='white',borderwidth=15,border=10)
    sub_lbl.pack(side=RIGHT)
    sub_lbl.place(x=560,y=600)
    
def credit():
    credit_window=Toplevel()
    credit_window.title('Credits')
    credit_window.config(bg='grey')
    credit_window.resizable(True,True)
    credit_window.geometry('1000x600')
    
    content="Developer,UI/UX designer,Tester,Product Manager and Product Owner \n is "
    detail_lbl=Label(credit_window,text=content,font=("Comic Sans MS", 20, "bold"),foreground='Blue',bg='grey')
    detail_lbl.pack(side=TOP,fill=BOTH)
    detail_lbl.place(x=10,y=70)
    credit_lbl=Label(credit_window,text='SENTHILNATHAN',font=("Comic Sans MS", 30, "bold"),foreground='White',bg='grey')
    credit_lbl.pack(side=TOP,fill=BOTH)
    credit_lbl.place(x=300,y=170)
    contact_lbl=Label(credit_window,text='For Details Contact: senthilnathanr2003@gmail.com',font=("Comic Sans MS", 20, "bold"),foreground='Black',bg='grey')
    contact_lbl.pack(side=BOTTOM,fill=BOTH)

root=Tk()
root.title("EMAIL SENDER")
root.resizable(True,True)
root.config(background='Grey')
root.geometry('1000x700')

age=IntVar()
name=StringVar()
email=StringVar()

subject=StringVar()
body=StringVar()
attachment=StringVar()

title=Label(root,text='Email Sender',font=("Lexend", 50,'bold'),background='Black',foreground='White')
title.pack(fill=X)
a=  Label(root, text="",bg="Grey",pady=40).pack()
add_lbl=Button(root,text='Add new member',font=("Lexend",20,'bold'),command=Addnew,height=2,width=25,pady=10,background='White',borderwidth=15,border=10,foreground='Black')
add_lbl.pack(side=TOP,fill=Y)
a=  Label(root, text="",bg="Grey",pady=20).pack()
add_lbl=Button(root,text='Send Email',font=("Lexend", 20,'bold'),height=2,command=sendmail,width=25,pady=10,background='White',borderwidth=15,border=10,foreground='Black')
add_lbl.pack(side=TOP,fill=Y)
a=  Label(root, text="",bg="Grey",pady=20).pack()
add_lbl=Button(root,text='Credits',font=("Lexend", 20,'bold'),height=2,command=credit,width=25,pady=10,background='White',borderwidth=15,border=10,foreground='Black')
add_lbl.pack(side=TOP,fill=Y)

root.mainloop()
