from tkinter import *
from PIL import Image,ImageTk #pip install pillow
from tkinter import messagebox, filedialog
from pymysql import *
import pandas as pd
import pymysql
import mailing_fun
from validate_email import validate_email

connection = pymysql.connect(
    host="localhost",
    user="root",
    passwd="",
    database="multiple_mail"
)
cursor = connection.cursor()

class email:
    def __init__(self,root):
        self.root = root
        self.root.title("Multiple Mailing System")
        self.root.geometry("1200x700+350+150")
        self.root.resizable(False,False)
        self.root.config(bg="#F5F5F6")

        #--icons--
        
        img1 = Image.open("C:\pythonpr\py_imgs\icon_1.jpg")
        mail_icon = img1.resize((70,70))
        self.image_icon1 = ImageTk.PhotoImage(mail_icon)

        img2 = Image.open("C:\pythonpr\py_imgs\setting.jpg")
        setting_icon = img2.resize((50,50))
        self.image_icon2 = ImageTk.PhotoImage(setting_icon)
        

        #--body--
        
        self.label = Label(self.root, text="Multiple Mailing System", image=self.image_icon1, compound="left", padx="10" ,bg="#F5F5F6", fg="black", font=("times new roman",32,"bold")).place(x=0,y=10,relwidth=1)

        self.settings_btn = Button(self.root, image=self.image_icon2, bg="#F5F5F6", activebackground="#F5F5F6", bd=0, command=self.settings_window).place(x=1130,y=630)

        #----

        self.optionvar=StringVar()
        self.label_note=Label(self.root, text="Note: For sending multiple mails use excel file...",font=("times new roman",14,"bold"), bg="yellow", width="1000").place(x=0,y=100, relwidth=1)
        self.single_option=Radiobutton(self.root, text="Single",value="single", command=self.check_disability ,variable=self.optionvar,bg="#F5F5F6",activebackground="#F5F5F6", font=("times new roman",20,"bold"))
        self.single_option.place(x=100,y=200)
        self.multiple_option=Radiobutton(self.root, text="Multiple", value="multiple", command=self.check_disability , variable=self.optionvar,bg="#F5F5F6", activebackground="#F5F5F6", font=("times new roman",20,"bold"))
        self.multiple_option.place(x=250,y=200)
        self.optionvar.set("single")

        #----

        self.label_email=Label(self.root, text="To (Email address):",font=("times new roman",18), bg="#F5F5F6").place(x=100,y=260)
        self.label_subject=Label(self.root, text="Subject:",font=("times new roman",18), bg="#F5F5F6").place(x=100,y=320)
        self.label_message=Label(self.root, text="Message:",font=("times new roman",18), bg="#F5F5F6").place(x=100,y=380)

        self.get_email=Entry(self.root, font=("arial",16),bg="lightyellow")
        self.get_email.place(x=300, y=260, width=260, height=30)
        self.get_subject=Entry(self.root, font=("arial",16),bg="lightyellow")
        self.get_subject.place(x=300, y=320, width=500, height=30)
        self.get_content=Text(self.root, font=("arial",16),bg="lightyellow")
        self.get_content.place(x=300, y=380, width=800, height=100)

        self.btn_browse = Button(self.root, text="BROWSE", font=("times new roman",15,"bold"), bg="lightblue", activebackground="lightblue", cursor="hand2", command=self.browse)
        self.btn_browse.place(x=600, y=253, width=130, height=40)
        self.btn_browse.config(state=DISABLED)
        self.btn_send = Button(self.root, text="SEND", font=("times new roman",18,"bold"), bg="lightgreen", activebackground="lightgreen", cursor="hand2", command=self.send).place(x=900, y=500, width=130, height=70)
        self.btn_clear = Button(self.root, text="CLEAR", font=("times new roman",17,"bold"), bg="grey", activebackground="grey", cursor="hand2",command=self.clear).place(x=740, y=500, width=130, height=70)

        #----status----

        self.label_total=Label(self.root, font=("times new roman",28), bg="#F5F5F6", fg="green")
        self.label_total.place(x=30,y=620)
        self.label_sent=Label(self.root, font=("times new roman",28), bg="#F5F5F6", fg="purple")
        self.label_sent.place(x=250,y=620)
        self.label_remaining=Label(self.root, font=("times new roman",28), bg="#F5F5F6", fg = "purple")
        self.label_remaining.place(x=430,y=620)
        self.label_failed=Label(self.root, font=("times new roman",28), bg="#F5F5F6", fg="red")
        self.label_failed.place(x=670,y=620)

        self.check_login()

    def send(self):
        content=len(self.get_content.get(1.0,END))
        if(self.get_email.get() == "" or self.get_subject.get()=="" or content == 1):
            messagebox.showerror("ERROR!","please enter valid input", parent=self.root)
        else:
            self.check_login()
            try:
                if self.optionvar.get()=="single":
                    self.valididty(self.get_email.get())
                    if self.valid_email == True:
                        status=mailing_fun.mail_sent(self.get_email.get(), self.from_logined, self.get_subject.get(), self.get_content.get(1.0,END), self.pwd_logined )
                    if status == "success":
                        messagebox.showinfo("SUCCESSFUL!","email has been successfully sent", parent=self.root)
                    else:
                        messagebox.showerror("Error!","Communication Failed!! Mail has not been send please try again")
            except:
                messagebox.showerror("Error!","Communication Failed! Mail has not been send please try again")

            if self.optionvar.get()=="multiple":
                self.s_count=0
                self.f_count=0
                #
                for i in self.emails:
                    self.valididty(i)
                    if self.valid_email == True:
                        status = mailing_fun.mail_sent(i, self.from_logined, self.get_subject.get(), self.get_content.get(1.0,END), self.pwd_logined )
                    else:
                        # self.f_count+=1
                        status = "failed"
                    if status == "success":
                        self.s_count+=1
                    if status == "failed":
                        self.f_count+=1
                    self.status()
                messagebox.showinfo("Success!","email has been successfully sent", parent= self.root)
        
        self.check_disability()

    def status(self):
        self.label_total.config(text="Total: "+str(len(self.emails)))
        self.label_remaining.config(text="Remaining: "+str(len(self.emails)-(self.s_count + self.f_count)))
        self.label_sent.config(text="Sent: "+str(self.s_count))
        self.label_failed.config(text="Failed: "+str(self.f_count))
        self.label_total.update()
        self.label_remaining.update()
        self.label_sent.update()
        self.label_failed.update()

    def clear(self):
        self.get_email.config(state=NORMAL)
        self.get_email.delete(0,END)
        self.get_email.config(state='readonly')
        self.get_subject.delete(0,END)
        
        self.get_content.delete(1.0,END)
        self.label_failed.config(text="")
        self.label_sent.config(text="")
        self.label_total.config(text="")
        self.label_remaining.config(text="")
        #

   
    
    def check_disability(self):
        if(self.optionvar.get()=="single"):
            self.btn_browse.config(state=DISABLED)
            self.get_email.delete(0,END)
            self.clear()
            self.get_email.config(state=NORMAL)
            
        else:
            self.clear()
            self.btn_browse.config(state=NORMAL)
            self.get_email.delete(0,END)
            self.get_email.config(state='readonly')
        
         
    def browse(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel file", filetypes=(("All Types","*.*"),("Excel Files",".xlsx")))
        if op != NULL:
            importfile = pd.read_excel(op.name , engine='openpyxl')

            if 'Email' in importfile.columns:
                self.emails = list(importfile['Email'])
                temp=[]
                for i in self.emails:
                    if pd.isnull(i) == False:
                        temp.append(i)
                self.emails = temp 
                filename = op.name
                if len(self.emails) >0:
                    self.get_email.config(state=NORMAL)
                    self.get_email.insert(0,filename)
                    self.get_email.config(state='readonly')
                    self.label_total.config(text="Total: "+str(len(self.emails)))
                    self.label_remaining.config(text="Remaining: ")
                    self.label_sent.config(text="Sent: ")
                    self.label_failed.config(text="Failed: ")
                else:
                    messagebox.showerror("Error", "Please enter emails", parent = self.root)
            else:
                messagebox.showerror("Error","'Email' column not present", parent = self.root)
        else:
            messagebox.showerror("Error","File nor selected", parent = self.root)
    
    def valididty(self, ae):
        is_valid = validate_email(
        email_address= ae,
        check_format=True,
        
        check_smtp=True,
        
        smtp_helo_host='my.host.name',
        )   
        self.valid_email =  is_valid


    def check_login(self):
        a="select count(*) where exists (select * from userdata)"
        if a < "1":
            useremail = self.get_from.get()
            pwd = self.get_from_pwd.get()
            userdata = (useremail,pwd)
            data = "insert into userdata(email, pwd) values(%s, %s)"
            cursor.execute(data,userdata)
            connection.commit()
            
        
        #self.login_info = []
        retrive1 = "select * from userdata"
        cursor.execute(retrive1)
        connection.commit()
        
        for i in cursor.fetchall():
            self.from_logined = str(i[1])
            self.pwd_logined = str(i[2])
            # self.login_info.append( [str(i[1]),str(i[2])] )
        

        # self.from_logined = self.login_info[0][1]
        # self.pwd_logined = self.login_info[0][1]


    #---------settings window----------------------

    def settings_window(self):
        self.check_login()
        self.root2 = Toplevel()
        self.root2.title("credential settings")
        self.root2.geometry("900x400+510+250")
        self.root2.focus_force()
        self.root2.grab_set()
        
        

        #-----

        self.label = Label(self.root2, text="Credential Settings", image=self.image_icon2, compound="left", padx="10" ,bg="#F5F5F6", fg="black", font=("times new roman",32,"bold")).place(x=0,y=10,relwidth=1)
        self.label_instruction=Label(self.root2, text="Enter you email from which you want all the emails sent from",font=("times new roman",14,"bold"), bg="green", width="1000").place(x=0,y=90, relwidth=1)

        self.label_from=Label(self.root2, text="Your Email:",font=("times new roman",18), bg="#F5F5F6").place(x=30,y=150)
        self.label_from_pwd=Label(self.root2, text="Password: ",font=("times new roman",18), bg="#F5F5F6").place(x=30,y=210)

        self.get_from=Entry(self.root2, font=("arial",16),bg="lightyellow")
        self.get_from.place(x=230, y=150, width=280, height=30)
        self.get_from_pwd=Entry(self.root2, font=("arial",16),bg="lightyellow")
        self.get_from_pwd.place(x=230, y=210, width=200, height=30)

        
        self.btn_sett_save = Button(self.root2, text="SAVE", font=("times new roman",18,"bold"), bg="lightgreen", activebackground="lightgreen", cursor="hand2", command=self.save).place(x=430, y=300, width=100, height=60)
        self.btn_sett_clear = Button(self.root2, text="CLEAR", font=("times new roman",17,"bold"), bg="grey", activebackground="grey", cursor="hand2", command=self.clear1).place(x=320, y=300, width=100, height=60)

        self.get_from.insert(0,self.from_logined)
        self.get_from_pwd.insert(0,self.pwd_logined)

    
    def clear1(self):
        self.get_from.delete(0,END)
        self.get_from_pwd.delete(0,END)

    def save(self):
        clear_table="truncate table userdata"
        cursor.execute(clear_table)
        connection.commit()
        if(self.get_from.get() == "" or self.get_from_pwd.get()==""):
            messagebox.showerror("ERROR!","please enter valid input", parent=self.root2)
        else:
            # f=open('C:\pythonpr\mailuserdata','w')
            # f.write(self.get_from.get()+", "+self.get_from_pwd.get())
            # f.close()
            try:
                useremail = self.get_from.get()
                pwd = self.get_from_pwd.get()
                userdata = (useremail,pwd)
                data = "insert into userdata(email, pwd) values(%s, %s)"
                cursor.execute(data,userdata)
                connection.commit()
                connection.close()
                messagebox.showinfo("SUCCESSFUL!","your email has been successfully stored", parent=self.root2)
                self.check_login
            except:
                messagebox.showerror("ERROR!","Enter valid user id",parent=self.root2)
            self.check_login()

root = Tk()
objct = email(root)
root.mainloop()