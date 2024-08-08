import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os
root = tkinter.Tk()
root.title("Data entry")
root.geometry("550x400")


def data_entry():

    ifaccepted=term_var.get()

    if ifaccepted =="Accepted":
        
         
        firstname_= first_entry.get()
        lastname_= last_entry.get()
        if firstname_ and lastname_:

            age_1=age_entry.get()
            contine=cont_entry.get()
            title_=title_entry.get()
        
            numcourse=num_course_entry.get()
            semester=sem_enter.get()
            regitration=reg_status_var.get()
            try:
                age_1 = int(age_1)
            except ValueError:
                tkinter.messagebox.showwarning(title="Error", message="Enter a valid age")
                return
            
            try:
                numcourse = int(numcourse)
            except ValueError:
                tkinter.messagebox.showwarning(title="Error", message="Enter a valid number of courses")
                return
            
            print("title:", title_ ,"frist name:",firstname_,"last name:",lastname_,age_1)
            print("continent",contine)
            print("no.ofcourse:",numcourse,"semester",semester)
            print("regsitration status",regitration)
            

            print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~") 

            filepath = "E:\code\project\score\data entry\data example.xlsx"
            if not os.path.exists(filepath):     #if the selected path is not exist 
                workbook = openpyxl.Workbook()
                sheet= workbook.active
                heading=["title","first name","last name","age","continent","courses","semester","registration status"]
                sheet.append(heading)
                workbook.save(filepath)

                
            workbook=openpyxl.load_workbook(filepath)
            sheet=workbook.active
            sheet.append([firstname_,lastname_,age_1,contine,title_,numcourse,semester,regitration])
            workbook.save(filepath)
        else:
            tkinter.messagebox.showwarning(title="error",message="entry your first name and last name")
    else:
        tkinter.messagebox.showwarning(title="error",message="You have not accepted the terms")
        


frame=tkinter.Frame()
user_info=tkinter.LabelFrame(frame, text="user information")

#for name
firstname=tkinter.Label(user_info,text="First name")
lastname=tkinter.Label(user_info,text="Last name")
first_entry=tkinter.Entry(user_info)
last_entry=tkinter.Entry(user_info)

# for title
title=tkinter.Label(user_info,text="Title")
title_entry=ttk.Combobox(user_info,values=["","Mr.","Ms.","Dr."])

# for age
age=tkinter.Label(user_info,text="Age")
age_entry=tkinter.Spinbox(user_info,from_=0,to=110)

# continent
continent=tkinter.Label(user_info,text="continent")
cont_entry=ttk.Combobox(user_info,values=["Asia", "Africa","Europe", "Australia", "North America","South America","Antarctica"])

# second frame
sec_frame=tkinter.LabelFrame(frame)
   
 #######  # checkbutton  
#  StringVar is a built-in Tkinter variable class. It allows you to store and retrieve string values in your Tkinter applications.
#  You can use StringVar to create labels, entries, and other widgets that display or accept text input


res_lable=tkinter.Label(sec_frame,text="resgitration status")
reg_status_var=tkinter.StringVar(value="not regsitered")
res_checkbox=tkinter.Checkbutton(sec_frame,text="currently registered",variable=reg_status_var, onvalue="regsitered",offvalue="not regsiter")

# no. of course
num_course=tkinter.Label(sec_frame,text="completed courses")
num_course_entry=tkinter.Spinbox(sec_frame,from_=0,to=110)

# sem 
sem_label=tkinter.Label(sec_frame,text="semsters")
sem_enter=tkinter.Spinbox(sec_frame,from_=0,to=110)

# frame 3
third_frame=tkinter.LabelFrame(frame,text="Terms and conditions")

# for terms and condition check box
term_var=tkinter.StringVar(value="not Accepted")


term_check=tkinter.Checkbutton(third_frame,text="I accept the terms and conditions.",variable=term_var,onvalue="Accepted",offvalue="not Accepted")

# button 

button=tkinter.Button(frame,text="Data entry",command=data_entry)
frame.pack()
user_info.grid(row=0,column=0)

# for name
firstname.grid(row=0,column=0)
first_entry.grid(row=1,column=0)

lastname.grid(row=0,column=1)
last_entry.grid(row=1,column=1)

# for title 
title.grid(row=0,column=2)
title_entry.grid(row=1,column=2)

# for age
age.grid(row=2,column=0)
age_entry.grid(row=3,column=0)
# continent
continent.grid(row=2,column=1)
cont_entry.grid(row=3,column=1)

# for second frame 
sec_frame.grid(row=1,column=0,stick="news",pady=20,padx=20)

# for checkbutton
res_lable.grid(row=0,column=0)
res_checkbox.grid(row=1,column=0)

# no. of course
num_course.grid(row=0,column=1)
num_course_entry.grid(row=1,column=1)

# sem 
sem_label.grid(row=0,column=2)
sem_enter.grid(row=1,column=2)

# third frame
third_frame.grid(row=2,column=0,sticky="news",padx=20,pady=10)
 
# for terms and condition
term_check.grid(row=0,column=0)

# button
button.grid(row=3,column=0,sticky="news",padx=20,pady=10)



for widget in user_info.winfo_children():
    widget.grid_configure(padx=10,pady=5)
# for frame 2
for widget in sec_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)



root.mainloop()