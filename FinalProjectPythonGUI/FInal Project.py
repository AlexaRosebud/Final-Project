from tkinter import*
from tkinter import messagebox
from openpyxl import workbook
import xlrd
import pandas as pd
from tkinter import *
from tkcalendar import *



root=Tk()
f=Frame(root)
frame1=Frame(root)
frame2=Frame(root)
frame3=Frame(root)
root.title("Employee Time Clock")
root.geometry("830x395")
root.configure(background="White")

scrollbar=Scrollbar(root)
scrollbar.pack(side=RIGHT, fill=Y)

firstname=StringVar()                    
lastname=StringVar()
id=StringVar()
designation=StringVar()
remove_firstname=StringVar()
remove_lastname=StringVar()
searchfirstname=StringVar()
searchlastname=StringVar()
sheet_data=[]
row_data=[]


def emp_dict(*args):
    workbook_name="finalProject.xlsx"
    workbook=xlrd.open_workbook(workbook_name)
    worksheet=workbook.sheet_by_index(0)
    
    wb=load_workbook(workbook_name)
    page=wb.active
    
    p=0
    for i in range(worksheet.nrows):
        for j in range(worksheet.ncols):
            cellvalue=worksheet.cell_value(i,j)
            print(cellvalue)   
            sheet_data.append([])
            sheet_data[p]=cellvalue
            p+=1
    print(sheet_data)
    fl=firstname.get()
    fsl=fl.lower()
    ll=lastname.get()
    lsl=ll.lower()
    if (fsl and lsl) in sheet_data:
        print("found")
        messagebox.showerror("Error","This employee already exist")
    else:
        print("not found")
        for info in args:
            page.append(info)
        messagebox.showinfo("Done","Successfully added the employee record")

    wb.save(filename=workbook_name)
    
def add_entries():                      
    a=" "
    f=firstname.get()
    f1=f.lower()
    l=lastname.get()
    l1=l.lower()
    de=designation.get()
    de1=de.lower()
    list1=list(a)
    list1.append(f1)
    list1.append(l1)
    list1.append(de1)
    emp_dict(list1)


def add_info():
    frame2.pack_forget()
    frame3.pack_forget()
    emp_first_name=Label(frame1,text="Enter first name of the employee: ",bg="purple",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e1=Entry(frame1,textvariable=firstname)
    e1.grid(row=1,column=2,padx=10)
    e1.focus()
    emp_last_name=Label(frame1,text="Enter last name of the employee: ",bg="purple",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e2=Entry(frame1,textvariable=lastname)
    e2.grid(row=2,column=2,padx=10)
    emp_desig=Label(frame1,text="Select Employee's Position: ",bg="purple",fg="white")
    emp_desig.grid(row=4,column=1,padx=10)
    designation.set("Select Option")
    e5=OptionMenu(frame1,designation,"Select Option","Supervisor","Foremen","Apprentice")
    e5.grid(row=4,column=2,padx=10)
    button4=Button(frame1,text="Add entries",command=add_entries)
    button4.grid(row=5,column=2,pady=10)
    
    frame1.configure(background="Purple")
    frame1.pack(pady=10)
    
def clear_all():
    frame1.pack_forget()
    frame2.pack_forget()
    frame3.pack_forget()

    
def remove_emp():
    clear_all()
    emp_first_name=Label(frame2,text="Enter first name of the employee",bg="purple",fg="white")
    emp_first_name.grid(row=1,column=1,padx=10)
    e6=Entry(frame2,textvariable=remove_firstname)
    e6.grid(row=1,column=2,padx=10)
    e6.focus()
    emp_last_name=Label(frame2,text="Enter last name of the employee",bg="purple",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e7=Entry(frame2,textvariable=remove_lastname)
    e7.grid(row=2,column=2,padx=10)
    remove_button=Button(frame2,text="Click to remove",command=remove_entry)
    remove_button.grid(row=3,column=2,pady=10)
    frame2.configure(background="Purple")
    frame2.pack(pady=10)

def remove_entry():
    rsf=remove_firstname.get()
    rsf1=rsf.lower()
    print(rsf1)
    rsl=remove_lastname.get()
    rsl1=rsl.lower()
    print(rsl1)
    workbook_name="finalProject.xlsx"
    path="finalProject.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==rsf1 and row_value[2]==rsl1):
            print(row_value)
            print("found")
            file="finalProject.xlsx"
            x=pd.ExcelFile(file)
            dfs=x.parse(x.sheet_names[0])
            dfs=dfs[dfs['First Name']!=rsf]
            dfs.to_excel("finalProject.xlsx",sheet_name='Employee',index=False)
            messagebox.showinfo("Done","Successfully removed the Employee record")
    clear_all()

def search_emp(): 
    clear_all()
    emp_first_name=Label(frame3,text="Enter first name of the employee",bg="Purple",fg="white") 
    emp_first_name.grid(row=1,column=1,padx=10)
    e8=Entry(frame3,textvariable=searchfirstname)
    e8.grid(row=1,column=2,padx=10)
    e8.focus()
    emp_last_name=Label(frame3,text="Enter last name of the employee",bg="Purple",fg="white")
    emp_last_name.grid(row=2,column=1,padx=10)
    e9=Entry(frame3,textvariable=searchlastname)
    e9.grid(row=2,column=2,padx=10)
    search_button=Button(frame3,text="Click to search",command=search_entry)
    search_button.grid(row=3,column=2,pady=10)
    
    frame3.configure(background="Purple")
    frame3.pack(pady=10)

    
def search_entry():
    sf=searchfirstname.get()
    ssf1=sf.lower()
    print(ssf1)
    sl=searchlastname.get()
    ssl1=sl.lower()
    print(ssl1)
    path="finalProject.xlsx"
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)

    for row_num in range(sheet.nrows):
        row_value = sheet.row_values(row_num)
        if (row_value[1]==ssf1 and row_value[2]==ssl1):
            print(row_value)
            print("found")
            messagebox.showinfo("Searched Employee Exist")
            clear_all()
            
    if(row_value[1]!=ssf1 and row_value[2]!=ssl1):
        print("Not found")
        messagebox.showerror("Employee Record does not Exist")
        clear_all()

        

        
label1=Label(root,text="Employee Time Clock")
label1.config(font=('Italic',16,'bold'), justify=CENTER, background="Black",fg="White", anchor="center")
label1.pack(fill=X)

label2=Label(f,text="Select an action: ",font=('bold',12), background="Black", fg="White")
label2.pack(side=LEFT,pady=10)
button1=Button(f,text="Add", background="Purple", fg="White", command=add_info, width=8)
button1.pack(side=LEFT,ipadx=20,pady=10)
button2=Button(f,text="Remove", background="Purple", fg="white", command=remove_emp, width=8)
button2.pack(side=LEFT,ipadx=20,pady=10)
button3=Button(f,text="Search", background="Purple", fg="White", command=search_emp, width=8)
button3.pack(side=LEFT,ipadx=20,pady=10)
button6=Button(f,text="Close", background="Purple", fg="White", width=8, command=root.destroy)
button6.pack(side=LEFT,ipadx=20,pady=10)
f.configure(background="Black")
f.pack()


import datetime
import time
import fileinput
from operator import index

#Add time in time out
def Add():
    Time_emp = input("Enter 'IN' For Clockin or Enter 'Out' for Clockout: ")
    Time = Time_emp.upper()
    if Time == "OUT":
        Clock_out()
    elif Time == "IN":

        #open txt file
        TimeCard_file = open('TimeCard.txt', 'a')

        # Add clock in time in file with current time.
        EmpName = input('Enter your First and Last Name: ')
        Emp_Name = EmpName.upper()
        now = datetime.datetime.now()
        print(EmpName)
        Current  = now.strftime("%Y-%m-%d")
        Clock_in = now.strftime("%H:%M")

        #Append data to the file
        TimeCard_file.write(Emp_Name + " ")
        TimeCard_file.write(Current + " ")
        TimeCard_file.write(Clock_in)
        print("Current Date: ",Current)
        print("Clock in Time: ", Clock_in)
        TimeCard_file.write('\n')

        #Close the file
        TimeCard_file.close()
    
    #Check if user wants to add another record to the file
    func = input("Enter C to Close or Press M for main menu:")
    if func == "C" or func =="c":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

#First Ask for employee name so can add clock out time for same person.
def Clock_out():
    found = False
    val = 'x'
    #input the name what you want to search
    employee_name = input("Please Enter your First and Last Name: ")
    employee = employee_name.upper()
    #open the time card file and search name
    TimeCard_file = open("TimeCard.txt", 'r')
    TimeCard = TimeCard_file.readline()

    #read the file if you have entered any name
    while TimeCard != '':
        found = TimeCard.startswith(employee)
        if found:
            val = TimeCard
        TimeCard = TimeCard_file.readline()
    TimeCard_file.close()

    if val != '':
        now = datetime.datetime.now()
        #Add clouckout time in file with current time
        Current  = now.strftime("%Y-%m-%d")
        Clock_out = now.strftime("%H:%M")
        search = val
        print(search.rstrip('\n') + ' ' + Clock_out)
        print("You have been successfully clocked out")

        #open the file
        fn = "TimeCard.txt"
        f = open(fn)
        output = []
        #for loop i you find search record
        for line in f:
            if line.startswith(val):
                output.append(line.replace(line, line.rstrip('\n') + ' ' + Clock_out) + '\n')
            else:
                output.append(line)
        f.close()
        f = open(fn, 'w')
        f.writelines(output)
        f.close()

    #open the time card file and search name
    TimeCard_file = open('TimeCard.txt', 'r')

    TimeCard = TimeCard_file.readline()
    sum_h = 0
    sum_m = 0

    # read the file if you have entered any name
    while TimeCard != '':
        found =  TimeCard.startswith(employee)

        if found:
            start = TimeCard[-12:-7]
            end = TimeCard[-6:-1]
            #split the date
            clock_in = start.split(':')
            clock_out = end.split(':')

            #convert tiem from string to integer
            clock_out_int_h = int(clock_out[0])
            clock_out_int_m = int (clock_out[1])
            clock_in_int_h = int(clock_in[0])
            clock_in_int_m = int (clock_in[1])

            #calculate Working Hours
            if clock_out_int_m > clock_in_int_m:
                hours = (clock_out_int_h - clock_in_int_h)
                minutes = (clock_out_int_m - clock_in_int_m)
            else:
                hours = (clock_out_int_h - clock_in_int_h) - 1
                minutes = 60 - (clock_in_int_m - clock_out_int_m)
            sum_h += hours
            sum_m += minutes
        TimeCard = TimeCard_file.readline()

    #Collect Total Hours of the Employee
    if sum_m > 60:
        sum = (sum_m / 60)
        split_min = str(sum).split('.')
        int_part = int(split_min[0])
        decimal = int((sum - int_part) * 60)
        Total_hours = sum_h + int_part
        print("Weekly Working Hours is:", Total_hours, "Hours and", decimal, "Minutes", '\n')
    else:
        print("Weekly Working Hours is:", sum_h, "Hours and", sum_m,"Minutes", '\n')

    #added function to make decision if you want to work on this file or exit

    func = input("Enter Q to quit or Press M for main menu:")
    if func == "Q" or func =="q":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

#This function will display all working time for all employee
def TimeReport():
    #open time card
    TimeCard_file = open('TimeCard.txt','r')
    TimeCard = TimeCard_file.readline()
    #read the rest of the file
    while TimeCard != '':
        #Display the record
        print('Employee Hours:', TimeCard)

        #Read the next Description
        TimeCard = TimeCard_file.readline()
    #close the file
    TimeCard_file.close()

    #Added function to make decision if want to work or exit
    func = input("Enter Q to quit or Press M for main menu:")
    if func == "Q" or func =="q":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

def Edit():
    found =  False

    #input the name you want to search
    EmpName = input("Please Enter Employee Name:")
    employee = EmpName.upper()

    #open the time card file and search name
    TimeCard_file = open('TimeCard.txt','r')
    TimeCard = TimeCard_file.readline()


    #read the file if you have entered any name
    while TimeCard != '':
        found = TimeCard.startswith(employee)
        if found:
            print(TimeCard)

        TimeCard = TimeCard_file.readline()
    TimeCard_file.close()
    
    Date = input("Enter the Date you want to edit:")
    search = (employee + " " + Date)

    #Open the file
    fn = "TimeCard.txt"
    f = open(fn)
    output = []
    #for loop i you find search record
    for line in f:
        if not line.startswith(search):
            output.append(line)
    f.close()
    f = open(fn, 'w')
    f.writelines(output)
    f.close()

    #open the TimeCard.txt file in append mode
    TimeCard_file = open('TimeCard.txt', 'a')
    print("enter Clocl in and Clock out for",employee)

    Clock_in = input("Enter Clock In time:")
    Clock_out = input("Enter Clock Out time:")

    TimeCard_file.write(employee + " ")
    TimeCard_file.write(Date + " ")
    TimeCard_file.write(str(Clock_in) + " ")
    TimeCard_file.write(str(Clock_out) + " ")
    TimeCard_file.write('\n')

    #Close the file
    TimeCard_file.close()
    
    #Check if user wants to add another record to the file
    func = input("Enter Q to quit or Press M for main menu:")
    if func == "Q" or func =="q":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

#This function will search emplyee and his working hours
def Search():
    found = False

    EmpName  =  input("Please enter Employee Name:")
    search = EmpName.upper()

    #open the time card file and search name
    TimeCard_file = open('TimeCard.txt', 'r')

    TimeCard = TimeCard_file.readline()
    sum_h = 0
    sum_m = 0

    #read the file if you have entered any name
    while TimeCard != '':
        found = TimeCard.startswith(search)
        if found:
            #print(TimeCard)
            start = TimeCard[-12:-7]
            #print("start",start)
            end = TimeCard[-6:-1]
            print("end",end)
            #split the date
            clock_in = start.split(':')
            #print(clock_in)
            clock_out = end.split(':')
            #print(clock_out)

            #convert tiem from string to integer
            clock_out_int_h = int(clock_out[0])
            clock_out_int_m = int (clock_out[1])
            clock_in_int_h = int(clock_in[0])
            clock_in_int_m = int (clock_in[1])

            #calculate Working Hours
            if clock_out_int_m > clock_in_int_m:
                hours = (clock_out_int_h - clock_in_int_h)
                minutes = (clock_out_int_m - clock_in_int_m)
                #print("Total Hours Worked:", hours, 'Hours and', minutes,"Minutes", '\n' )
            else:
                hours = (clock_out_int_h - clock_in_int_h) - 1
                minutes = 60 - (clock_in_int_m - clock_out_int_m)
            sum_h += hours
            sum_m += minutes
            print("Total Hours Worked:", hours, 'Hours and', minutes,"Minutes", '\n' )
        TimeCard = TimeCard_file.readline()

    #Collect Total Hours of the Employee
    if sum_m > 60:
        sum = (sum_m / 60)
        split_min = str(sum).split('.')
        int_part = int(split_min[0])
        decimal = int((sum - int_part) * 60)
        Total_hours = sum_h + int_part
        print("Total Working Hours is:", Total_hours, "Hours and", decimal, "Minutes", '\n')
    else:
        print("Total Working Hours is:", sum_h, "Hours and", sum_m,"Minutes", '\n')

    #close the file
    TimeCard_file.close()
    #added function to make decision if you want to work on this file or exit

    func = input("Enter Q to quit or Press M for main menu:")
    if func == "Q" or func =="q":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

#this fucntion will remove unwanted data from report
def Delete():
    found = False


    #input the name you want to search
    EmpName = input("Please Enter Employee Name:")
    employee = EmpName.upper()

    #open the time card file and search name
    TimeCard_file = open('TimeCard.txt','r')
    TimeCard = TimeCard_file.readline()


    #read the file if you have entered any name
    while TimeCard != '':
        found = TimeCard.startswith(employee)
        if found:
            print(TimeCard)

        TimeCard = TimeCard_file.readline()
    TimeCard_file.close()
    TimeCard = filter(lambda x: not x.isspace(), TimeCard)
    #Find the blank space and delete it
    Date = input("Enter the date you want to delete:")
    search = (employee + '' + Date)

    #open the file
    fn = 'TimeCard.txt'
    f = open(fn)
    output = []
    #for loop i you find search record
    for line in f:
        if not line.startswith(search):
            output.append(line)
    f.close()
    f = open(fn, 'w')
    f.writelines(output)
    f.write("".join(TimeCard))
    f.close()

    print("You successfully deleted " + search + "'s Record")
    func = input("Enter Q to quit or Press M for main menu:")
    if func == "Q" or func =="q":
        quit()
    elif func == "M" or func =="m":
        return main()
    else:
        print("Incorrect input, Please Try Again \n")

#Add date in time Card
def main():
    #Select the option what you like to do in you time card
    print("What would you like to do in Time Card System?")
    print("A = Add :: S = Search :: D = Delete :: E = Edit :: R = Report :: Q = Quit")
    func = input("Please select a function from the list above:")

    #Use while loop to enter in time card as selected option
    while func == '' or func != 'A'  or func != 'a' or func != 'S' or func != 's' or func != 'D' or func != 'd' or func != 'E' or func != 'e' or func != 'R' or func != 'r':
        if func == 'A' or func == 'a':
            Add()
        elif func == 'S' or func == 's':
            Search()
        
        # only manager can use Report, Delte and Edit Optipns
        elif func == 'D' or func == 'd':
            print('Are you Supervisor?:')
            Mrg_emp = input("Y=Yes, Anything Else = No:")
            Mgr= Mrg_emp.upper()
            if (Mgr == 'YES' or Mgr == 'Y'):
                Delete()
            else:
                print('\n')
                main()
        elif func == 'E' or func == 'e':
            print('Are you Supervisor?:')
            Mrg_emp = input("Y=Yes, Anything Else = No:")
            Mgr= Mrg_emp.upper()
            if (Mgr == 'YES' or Mgr == 'Y'):
                Edit()
            else:
                print('\n')
                main()
        elif func == 'R' or func == 'r':
            print('Are you Supervisor?:')
            Mrg_emp = input("Y=Yes, Anything Else = No:")
            Mgr= Mrg_emp.upper()
            if (Mgr == 'YES' or Mgr == 'Y'):
                TimeReport()
            else:
                print('\n')
                main()
        elif func == "Q" or func == "q":
            quit()
        #if you press any random key it will bring you here
        else:
            print("Incorrect input please try again \n")
            func = input("Enter A for Add, S for Search, D for Delete, E for Edit, R for Report:")


root.mainloop()









