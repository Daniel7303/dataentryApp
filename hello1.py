import tkinter

from tkinter import messagebox, ttk

import openpyxl
from openpyxl import load_workbook
root = tkinter.Tk()

root.geometry("300x400")
root.configure(bg="")

file_path = r"C:\Users\LENOVO\OneDrive\Documents\dataentrywithPY.xlsx"
A = openpyxl.load_workbook(file_path)
B = A["Registration"]
def onClick_submit():
    name = name_textbox.get()
    email = email_textbox.get()
    phone = phone_textbox.get()
    branch = branch_dropdown.get()

    if name and email and phone and branch:
        if not phone.isdigit():
            messagebox.showwarning("Warning", "Tel details must be a number")
        else:
            B.append([name, email, phone, branch])
            A.save(file_path)
            messagebox.showinfo("Status","Captured Data")
    else:
        messagebox.showwarning("Warning", "fill all the fields")


root.title("student Registration form")
text_label = tkinter.Label(root, text="Enter Name")
text_label.pack(anchor=tkinter.W,padx=10)
name_textbox = tkinter.Entry(root)
name_textbox.pack(anchor=tkinter.W,padx=10)

text_label = tkinter.Label(root, text="Enter Email")
text_label.pack(anchor=tkinter.W,padx=10)
email_textbox = tkinter.Entry(root)
email_textbox.pack(anchor=tkinter.W,padx=10)

text_label = tkinter.Label(root, text="Phone number")
text_label.pack(anchor=tkinter.W,padx=10)
phone_textbox = tkinter.Entry(root)
phone_textbox.pack(anchor=tkinter.W,padx=10)

# a dropdown
choices = ['CS', 'EC', 'Mechanical']
branch_label = tkinter.Label(root, text="Select Department")
branch_label.pack(anchor=tkinter.W,padx=10)
branch_dropdown = ttk.Combobox(root, values=choices)
branch_dropdown.pack(anchor=tkinter.W,padx=10)

# submit button
submit_button = tkinter.Button(root, text="Submit", command=onClick_submit)
submit_button.pack(anchor=tkinter.W,padx=10)



root.mainloop()





