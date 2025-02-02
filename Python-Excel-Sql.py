import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl
import sqlite3

def enter_data():
    accepted = accept_var.get()
    
    if accepted=="Accepted":
        # Getting user info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        
        if firstname and lastname:
            gender = gender_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            products = product_combobox.get()
            produuct_num = product_num_spinbox.get()
            
            
            print("First name: ", firstname, "Last name: ", lastname,)
            print("Gender: ", gender, "Age: ", age, "Nationality: ", nationality)
            print("Product Name: ", products, "No. of Products: ", produuct_num)

            print("------------------------------------------")
            
            filepath = "path-diretion.xlsx" # Give the file path, where you want to save the file
            
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading1 = ["Product Sale"]
                heading2 = ["First Name", "Last Name", "Gender", "Age", "Nationality",
                           "Product Name", "Number of Products"]
                sheet.append(heading1)
                sheet.append(heading2)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, gender, age, nationality, products, produuct_num])
            workbook.save(filepath)

            # Create Database (DB) Table
            conn = sqlite3.connect('path-direction.db') # connecting to sqlite3 and give your filepath
            table_create_query = '''CREATE TABLE IF NOT EXISTS Product_Sale
                    (Firstname TEXT, Lastname TEXT, Gender TEXT, Age INT, Nationality TEXT, Products TEXT, Produuct_num INT)
            '''
            conn.execute(table_create_query)

            # Insert data
            data_insert_query = '''INSERT INTO Product_Sale (firstname, lastname, gender, age, nationality, products, produuct_num)
                    VALUES (?,?,?,?,?,?,?)
            '''
            data_insert_tuple = (firstname, lastname, gender, age, nationality, products, produuct_num)
            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()


            conn.close()
                
        else: #showing messsage box if required input is missing
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title= "Attention", message="You have not accepted the terms & conditions")

window = tkinter.Tk()
window.title("Company Name")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame =tkinter.LabelFrame(frame, text="Product Sale")
user_info_frame.grid(row= 0, column=0, padx=20, pady=10)


first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

gender_label = tkinter.Label(user_info_frame, text="Gender")
gender_combobox = ttk.Combobox(user_info_frame, values=["", "Male", "Female"])
gender_label.grid(row=2, column=0)
gender_combobox.grid(row=3, column=0)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=10, to=110)
age_label.grid(row=2, column=1)
age_spinbox.grid(row=3, column=1)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["Africa", "Asia", "Europe", "North America", "Oceania", "South America"])
nationality_label.grid(row=2, column=2)
nationality_combobox.grid(row=3, column=2)

product_label = tkinter.Label(user_info_frame, text= "Products")
product_combobox = ttk.Combobox(user_info_frame, values=["Cleanroom garments", "Cleanroom undergarments", "Disposable garments & PPE", "Cleanroom gloves", "Shoes & socks", "Cleanroom wipes", "Disinfectants", "Tacky mats", "Cleaning equipment", "Disposal systems and accessories", "Dispenser systems", "Furniture", "Cleanroom paper & accessories", "Tapes & Labels", "Specific products", "Glossary"])
product_label.grid(row=4, column=0)
product_combobox.grid(row=5, column=0)

product_num_label= tkinter.Label(user_info_frame, text="Number of products")
product_num_spinbox = tkinter.Spinbox(user_info_frame, from_=0, to=100)
product_num_label.grid(row=4, column=1)
product_num_spinbox.grid(row=5, column=1)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)


# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text= "I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command= enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)
 
window.mainloop()