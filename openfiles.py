import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import csv
import openpyxl 
import os


root=tk.Tk()
def open_file_dialog():
    root.withdraw()
    root.lift()
    root.focus_set()
    file_path = filedialog.askopenfilenames()
    #print("file_path:")
    #print(file_path)
    return file_path
def open_directory_dialog(FileName):
    root.withdraw()
    root.lift()
    root.focus_set()
    file_path = filedialog.askdirectory() + FileName + ".csv"
    return file_path
def askdirectory():
    root.withdraw()
    root.lift()
    root.focus_set()
    filedirectory = filedialog.askdirectory()
    return filedirectory 
def makeDirectory(FileName):

    rootPath = askdirectory()
    Directory = rootPath + "/" + FileName
    Path = Directory + "/" + FileName
    #print(f"makeDirectory FileName  is: {FileName}")
    #print(f"makeDirectory directory is: {Directory}")
    #print(f"makeDirectory root path is: {rootPath}")
    #print(f"makeDirectory Path      is: {Path}")
    try:
        os.mkdir(Directory)
        print(f"Directory '{Directory}' created successfully")
    except FileExistsError:
        print(f"Directory '{Directory}' already exists")
    except OSError as e:
        print(f"Error creating directory: {e}")
    return Path

def csv_to_xlsx_openpyxl(csv_file_path, currfile):
    xlsx_file_path = csv_file_path[currfile].split('.')
    xlsx_file_path[1] = '.xlsx'
    file_path = "".join(xlsx_file_path)
    #file_path = str(xlsx_file_path) + 'xlsx'
    print(file_path)
    """Converts a CSV file to an XLSX file using csv and openpyxl."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    with open(csv_file_path[currfile], 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            sheet.append(row)

    workbook.save(file_path)
    return file_path
def deletefile(file_path):

    try:
        os.remove(file_path)
        print(f"File '{file_path}' deleted successfully.")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
    except PermissionError:
        print(f"Error: Permission denied to delete '{file_path}'.")
    except Exception as e:
        print(f"An error occurred: {e}")
def button_clicked():
    print("button Clicked!")
def ButtonShell():
    # Creating a button with specified options
    button = tk.Button(root, 
                    text="Click Me", 
                    command=button_clicked,
                    activebackground="blue", 
                    activeforeground="white",
                    anchor="center",
                    bd=3,
                    bg="lightgray",
                    cursor="hand2",
                    disabledforeground="gray",
                    fg="black",
                    font=("Arial", 12),
                    height=2,
                    highlightbackground="black",
                    highlightcolor="green",
                    highlightthickness=2,
                    justify="center",
                    overrelief="raised",
                    padx=10,
                    pady=5,
                    width=15,
                    wraplength=100)
    button.pack(padx=20, pady=20)

    root.mainloop()
# declaring string variable
# for storing name and password
name_var=tk.StringVar()
passw_var=tk.StringVar()
def loginShell():


    root = tk.Tk()
    # setting the windows size
    root.geometry("600x400")
    

    
    # defining a function that will
    # get the name and password and 
    # print them on the screen
    # creating a label for 
    # name using widget Label
    name_label = tk.Label(root, text = 'Username', font=('calibre',10, 'bold'))
    
    # creating a entry for input
    # name using widget Entry
    name_entry = tk.Entry(root,textvariable = name_var, font=('calibre',10,'normal'))
    
    # creating a label for password
    passw_label = tk.Label(root, text = 'Password', font = ('calibre',10,'bold'))
    
    # creating a entry for password
    passw_entry=tk.Entry(root, textvariable = passw_var, font = ('calibre',10,'normal'), show = '*')
    
    # creating a button using the widget 
    # Button that will call the submit function 
    sub_btn=tk.Button(root,text = 'Submit', command = submit)
    
    # placing the label and entry in
    # the required position using grid
    # method
    name_label.grid(row=0,column=0)
    name_entry.grid(row=0,column=1)
    passw_label.grid(row=1,column=0)
    passw_entry.grid(row=1,column=1)
    sub_btn.grid(row=2,column=1)
    root.mainloop()
def submit():

    name=name_var.get()
    password=passw_var.get()
    
    print("The name is : " + name)
    print("The password is : " + password)
    
    name_var.set("")
    passw_var.set("")
def dropdownShell():
    # Creating tkinter window
    window = tk.Tk()
    window.title('Combobox')
    window.geometry('500x250')

    # label text for title
    ttk.Label(window, text = "GFG Combobox Widget", 
            background = 'green', foreground ="white", 
            font = ("Times New Roman", 15)).grid(row = 0, column = 1)

    # label
    ttk.Label(window, text = "Select the Month :",
            font = ("Times New Roman", 10)).grid(column = 0,
            row = 5, padx = 10, pady = 25)

    # Combobox creation
    n = tk.StringVar()
    monthchoosen = ttk.Combobox(window, width = 27, textvariable = n)

    # Adding combobox drop down list
    monthchoosen['values'] = (' January', 
                            ' February',
                            ' March',
                            ' April',
                            ' May',
                            ' June',
                            ' July',
                            ' August',
                            ' September',
                            ' October',
                            ' November',
                            ' December')

    monthchoosen.grid(column = 1, row = 5)
    monthchoosen.current(0)
    tk.Button(window, text="Quit?", command=root.destroy)
    window.mainloop()
