import tkinter as tk
from tkinter import filedialog
from tkinter import ttk

test = "best"
files = "best"
testfile = ""
class GUIshell():
    # declaring string variable
    # for storing name and password
    def __init__():
        window = tk.Tk()
        print("GUIshell")
        window.title("Python Script Generator")
        global test
        global files
        global testfile
        testName=tk.StringVar()
        fileName=tk.StringVar()

        label1 = tk.Label(window, text="Insert name of test: ")
        label2 = tk.Label(window, text="Insert File name: ")
        test_entry = tk.Entry(window,textvariable = testName, font=('calibre',10,'normal'))
        file_entry = tk.Entry(window,textvariable = fileName, font=('calibre',10,'normal'))
        test = test_entry.get()
        files = file_entry.get()
        label3 = tk.Label(window, text="Single Insert?")


        n = tk.StringVar()
        tempSelection = ttk.Combobox(window, width = 27, textvariable = n)

        # Adding combobox drop down list
        tempSelection['values'] = ('True', 
                                'False')
        tempSelection.current(0)


        def close_window():
            window.quit()
            window.destroy()
        label4 = tk.Label(window, text="Want to quit?")
        quitbutton = tk.Button(window, text="Quit?", command=close_window)

        #label1.grid(row=0, column=0)
        #test_entry.grid(row=0,column=1)
        label2.grid(row=0,column=0)
        file_entry.grid(row=0,column=1)
        label3.grid(row=1,column=0)
        tempSelection.grid(row = 1,column= 1)
        label4.grid(row=2,column=0)
        quitbutton.grid(row=2,column=1)

        def testfileName(testfileName):
            global testfile
            testfile = testfileName
        def Submit():
            test = test_entry.get()
            files = file_entry.get()
            temps = tempSelection.get()
            print("Test Name: " + test)
            print("File name: " + files)
            close_window()
            testfileNames = [test, files, temps]
            testfileName(testfileNames)

        def GUIshellSubmit(event):
            testfileNames = Submit()

        file_entry.bind('<Return>', GUIshellSubmit)

        window.mainloop()
        return testfile