import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import DataConversion
import Plot
import openfiles

#grab worksheet
DataSheet = openpyxl.load_workbook('pydata.xlsx', data_only=True)

#activate the sheet
sheet1 = DataSheet.active


#user input for number of tests and temps
#NoTests = input("Please enter number of tests: ")
#DataLength = input("Please enter data length: ")
#InsertType = input("SingleInsert y/n: ")
#   if(InsertType = "y" || InsertType = "Y"):
#       SingleInsert = True
#   else:
#       SingleInsert = False
#user input automated for testing
#NoTests = "4"
#DataLength = "25"
SingleInsert = True

#working on converting to num based  call to sheet for ease of working
#lolipop = sheet1.cell(row=9,column=3)

#call data Conversion
#exampleData = DataConversion.Conversion(int(NoTests), SingleInsert, int(DataLength))
print("openfiles.open_file_dialog")
pydata = openfiles.open_file_dialog()
filenumber = 0
print("dataconversion.getnotests")
NoTests = DataConversion.GetNoTests(pydata)
print(NoTests)
for thisfile in pydata:
    print("dataconversion.conversion")
    data = DataConversion.Conversion(int(NoTests[filenumber]), SingleInsert, pydata[filenumber])
    print("pydata[filenumber]")
    print(pydata[filenumber])
    print("data")
    print(data)
    print('hasTemp')
    print(DataConversion.GethasTemp(pydata[filenumber]))
    #`RunName = input("Please Enter the name of the run: \n")
    #creates output file so you don't clog the dir
    #Path = "./" + input("inputFileName: \n")
    import GUIshell
    testname = GUIshell.GUIshell.__init__()
    RunName = testname[0]
    Path = "./" + testname[1]
    if(testname[2] == "True"):
        SingleInsert = True
    else:
        SingleInsert = False
    if(DataConversion.GethasTemp(pydata[filenumber]) == True):
        #example of a ploting function
        Plot.PlotTempVsTest(int(NoTests[filenumber]), SingleInsert, data, pydata[filenumber], RunName, Path)
    else:
        Plot.PlotTempVsTest(int(NoTests[filenumber]), SingleInsert, data, pydata[filenumber], RunName, Path)
    #elif(DataConversion.GethasTemp(pydata[filenumber]) == False):
    #    SingleInsert = False
    #    Plot.PlotSnVsTest(int(NoTests[filenumber]), SingleInsert, data, pydata[filenumber], RunName)
   #elif(input("repeatablity") == "True"):
   #    SingleInsert = True
   #    i230765Repeatablity.Repeatablity(int(NoTests[filenumber]), SingleInsert, data, pydata[filenumber], RunName)
    #SingleInsert = True
    #Plot.PlotSnVsTest(int(NoTests[filenumber]), SingleInsert, data, pydata[filenumber], RunName)
    filenumber += 1


print("ran the script \n")
plt.show()
#openfiles.loginShell()



