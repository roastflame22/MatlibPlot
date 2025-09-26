import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import DataConversion
import Plot
import openfiles
from   pypetb import RnR
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
#set seaborn style to improve the figure sight
sns.set()


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
#SingleInsert = True
gageRnR = False

#working on converting to num based  call to sheet for ease of working
#lolipop = sheet1.cell(row=9,column=3)

#call data Conversion
#exampleData = DataConversion.Conversion(int(NoTests), SingleInsert, int(DataLength))
#print("openfiles.open_file_dialog")
pydata = openfiles.open_file_dialog()
filenumber = 0
#print("dataconversion.getnotests")
NoTests = DataConversion.GetNoTests(pydata)
#print(NoTests)
gageRnR_Data = [[[""] * len(pydata)] * int(NoTests[0]*3)] * 9
tester = [""] * len(pydata)
for thisfile in pydata:
    SingleInsert = DataConversion.GetSingleInsert(pydata, filenumber)
    #print("single insert: ")
    #print(SingleInsert)
    #print("dataconversion.conversion")
    #print("pydata[filenumber]")
    #print(pydata[filenumber])
    #print("NoTests[filenumber]")
    #print(NoTests[filenumber])
    data = DataConversion.Conversion(int(NoTests[filenumber]), SingleInsert, pydata[filenumber])
    #print("pydata[filenumber]")
    #print(pydata[filenumber])
    #print("data")
    #print(data)
    #print('hasTemp')
    #print(DataConversion.GethasTemp(pydata[filenumber]))
    #`RunName = input("Please Enter the name of the run: \n")
    #creates output file so you don't clog the dir
    #Path = "./" + input("inputFileName: \n")
    import GUIshell
    testname = GUIshell.GUIshell.__init__()
    RunName = testname[1]
    
    #Path = openfiles.makeDirectory(testname[1])
    Path = openfiles.askdirectory() + "/" + testname[1]
    if(gageRnR):
        print("gageRnR_Data")
        print(gageRnR_Data)
        gageRnR_Data[filenumber] = data
        tester[filenumber] = DataConversion.GetTester(pydata[filenumber])
        print("gageRnR_Data")
        print(gageRnR_Data)
   #if(testname[2] == "True"):
   #    SingleInsert = True
   #else:
   #    SingleInsert = False
    if(SingleInsert):
        print("starting plot function")
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
    filenumber = filenumber + 1

if(gageRnR):
    df = pd.read_csv(openfiles.open_file_dialog()[0], sep=';')
    print(df.info())
    #Build up the model
    dict_key={'1':'Technician','2':'Parts','3':'Valor'}
    RnRModel=RnR.RnRNumeric(
        mydf_Raw=df,
        mydict_key=dict_key
        )
    #Solve it
    RnRModel.RnRSolve()
    #Check the calculation
    print(RnRModel.getLog())
    df_Result=RnRModel.RnR_varTable()
    #Checking var. table
    print(df_Result)
    #accesing one individual value
    print('\nRnR RESULT:\n-------------------')

    dbl_RnR=df_Result['% Contribution'].loc['Total Gage R&R']
    print(f"Total Gage R&R: {dbl_RnR:.3f}")
    if dbl_RnR<1:
        print('<1% --> Acceptable measurement system')
    elif dbl_RnR>=1 and dbl_RnR<=9:
        print(
            '1-9%--> It may be acceptable depending on application and cost'
            )
    else:
        print(
            '>9% --> Unacceptable measurement system, it must be improved'
            )
    print("gage RnR")
    gage = 0
    for n in tester:
        paths = DataConversion.GageRnR(tester, gageRnR_Data, pydata)
        print("Paths")
        print(paths)
        gage += 1
print("ran the script \n")
plt.show()
#openfiles.loginShell()



