from tkinter import *
from PIL import ImageTk,Image
import time
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from datetime import datetime
#from wordcloud import WordCloud, STOPWORDS, Image

def generatefilters():
    print("ExcelFilePath: " +str(FilePath.get()))
    str_ExcelFilepath = str(FilePath.get())
    data = pd.read_excel(str_ExcelFilepath)
    Li_ColumnNames = data.columns
    for i in Li_ColumnNames:
        print(i)

'''
def buildgraph():
    #Print output values
    print("ExcelFilePath: "+str(FilePath.get()))
    print("ColumnName1: " + str(ColumnName1.get()))
    print(DefaultValueFilter1.get())
    print(DefaultValueBG1.get())
    print("ColumnName2: " + str(ColumnName2.get()))
    print(DefaultValueFilter2.get())
    print(DefaultValueBG2.get())
    #Initialize Variables:
    str_ExcelFilepath = str(FilePath.get())
    str_ColumnName1 = ColumnName1.get()
    str_ColumnName2 = ColumnName2.get()
    str_FilterValue1 = DefaultValueFilter1.get()
    str_FilterValue2 = DefaultValueFilter2.get()
    str_GraphValue1 = DefaultValueBG1.get()
    str_GraphValue2 = DefaultValueBG2.get()
    #Build Graph1:
    # If Column Name is not blank, generate graph:
    if (str_ColumnName1 != "") :
        #Preprocess the data
        data = pd.read_excel(str_ExcelFilepath)
        UniqueValuesCount1 = pd.value_counts(data[str_ColumnName1])
        UniqueValues1 = data[str_ColumnName1].dropna().unique().tolist()
        str_graphname1 = str(datetime.now()).replace(":",".").replace(" ",".").replace("-",".")+"_Graph1.PNG"
        # Generate the garph
        if (str_GraphValue1 == "PieChart"):
            plt.style.use("fivethirtyeight")
            slices = UniqueValuesCount1
            labels = UniqueValues1
            plt.pie(slices, labels=labels, wedgeprops={'edgecolor': 'Black'}, shadow=True, autopct='%1.1f%%')
            plt.title("Graph1 as per " + str_FilterValue1)
            plt.tight_layout()
            plt.savefig(str_graphname1)
            plt.clf()
        elif(str_GraphValue1 == "BarGraph"):
            plt.bar(UniqueValues1, UniqueValuesCount1)
            plt.title("Graph1 as per " + str_FilterValue1)
            plt.xlabel('Values')
            plt.ylabel('Count')
            plt.grid(True)
            plt.rcParams['figure.figsize'] = (20,20)
            plt.savefig(str_graphname1)
            plt.clf()
    if (str_ColumnName2 != "") :
        #Preprocess the data
        data = pd.read_excel(str_ExcelFilepath)
        UniqueValuesCount2 = pd.value_counts(data[str_ColumnName2])
        UniqueValues2 = data[str_ColumnName2].dropna().unique().tolist()
        str_graphname2 = str(datetime.now()).replace(":",".").replace(" ",".").replace("-",".") + "_Graph2.PNG"
        # Generate the garph
        if (str_GraphValue2 == "PieChart"):
            plt.style.use("fivethirtyeight")
            slices = UniqueValuesCount2
            labels = UniqueValues2
            plt.pie(slices, labels=labels, wedgeprops={'edgecolor': 'Black'}, shadow=True, autopct='%1.1f%%')
            plt.title("Graph2 as per " + str_FilterValue2)
            plt.tight_layout()
            plt.savefig(str_graphname2)
            plt.clf()
        elif(str_GraphValue2 == "BarGraph"):
            plt.bar(UniqueValues2, UniqueValuesCount2)
            plt.title("Graph2 as per " + str_FilterValue2)
            plt.xlabel('Values')
            plt.ylabel('Count')
            plt.grid(True)
            plt.rcParams['figure.figsize'] = (20, 20)
            plt.savefig(str_graphname2)
            plt.clf()
        elif(str_GraphValue2 == "WordCloud"):
            pass
'''


def openNewWindow():

    #Get list of all columns in excel
    print("ExcelFilePath: " + str(FilePath.get()))
    str_ExcelFilepath = str(FilePath.get())
    data = pd.read_excel(str_ExcelFilepath)
    Li_ColumnNames = data.columns


    # Building a new window
    newWindow = Toplevel()


    # sets the geometry of toplevel
    newWindow.geometry("1200x700")
    newWindow.maxsize(1200, 700)
    newWindow.minsize(1200, 700)

    newWindow.title("Analytics Dashboard by T-systems")
    newWindow.configure(background="black")

    # Add backgroundImage
    image1 = ImageTk.PhotoImage(Image.open("BgImg.PNG"))
    newWindow_canvas = Canvas(newWindow, width=1080, height=2160)
    newWindow_canvas.pack(fill="both", expand=True)
    newWindow_canvas.create_image(0, 0, image=image1, anchor="nw")

    # build title and project name
    newWindow_canvas.create_text(550, 35, text="Analytics Dashboard", font=("comicssansns", 20, "bold"), fill='White')
    LogoImg = ImageTk.PhotoImage(Image.open("TSysLogo.PNG"))
    newWindow_canvas.create_image(550, 100, image=LogoImg)

    # Add separator line
    newWindow_canvas.create_line(1, 150, 1360, 150, fill="#fb0")

    # build label Column Name
    newWindow_canvas.create_text(85, 170, text="Choose Column Name:", font=("comicssansns", 10, "bold"), fill='White')

    x = 1
    int_Rowcounter = 220
    int_ColumnCounter = 100
    int_Dropdowncounter = 0
    BGOptions1 = ["PieChart", "BarGraph", "WordCloud"]


    for i in Li_ColumnNames:


        varx = IntVar()
        Chk_Button = Checkbutton(newWindow_canvas, text=i, variable=varx, width=20)
        newWindow_canvas.create_window(int_ColumnCounter, int_Rowcounter, window=Chk_Button)
        DefaultValueBG1 = StringVar(newWindow)
        DefaultValueBG1.set(BGOptions1[0])

        BGOption1 = OptionMenu(newWindow, DefaultValueBG1, *BGOptions1)
        BGOption1.config(width=8, font=('Helvetica', 10))
        newWindow_canvas.create_window(int_ColumnCounter+180, int_Rowcounter , window=BGOption1)

        x = x + 1
        int_Rowcounter = int_Rowcounter + 50
        int_Dropdowncounter = int_Dropdowncounter + 50
        if (x==10) :
            int_ColumnCounter = int_ColumnCounter+500
            int_Rowcounter = 220


    newWindow.mainloop()
    app_root.destroy()

#>>>>>>>>>>>>>>>>>>>> Build Main Application <<<<<<<<<<<<<<<<<<<<<<<<<<
app_root = Tk()

#GUI Framework
app_root.geometry("500x400")
app_root.maxsize(500,400)
app_root.minsize(500,400)

app_root.title("Analytics Dashboard by T-systems")
app_root.configure(background="black")

#Add backgroundImage
image1=ImageTk.PhotoImage(Image.open("BgImg.PNG"))
app_canvas = Canvas(app_root,width=1080,height=2160)
app_canvas.pack(fill="both",expand=True)
app_canvas.create_image(0,0,image=image1,anchor="nw")

#build title and project name
app_canvas.create_text(250,35,text="Analytics Dashboard",font=("comicssansns",20,"bold"),fill='White')
LogoImg = ImageTk.PhotoImage(Image.open("TSysLogo.PNG"))
app_canvas.create_image(250, 100, image=LogoImg)


#Add separator line
app_canvas.create_line(1, 150, 1360, 150,fill="#fb0")

#build input filepath input box
app_canvas.create_text(85,170,text="Enter Input File Path:",font=("comicssansns",10,"bold"),fill='White')
FilePath = Entry (app_root,width=50)
app_canvas.create_window(170, 200, window=FilePath)

# Function for opening the file
def file_opener():
    rep = filedialog.askopenfilenames(
        parent=app_root,
        initialdir='/',
        initialfile='tmp',
        filetypes=[("All files", "*")])
    print(rep)


# Browse Button label
Btn_Browse = Button(app_root, text ='Browse', command = lambda:file_opener())
Win_BtnBrowseWindow = app_canvas.create_window(365,200,window=Btn_Browse)

#Add separator line
app_canvas.create_line(1, 230, 1360, 230,fill="#fb0")

#build output folder input box
app_canvas.create_text(95,250,text="Enter Output Folder Path:",font=("comicssansns",10,"bold"),fill='White')
OutputFolderPath = Entry (app_root,width=50)
app_canvas.create_window(170, 280, window=OutputFolderPath)

#Add separator line
app_canvas.create_line(1, 300, 1360, 300,fill="#fb0")

#Add Generate Filter Button
Btn_GenerateFilters = Button(app_root,text='Generate Filters > > >',command=openNewWindow)
Win_BtnGenerateFiltersWindow = app_canvas.create_window(80,330,window=Btn_GenerateFilters)

#Add separator line
app_canvas.create_line(1, 355, 1360, 355,fill="#fb0")

# function to open a new window
# on a button click




'''
#Build Filter Options
FilterOptions1 = ["Priority","Impact","Region","Day","Generic"]
FilterOptions2 = ["Priority","Impact","Region","Day","Generic"]
DefaultValueFilter1 = StringVar(app_root)
DefaultValueFilter1.set(FilterOptions1[0])
DefaultValueFilter2 = StringVar(app_root)
DefaultValueFilter2.set(FilterOptions2[0])
#Build filter1
app_canvas.create_text(97,250,text="Enter First Graph Details:",font=("comicssansns",10,"bold"),fill='White')
ColumnName1 = Entry (app_root)
app_canvas.create_window(80, 278, window=ColumnName1)
FilterOption1 = OptionMenu(app_root,DefaultValueFilter1, *FilterOptions1)
FilterOption1.config(width=8, font=('Helvetica', 10))
app_canvas.create_window(75, 314, window=FilterOption1)
#Build filter2
app_canvas.create_text(375,250,text="Enter Second Graph Details:",font=("comicssansns",10,"bold"),fill='White')
ColumnName2 = Entry (app_root)
app_canvas.create_window(350, 275, window=ColumnName2)
FilterOption2 = OptionMenu(app_root,DefaultValueFilter2, *FilterOptions2)
FilterOption2.config(width=8, font=('Helvetica', 10))
app_canvas.create_window(345, 315, window=FilterOption2)
#Build Graph Options
BGOptions1 = ["PieChart","BarGraph","WordCloud"]
BGOptions2 = ["PieChart","BarGraph","WordCloud"]
DefaultValueBG1 = StringVar(app_root)
DefaultValueBG1.set(BGOptions1[0])
DefaultValueBG2 = StringVar(app_root)
DefaultValueBG2.set(BGOptions2[0])
#Build Graph dropdowns
BGOption1 = OptionMenu(app_root,DefaultValueBG1, *BGOptions1)
BGOption1.config(width=8, font=('Helvetica', 10))
app_canvas.create_window(75, 360, window=BGOption1)
BGOption2 = OptionMenu(app_root,DefaultValueBG2, *BGOptions2)
BGOption2.config(width=8, font=('Helvetica', 10))
app_canvas.create_window(345, 360, window=BGOption2)
#Add separator line
app_canvas.create_line(1, 390, 1360, 390,fill="#fb0")
# Add Submit button
SubmitBtn = Button(app_root,text='< < < Submit > > >',command=buildgraph)
SubmitBtn_window = app_canvas.create_window(220,450,window=SubmitBtn)
'''

#Add separator line
app_canvas.create_line(1, 500, 1360, 500,fill="#fb0")

app_root.mainloop()
