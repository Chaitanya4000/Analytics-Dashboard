#>>>>>>>>>>>>>>> Analytics Dashboard <<<<<<<<<<<<<<<

#>>>>>>>>>>>>>>> Import all required libraries <<<<<<<<<<<<<<<
from tkinter import *
from PIL import ImageTk,Image
import time
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog
from datetime import datetime
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_DATA_LABEL_POSITION
from pptx.util import Cm
#from wordcloud import WordCloud, STOPWORDS, Imagecolourgenerator

#>>>>>>>>>>>>>>> Initialize Variables <<<<<<<<<<<<<<<
str_OutputFolderPath = ""
str_InputFilePath = ""
str_InputFileName = ""
str_OutputFileName = ""
Li_ColumnNames = []
varx = []
vary = []
dic_chkvar = {}
dic_dropdwnvar = {}

#>>>>>>>>>>>>>>> All functions that are used in applications <<<<<<<<<<<<<<<

#KillMainApp function will kill the main app window
def KillMainApp():
    app_root.destroy()

#GenerateFilters function is called by generate filters button from main app.
def GenerateFilters():
    global str_OutputFileName
    global str_OutputFolderPath
    str_OutputFolderPath = str(OutputFolderPath.get())
    str_OutputFileName = os.path.basename(str_OutputFolderPath)
    print("Output folder path is: "+str_OutputFolderPath)
    print("Output file name is: "+str_OutputFileName)
    KillMainApp()
    OpenFilterWindow()



# Function for opening the file
def file_opener():
    global str_InputFileName
    global str_InputFilePath
    str_InputFilePath = filedialog.askopenfilenames(
        parent=app_root,
        initialdir='/',
        initialfile='tmp',
        filetypes=[("Excel", "xlsx")])
    print(str_InputFilePath)
    str_InputFilePath = str(str_InputFilePath).replace("(","").replace(")","").replace("'","").replace(",","")
    str_InputFileName = os.path.basename(str_InputFilePath)

    if (str_InputFilePath != ""):
        app_canvas.create_text(270, 200, text=str_InputFileName, font=("comicssansns", 10, "bold"), fill='White')


#OpenFilterWindow fundction is called to display all column names and filters for graphs.

def OpenFilterWindow():

    def BuildGraphs():
        int_counter = 0
        print("Build graph function executing...")

        pptx = Presentation()
        first_slide_layout = pptx.slide_layouts[int_counter]
        slide = pptx.slides.add_slide(first_slide_layout)
        slide.shapes.title.text = "< < < Data Analysis > > >"
        slide.placeholders[1].text = " -Created by Analytics Dashboard."

        for column in Li_ColumnNames:
            if (dic_chkvar["ChkVar_"+column].get()== 1):
                print("Selected Column name is : " + column)
                print("Selected graph value for respective column is : "+str(dic_dropdwnvar["DropdwnVar_"+column].get()))


                if (dic_dropdwnvar["DropdwnVar_"+column].get() == "BarGraph"):

                    print("Output Folder Path: " + str_OutputFolderPath)
                    print("Output File Name: "+ str_OutputFileName)

                    data = pd.read_excel(str_InputFilePath)
                    UniqueValuesCount1 = pd.value_counts(data[column])
                    UniqueValues1 = data[column].dropna().unique().tolist()

                    slide = pptx.slides.add_slide(pptx.slide_layouts[int_counter])
                    chart_data = CategoryChartData()
                    chart_data.categories = UniqueValues1
                    chart_data.add_series('Series 1', UniqueValuesCount1)
                    x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
                    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
                    slide.shapes.title.text = "Graph as per "+column

                    int_counter = int_counter + 1

                if (dic_dropdwnvar["DropdwnVar_" + column].get() == "PieChart"):
                    print("Output Folder Path: " + str_OutputFolderPath)
                    print("Output File Name: " + str_OutputFileName)

                    data = pd.read_excel(str_InputFilePath)
                    UniqueValuesCount1 = pd.value_counts(data[column])
                    UniqueValues1 = data[column].dropna().unique().tolist()

                    slide = pptx.slides.add_slide(pptx.slide_layouts[int_counter])
                    slide.shapes.title.text = "Graph as per " + column
                    chart_data = CategoryChartData()
                    chart_data.categories = UniqueValues1
                    chart_data.add_series('Series 1', UniqueValuesCount1)

                    x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
                    chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE,  x, y, cx, cy, chart_data).chart

                    chart.has_legend = True
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                    chart.legend.include_in_layout = False

                    chart.plots[0].has_data_labels = True
                    data_labels = chart.plots[0].data_labels
                    data_labels.number_format = '0.0000%'
                    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END


                    int_counter = int_counter + 1


                if (dic_dropdwnvar["DropdwnVar_" + column].get() == "WordCloud"):
                    data = pd.read_excel(str_InputFilePath)
                    InputColumn = data[column].dropna()
                    str_InputText = " "
                    for row in InputColumn:
                        str_InputText = str_InputText + " " + str(row)
                    print(str_InputText)
                    mask = np.array(Image.open("Images\\cloud.png"))
                    #stopwords = set(STOPWORDS)
                    #wc = WordCloud(background_color="black",mask=mask,max_words=200,stopwords=stopwords)
                    #wc.generate(str_InputText)
                    #wc.to_file("Images\\wc.png")

        pptx.save(str_OutputFolderPath)


    #Get list of all columns in excel
    global Li_ColumnNames
    print("ExcelFilePath: " + str_InputFilePath)
    str_ExcelFilepath = str_InputFilePath
    data = pd.read_excel(str_ExcelFilepath)
    Li_ColumnNames = data.columns



    # Building a new window
    newWindow = Tk()

    # sets the geometry of toplevel
    newWindow.geometry("1300x700")
    newWindow.maxsize(1300, 700)
    newWindow.minsize(1300, 700)

    #Add title to the main window
    newWindow.title("Analytics Dashboard by T-systems")
    newWindow.configure(background="black")

    # Add backgroundImage
    img_BgImg = ImageTk.PhotoImage(Image.open("Images\\BgImg.PNG"))
    newWindow_canvas = Canvas(newWindow, width=1080, height=2160)
    newWindow_canvas.pack(fill="both", expand=True)
    newWindow_canvas.create_image(0, 0, image=img_BgImg, anchor="nw")

    # build title and project name
    newWindow_canvas.create_text(680, 35, text="Analytics Dashboard", font=("comicssansns", 20, "bold"), fill='White')
    LogoImg = ImageTk.PhotoImage(Image.open("Images\\TSysLogo.PNG"))
    newWindow_canvas.create_image(680, 100, image=LogoImg)

    # Add separator line
    newWindow_canvas.create_line(1, 150, 2000, 150, fill="#fb0")

    # build input filepath input box
    newWindow_canvas.create_text(170, 200, text="Selected input file name is : "+str_InputFileName, font=("comicssansns", 10, "bold"),
                                    fill='White')

    # build input filepath input box
    newWindow_canvas.create_text(170, 260, text="Selected output file name is : "+str_OutputFileName, font=("comicssansns", 10, "bold"),
                                    fill='White')

    # build label Column Name
    newWindow_canvas.create_text(740, 170, text="Select Graph Type::", font=("comicssansns", 10, "bold"),
                                    fill='White')

    # build label Column Name
    newWindow_canvas.create_text(540, 170, text="Select Column Names:", font=("comicssansns", 10, "bold"), fill='White')


    # Submit button on new window
    Btn_Submit = Button(newWindow, text='<<< Reset <<<', command=BuildGraphs)
    Win_Btn_Submit = newWindow_canvas.create_window(75, 380, window=Btn_Submit)

    # Submit button on new window
    Btn_Submit = Button(newWindow, text='>>> Submit >>>', command=BuildGraphs)
    Win_Btn_Submit = newWindow_canvas.create_window(200, 380, window=Btn_Submit)

    #Logic to display columns as a filter options.

    x = 1
    int_Rowcounter = 220
    int_ColumnCounter = 550
    int_Dropdowncounter = 0
    BGOptions1 = ["PieChart", "BarGraph", "WordCloud"]

    for i in Li_ColumnNames:

        dic_chkvar["ChkVar_{0}".format(i)] = IntVar()
        dic_dropdwnvar["DropdwnVar_{0}".format(i)] = StringVar(newWindow)
        vary = list(dic_dropdwnvar)
        varx = list(dic_chkvar)
        print(varx[x-1])
        print(vary[x - 1])


        Chk_Button = Checkbutton(newWindow_canvas, text=i, variable=dic_chkvar[varx[x-1]], width=20)
        newWindow_canvas.create_window(int_ColumnCounter, int_Rowcounter, window=Chk_Button)
        dic_dropdwnvar[vary[x-1]].set(BGOptions1[0])

        BGOption1 = OptionMenu(newWindow, dic_dropdwnvar[vary[x-1]], *BGOptions1)
        BGOption1.config(width=8, font=('Helvetica', 10))
        newWindow_canvas.create_window(int_ColumnCounter+180, int_Rowcounter , window=BGOption1)

        x = x + 1
        int_Rowcounter = int_Rowcounter + 50
        int_Dropdowncounter = int_Dropdowncounter + 50
        if (x==10) :
            # build label Column Name
            newWindow_canvas.create_text(1130, 170, text="Select Graph Type:",
                                            font=("comicssansns", 10, "bold"),
                                            fill='White')

            # build label Column Name
            newWindow_canvas.create_text(950, 170, text="Select Column Names:", font=("comicssansns", 10, "bold"),
                                            fill='White')

            int_ColumnCounter = int_ColumnCounter+400
            int_Rowcounter = 220

        if (x==21):
            break

    newWindow.mainloop()

#>>>>>>>>>>>>>>>>>>>> Build Main Application <<<<<<<<<<<<<<<<<<<<<<<<<<
app_root = Tk()

#GUI Framework
app_root.geometry("500x400")
app_root.maxsize(500,400)
app_root.minsize(500,400)

app_root.title("Analytics Dashboard by T-systems")
app_root.configure(background="black")

#Add backgroundImage
image1=ImageTk.PhotoImage(Image.open("Images\\BgImg.PNG"))
app_canvas = Canvas(app_root,width=1080,height=2160)
app_canvas.pack(fill="both",expand=True)
app_canvas.create_image(0,0,image=image1,anchor="nw")

#build title and project name
app_canvas.create_text(250,35,text="Analytics Dashboard",font=("comicssansns",20,"bold"),fill='White')
LogoImg = ImageTk.PhotoImage(Image.open("Images\\TSysLogo.PNG"))
app_canvas.create_image(250, 100, image=LogoImg)


#Add separator line
app_canvas.create_line(1, 150, 1360, 150,fill="#fb0")

#build input filepath input box
app_canvas.create_text(85,170,text="Enter Input File Path:",font=("comicssansns",10,"bold"),fill='White')


# Browse Button label
Btn_Browse = Button(app_root, text ='Browse & Select file', command = lambda:file_opener())
Win_BtnBrowseWindow = app_canvas.create_window(75,200,window=Btn_Browse)

#Add separator line
app_canvas.create_line(1, 230, 1360, 230,fill="#fb0")

#build output folder input box
app_canvas.create_text(95,250,text="Enter Output Folder Path:",font=("comicssansns",10,"bold"),fill='White')
OutputFolderPath = Entry (app_root,width=50)
app_canvas.create_window(170, 280, window=OutputFolderPath)

#Add separator line
app_canvas.create_line(1, 300, 1360, 300,fill="#fb0")

#Add Generate Filter Button
Btn_GenerateFilters = Button(app_root,text='Generate Filters > > >',command=GenerateFilters)
Win_BtnGenerateFiltersWindow = app_canvas.create_window(80,330,window=Btn_GenerateFilters)

#Add separator line
app_canvas.create_line(1, 355, 1360, 355,fill="#fb0")

#Add separator line
app_canvas.create_line(1, 500, 1360, 500,fill="#fb0")

app_root.mainloop()