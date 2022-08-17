from calendar import week
from cgitb import text
from doctest import master
from tkinter import *
from tkinter import filedialog
from tkinter.tix import COLUMN
import customtkinter
import pandas as pd
import processDataFrames
from tkinter import messagebox

#---------------------------- Window Configuration ----------------------------#

customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
    
window=customtkinter.CTk()
window.title('Hours to charge calculator')
window.geometry('660x560')
window.minsize(660, 560)
#---------------------------- Constant Values ----------------------------#

NO_FILE_SELECTED="Not File Selected"
NO_FOLDER_SELECTED="No Folder Selected"

#---------------------------- Data ----------------------------#

# Uploaded Files Name
UploadClappedActionFileName=NO_FILE_SELECTED
UploadAccessActualActionFileName=NO_FILE_SELECTED
UploadCustomReportActionFileName=NO_FILE_SELECTED
selectUploadFolder=NO_FOLDER_SELECTED

#uploaded Files Data frames
clappedHoursDataframe=0
clientActualAccessDataframe=0
clientCustomReportDataframe=0
numberOfWeeks=0

#---------------------------- Utility Function ----------------------------#

# Function for uploading client clapped hours file
def UploadClappedAction(event=None):
    global UploadClappedActionFileName 
    global clappedHoursDataframe
    try:
        UploadClappedActionFileName = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        df = pd.read_excel (UploadClappedActionFileName)
        clappedHoursDataframe=df
    except:
        UploadClappedActionFileName=NO_FILE_SELECTED

    label_2.config(text=UploadClappedActionFileName)

# Function for uploading client's actual access file
def UploadAccessActualAction(event=None):
    global UploadAccessActualActionFileName
    global clientActualAccessDataframe
    try:
        UploadAccessActualActionFileName = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        df = pd.read_excel (UploadAccessActualActionFileName)
        clientActualAccessDataframe=df
    except:
        UploadAccessActualActionFileName=NO_FILE_SELECTED

    label_3.config(text=UploadAccessActualActionFileName)

# Function for uploading client's custom report
def UploadCustomReportAction(event=None):
    global UploadCustomReportActionFileName
    global clientCustomReportDataframe
    try:
        global UploadCustomReportActionFileName
        global clientCustomReportDataframe
        UploadCustomReportActionFileName = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        df = pd.read_excel (UploadCustomReportActionFileName)
        clientCustomReportDataframe=df
    except:
        UploadCustomReportActionFileName=NO_FILE_SELECTED

    label_4.config(text=UploadCustomReportActionFileName)

# Function selects upload folder
def selectUploadFolderUtil(event=None):
    global selectUploadFolder
    selectUploadFolder = filedialog.askdirectory()
    label_5.config(text=selectUploadFolder)

# Function selects number of weeks-client clapped
def clientClappedNumberOfWeeks():
    global numberOfWeeks
    dialog = customtkinter.CTkInputDialog(master=None, text="Select number of weeks - Client Clapped :", title="Select number of weeks - Client Clapped")
    numberOfWeeks=dialog.get_input()
    label_6.config(text=(f"Number of weeks - {numberOfWeeks}"))

def framesProcess():
    processDataFrames.processDataFrame(clappedHoursDataframe,clientActualAccessDataframe,clientCustomReportDataframe,numberOfWeeks,selectUploadFolder)
#---------------------------- UI Elements ----------------------------#
frameLeft = customtkinter.CTkFrame(master=window,
                               width=300,
                               corner_radius=10)
frameLeft.pack(fill="x",padx=40,pady=40)

frameRight = customtkinter.CTkFrame(master=window,
                               width=110,
                               corner_radius=10)

frameRight.pack(fill="x",padx=40)

label_1 = customtkinter.CTkLabel(master=frameLeft,
                                              text="Select Files",
                                              text_font=("Roboto Medium", -16)) 

label_1.pack(fill="x",pady=10,padx=4)

clientClappedbutton = customtkinter.CTkButton(master=frameLeft, text="Client Capped Hours File", command=UploadClappedAction)
clientClappedbutton.pack(fill="x",pady=5,padx=20)

label_2 = customtkinter.CTkLabel(master=frameLeft,
                                              text=UploadClappedActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_2.pack(fill="x",pady=1,padx=4)

clientAccessActual = customtkinter.CTkButton(master=frameLeft, text="Client Access File", command=UploadAccessActualAction)
clientAccessActual.pack(fill="x",pady=5,padx=20)
label_3 = customtkinter.CTkLabel(master=frameLeft,
                                              text=UploadAccessActualActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_3.pack(fill="x",pady=1,padx=4)

clientCustomReport=customtkinter.CTkButton(master=frameLeft,text="CRIB Sheet",command=UploadCustomReportAction)
clientCustomReport.pack(fill="x",pady=5,padx=20)

weekSelect = customtkinter.CTkButton(master=frameLeft, text="Number of weeks - Client Clapped", command=clientClappedNumberOfWeeks)
weekSelect.pack(fill="x",pady=5,padx=20)

label_4 = customtkinter.CTkLabel(master=frameLeft,
                                              text=UploadCustomReportActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_4.pack(fill="x",pady=1,padx=4)

label_6 = customtkinter.CTkLabel(master=frameLeft,
                                              text="Number of weeks - 0",
                                              text_font=("Roboto Medium", -11))
label_6.pack(fill="x",pady=1,padx=4)

selectOutputFolder=customtkinter.CTkButton(master=frameLeft,text="Select Output Folder",command=selectUploadFolderUtil)
selectOutputFolder.pack(fill="x",pady=5,padx=20)

label_5 = customtkinter.CTkLabel(master=frameLeft,
                                              text=selectUploadFolder,
                                              text_font=("Roboto Medium", -11)) 

label_5.pack(fill="x",pady=1,padx=4)

processButton=customtkinter.CTkButton(master=frameRight, text="Process Files", command=framesProcess)
processButton.pack(fill="x",pady=20,padx=20)


window.grid_rowconfigure(0, weight=1)
window.grid_rowconfigure(1, weight=1)
window.grid_rowconfigure(2, weight=1)
window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)

window.mainloop()
