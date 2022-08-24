# Language - Python v3.10.6
# Framework used for developing desktop application - customtkinter v0.3

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
import os

#---------------------------- Window Configuration ----------------------------#

customtkinter.set_appearance_mode("dark")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
    
window=customtkinter.CTk() #Initializing root window
window.title('Hours to charge calculator')
window.geometry('660x560')
window.minsize(660, 560)

#---------------------------- Constant Values ----------------------------#

NO_FILE_SELECTED="Not File Selected"
NO_FOLDER_SELECTED="No Folder Selected"

#---------------------------- Data ----------------------------#

# Uploaded Files Name
UploadClappedActionFileName=NO_FILE_SELECTED #The clapped hours file location
UploadAccessActualActionFileName=NO_FILE_SELECTED #The actual access file location
UploadCustomReportActionFileName=NO_FILE_SELECTED # The umanity custom report file location
selectUploadFolder=NO_FOLDER_SELECTED #The upload folder location

#uploaded Files Data frames
clappedHoursDataframe=0 #The clapped hours file
clientActualAccessDataframe=0 #The actual access file
clientCustomReportDataframe=0 # The umanity custom report file
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

# Function to select output folder
def selectUploadFolderUtil(event=None):
    try:
        global selectUploadFolder
        selectUploadFolder = filedialog.askdirectory()
        if selectUploadFolder=="":
            selectUploadFolder=NO_FOLDER_SELECTED
    except:
        selectUploadFolder = NO_FOLDER_SELECTED

    label_5.config(text=selectUploadFolder)

# Function selects number of weeks-client clapped
def clientClappedNumberOfWeeks():
    global numberOfWeeks
    dialog = customtkinter.CTkInputDialog(master=None, text="Select number of weeks - Client Clapped :", title="Select number of weeks - Client Clapped")
    numberOfWeeks=dialog.get_input()
    label_6.config(text=(f"Number of weeks - {numberOfWeeks}"))

#Utility function used to process all selected input files and calculate cost
def framesProcess():
    processDataFrames.processDataFrame(window,clappedHoursDataframe,clientActualAccessDataframe,clientCustomReportDataframe,numberOfWeeks,selectUploadFolder)

#---------------------------- UI Elements ----------------------------#

# Whole UI divided into two frames - Upper and lower frame

#Configuring and placing upper frame
frameUpper = customtkinter.CTkFrame(master=window,
                               width=300,
                               corner_radius=10)
frameUpper.pack(fill="x",padx=40,pady=40)

#Configuring and placing the lower frame
frameLower = customtkinter.CTkFrame(master=window,
                               width=110,
                               corner_radius=10)

frameLower.pack(fill="x",padx=40)

# Text displaying - select input files
label_1 = customtkinter.CTkLabel(master=frameUpper,
                                              text="Select Input Files",
                                              text_font=("Roboto Medium", -16)) 

label_1.pack(fill="x",pady=10,padx=4)

# Configuring and placing button for selecting client clapped hours file
clientClappedbutton = customtkinter.CTkButton(master=frameUpper,fg_color="#1597BB",hover_color="#107793", text="Client Capped Hours File", command=UploadClappedAction)
clientClappedbutton.pack(pady=5,padx=20)

# Label displaying the clapped hours file upload status
label_2 = customtkinter.CTkLabel(master=frameUpper,
                                              text=UploadClappedActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_2.pack(pady=1,padx=4)

# Configuring and placing button for selecting client actual access file
clientAccessActual = customtkinter.CTkButton(master=frameUpper, fg_color="#1597BB",hover_color="#107793",text="Client Access File", command=UploadAccessActualAction)
clientAccessActual.pack(pady=5,padx=20)

# Label displaying the client actual access file upload status
label_3 = customtkinter.CTkLabel(master=frameUpper,
                                              text=UploadAccessActualActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_3.pack(pady=1,padx=4)

# Configuring and placing button for selecting umanity custom report file
clientCustomReport=customtkinter.CTkButton(master=frameUpper,text="Humanity custom report",fg_color="#1597BB",hover_color="#107793",command=UploadCustomReportAction)
clientCustomReport.pack(pady=5,padx=20)

# Label displaying the umanity custom report file upload status
label_4 = customtkinter.CTkLabel(master=frameUpper,
                                              text=UploadCustomReportActionFileName,
                                              text_font=("Roboto Medium", -11)) 

label_4.pack(fill="x",pady=1,padx=4)

# Configuring and placing button for selecting selecting the number of weeks
weekSelect = customtkinter.CTkButton(master=frameUpper, text="Number of weeks",fg_color="#1597BB",hover_color="#107793", command=clientClappedNumberOfWeeks)
weekSelect.pack(pady=5,padx=20)

# Label displaying the number of weeks
label_6 = customtkinter.CTkLabel(master=frameUpper,
                                              text="Number of weeks - 0",
                                              text_font=("Roboto Medium", -11))
label_6.pack(fill="x",pady=1,padx=4)

# Configuring and placing button for selecting selecting the output folder
selectOutputFolder=customtkinter.CTkButton(master=frameUpper,text="Select Output Folder",fg_color="#1597BB",hover_color="#107793",command=selectUploadFolderUtil)
selectOutputFolder.pack(pady=5,padx=20)

# Label displaying the output folder
label_5 = customtkinter.CTkLabel(master=frameUpper,
                                              text=selectUploadFolder,
                                              text_font=("Roboto Medium", -11)) 

label_5.pack(fill="x",pady=1,padx=4)

# Configuring and placing button for processing input
processButton=customtkinter.CTkButton(master=frameLower,text="Calculate Hours", command=framesProcess)
processButton.pack(pady=20,padx=20)

#---------------- Configuration to root window for making it responsive ---------------------#
window.grid_rowconfigure(0, weight=1)
window.grid_rowconfigure(1, weight=1)
window.grid_rowconfigure(2, weight=1)
window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)

window.mainloop()
