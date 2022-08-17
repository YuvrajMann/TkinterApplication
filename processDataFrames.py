from http import client
from tkinter import messagebox
import pandas as pd

#---------------------------- Constant Values ----------------------------#
NO_FILE_SELECTED="Not File Selected"
NO_FOLDER_SELECTED="No Folder Selected"

def processDataFrame(
    clappedHoursDataframe,
    clientActualAccessDataframe,
    clientCustomReportDataframe,
    numberOfWeeks,
    selectUploadFolder
    ):
    # Null vaue checking
    if not isinstance(clappedHoursDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide client clapped hours file')
        return
    if not isinstance(clientActualAccessDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide client actual access file')
        return
    if not isinstance(clientCustomReportDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide CRIB sheet')
        return
    if numberOfWeeks==0:
        messagebox.showerror('Processing Error', f'Error: Please provide number of weeks-Client clapped')
        return
    if selectUploadFolder==NO_FOLDER_SELECTED:
        messagebox.showerror('Processing Error', f'Error: Please provide output folder directory')
        return

    outputFrameColumns=["Employee ID","Employee Name","Employee Level","Total Working Hours"]
    for client in clappedHoursDataframe['Client Name']:
        outputFrameColumns.append(client)
    outputFrameColumns.append('Total Hours Charged by Employee')
    outputFrameColumns.append('Hours Worked - Hours Charged')
    outputFrameColumns.append('Ratio')

    gk=clientCustomReportDataframe.groupby('employee')

    outputFrameData=[]

    for idx in clientCustomReportDataframe.index:
        totalHrs=gk.get_group(clientCustomReportDataframe['employee'][idx])['total_time'].agg('sum')
        outputSheetRow=[clientCustomReportDataframe['id'][idx],
            clientCustomReportDataframe['employee'][idx],
            clientCustomReportDataframe['Level'][idx],
            totalHrs
        ]
        for client in clappedHoursDataframe['Client Name']:
            outputSheetRow.append(0)
        outputSheetRow.append(0)
        outputSheetRow.append(0)
        outputSheetRow.append(0)

        outputFrameData.append(outputSheetRow)

    outputDataFrame=pd.DataFrame(outputFrameData,columns=outputFrameColumns)
    print(outputDataFrame)
    outputDataFrame.to_excel(f'{selectUploadFolder}/HoursToCharge.xlsx')
    messagebox.showinfo(f'Processing Done',f'Successfully processed , you can find output file here - {selectUploadFolder}/HoursToCharge.xlsx')