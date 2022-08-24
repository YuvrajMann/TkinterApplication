# Pandas library in python is used for easing the process of excel files processing
# Please get a basic understanding of pandas before working on this file
     
from http import client
from tkinter import messagebox
import pandas as pd
import os
from threading import Thread

#---------------------------- Constant Values ----------------------------#
NO_FILE_SELECTED="Not File Selected"
NO_FOLDER_SELECTED="No Folder Selected"

def processDataFrame(
    root,
    clappedHoursDataframe,
    clientActualAccessDataframe,
    clientCustomReportDataframe,
    numberOfWeeks,
    selectUploadFolder
    ):

    ## -- Null vaue checking and popping error boxes corresponding to missing files -- ##
    if not isinstance(clappedHoursDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide client capped hours file')
        return
    if not isinstance(clientActualAccessDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide client actual access file')
        return
    if not isinstance(clientCustomReportDataframe,pd.DataFrame):
        messagebox.showerror('Processing Error', f'Error: Please provide humanity custom report file')
        return
    if numberOfWeeks==0:
        messagebox.showerror('Processing Error', f'Error: Please provide number of weeks')
        return
    if selectUploadFolder==NO_FOLDER_SELECTED:
        messagebox.showerror('Processing Error', f'Error: Please provide output folder directory')
        return
    # ------------------------------------------------------------------------------------#
    
    # Getting the columns in output data frame
    outputFrameColumns=["Employee ID","Employee Name","Employee Level","Total Working Hours"]
    for client in clappedHoursDataframe['Client Name']:
        outputFrameColumns.append(client)
    outputFrameColumns.append('Total Hours Charged by Employee')
    outputFrameColumns.append('Hours Worked - Hours Charged')
    outputFrameColumns.append('Ratio')
    # ----------------------------------------------------------------------------------------#

    # Grouping the data in cleint custom report for each employee by employee id to get total hours of each employee
    gk=clientCustomReportDataframe.groupby('employee')
    
    # The data/rows in output data frame
    outputFrameData=[]

    # Lopping for each row in client's custom report data frame  
    for idx in clientCustomReportDataframe.index:
        # Getting the total hours for particular employee using grouped by 'gk' dataframe
        totalHrs=gk.get_group(clientCustomReportDataframe['employee'][idx])['total_time'].agg('sum')
        
        # -- Gathering data for output sheet row -- # 
        outputSheetRow=[clientCustomReportDataframe['id'][idx],
            clientCustomReportDataframe['employee'][idx],
            clientCustomReportDataframe['Level'][idx].split()[0],
            totalHrs
        ]

        # Getting all distinct cleints
        for client in clappedHoursDataframe['Client Name']:
            # checking if 'cleint' is accessible by the employee using client actual access file
            rw=clientActualAccessDataframe[clientActualAccessDataframe['Id']==clientCustomReportDataframe['id'][idx]][client]
            accessible=rw.to_string(index=False)

            # setting 0 if not accessible
            if accessible=='N':
                outputSheetRow.append(0)
            else:
                outputSheetRow.append("")
                 
        # Dummy values for other 3 columns
        outputSheetRow.append(0)
        outputSheetRow.append(0)
        outputSheetRow.append(0)
        
        #------------------------------------------#

        #Appending row to otuput frame data
        outputFrameData.append(outputSheetRow)

    # Generating ouput data frame using earier gathered output frame columns and output frame data
    outputDataFrame=pd.DataFrame(outputFrameData,columns=outputFrameColumns)
    # Removing duplicate values in dataframe based on employee id
    outputDataFrame=outputDataFrame.drop_duplicates(subset=['Employee ID'],keep='first')

    # Dictionary stores for each cleint - the total working hours for each employee who have access to the client
    clientTotalHours={}

    #Initializing the dictionary 
    for client in clappedHoursDataframe['Client Name']:
            clientTotalHours[client]=0

    for idx in outputDataFrame.index:
        for client in clappedHoursDataframe['Client Name']:
            #checking if client is accessible for employee if it is we will add that employee's total working hours to dictionary
            if outputDataFrame[client][idx]!=0:
                clientTotalHours[client]+=outputDataFrame['Total Working Hours'][idx]

    # Calculate hours = (employee total hours/sum of all employees' total hours)*client capped hours
    # Caluclating the hours value for each employee and each client 
    for idx in outputDataFrame.index:
        for client in clappedHoursDataframe['Client Name']:
            if outputDataFrame[client][idx]!=0:
                clappedHours=clappedHoursDataframe[clappedHoursDataframe['Client Name']==client]['Capped Hours Per Week']
                clappedHours=(float)(clappedHours.to_string(index=False))
                calc=(outputDataFrame['Total Working Hours'][idx]/clientTotalHours[client])*(clappedHours)
                outputDataFrame.at[idx,client]=calc
    
    # Developing a function to round to a multiple
    def round_to_multiple(number, multiple):
        return multiple * round(number / multiple)

    #Dictionary for storing sum of cleint hours
    clientsSum={}
    #Dictionary for storing rounded sum of cleint hours
    roundedClientSum={}
    # Dictionary for storing difference in sum of cleint hours and rounded sum of cleint hours
    differenceDict={}

    #Initializing the dictionaries
    for client in clappedHoursDataframe['Client Name']:
            clientsSum[client]=0
            roundedClientSum[client]=0
            differenceDict[client]=0

    #Calculating rounded off value to 0.25 multiple and storing it in output frame
    for idx in outputDataFrame.index:
        for client in clappedHoursDataframe['Client Name']:
            clientsSum[client]+=outputDataFrame[client][idx]
            multiple_round=round_to_multiple(outputDataFrame[client][idx],0.25)
            roundedClientSum[client]+=multiple_round
            outputDataFrame.at[idx,client]=multiple_round

    # Computation performend - 
    # Find the sum of hours for each client 
    # If sum exceeds capped hours -> sort in ascending order of "total hours" of employees and start subtracting 0.25 from each individual. If the loop reaches the end with still some hours left to remove, have another iteration
    # If sum is less than capped hours -> sort in descending order of "total hours" of employees and start adding 0.25 to each individual. If the loop reaches the end with still some hours left to give, have another iteration

    for client in clappedHoursDataframe['Client Name']:
        differenceDict[client]=clientsSum[client]-roundedClientSum[client]
        differenceDict[client]=round(differenceDict[client],2)
        if differenceDict[client]>0:
            outputDataFrame=outputDataFrame.sort_values('Total Working Hours',ascending=False)
            while True:
                for idx in outputDataFrame.index:
                    if differenceDict[client]>=0.25 :
                        outputDataFrame.at[idx,client]=outputDataFrame[client][idx]+0.25
                        differenceDict[client]=differenceDict[client]-0.25
                    else:
                        break
                        
                if differenceDict[client]<0.25:
                    break;
        else:
            outputDataFrame=outputDataFrame.sort_values('Total Working Hours')
            while True:
                for idx in outputDataFrame.index:
                    if differenceDict[client]>=-0.25 :
                        outputDataFrame.at[idx,client]=outputDataFrame[client][idx]-0.25
                        differenceDict[client]=differenceDict[client]+0.25
                    else:
                        break

                if differenceDict[client]>-0.25:
                    break;
    
    # Sorting the output frame using index value
    outputDataFrame=outputDataFrame.sort_index()
    # Converting the data frame excel and saving it at provided output location with name - HoursToCharge.xlsx
    outputDataFrame.to_excel(f'{selectUploadFolder}/HoursToCharge.xlsx',index=False)
    # Success message dialog box
    messagebox.showinfo(f'Processing Done',f'Successfully processed , you can find output file here - {selectUploadFolder}/HoursToCharge.xlsx')
    # Opeing the final output file on a new thread
    # More about threads in os - https://www.geeksforgeeks.org/thread-in-operating-system/
    exceFileOpenThread = Thread(target = lambda: os.system(f'{selectUploadFolder}/HoursToCharge.xlsx'))
    exceFileOpenThread.start()
    