#Version: 2
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os # for file operations

workingFolder =  os.getcwd()

DTS = pd.read_excel(workingFolder + '\\' + 'DTS.xlsx')


suffix = '_DTS_Corrected'
timeshifts = []
previousTime = datetime(2000,1,1)
convert_mm_to_m_all = True

 
warningStrings = []

for f in os.listdir(workingFolder):

    #if (f[-5:]==".xlsx" or f.endswith('csv')) and not suffix in f and f != 'DTS.xlsx':
    if (f[-5:]==".xlsx") and not suffix in f and f != 'DTS.xlsx':

        print('Process ' + f)
        
        convert_mm_to_m = convert_mm_to_m_all
        
        if f.endswith('.csv'):
            df1 = pd.read_csv(workingFolder + '\\' + f)
        else:
            df1 = pd.read_excel(workingFolder + '\\' + f)
            
        # Find the index of the first date
        for i in range(10):
            if type(df1.iloc[i,0])==datetime:
                first_date_index = i
                break
        
        # Add empty rows if the first date index is less than 8
        if first_date_index is not None and first_date_index < 8:
            rows_to_add = 8 - first_date_index
            empty_rows = pd.DataFrame(np.nan, index=range(rows_to_add), columns=df1.columns)
            #df1 = pd.concat([empty_rows, df1], ignore_index=True)
            df1 = pd.concat([df1.iloc[:first_date_index], empty_rows, df1.iloc[first_date_index:]], ignore_index=True)

        if 'level_col_index' in locals():
            del level_col_index
        for i in range(4):
            if df1.iloc[1,i].lower() == 'm' or df1.iloc[1,i].lower() == 'mm':
               df1.iloc[6,i] = 'Level'
               level_col_index = i
               if convert_mm_to_m and df1.iloc[1,i].lower() == 'm':
                   convert_mm_to_m = False #Since it is already in m.
                   
            elif df1.iloc[1,i].lower() == 'l/s' or df1.iloc[1,i].lower() == 'l/sec':
                df1.iloc[6,i] = 'Discharge'
            elif df1.iloc[1,i].lower() == 'm/s' or df1.iloc[1,i].lower() == 'meter/sec':
                df1.iloc[6,i] = 'Velocity'
        
        if not 'level_col_index' in locals():
            convert_mm_to_m = False #Since the column was not found.
                            
        
        df1.insert(4,'OriginalTime','')
        df1.insert(5,'ShiftedTimeFlag','')
        df1.insert(6, 'TimeShiftMarker', '')
        
        

        for index, row in df1[8:].iterrows():

            oYear = row[0].year
            DTS_row = DTS[DTS['Year']==oYear]
            DTS_Begin = DTS_row.iloc[0]['Begin']
            DTS_End = DTS_row.iloc[0]['End']

            df1.at[index, 'OriginalTime'] = row[0]

            if index > 8:
                timestep = (row[0]-previousTime).total_seconds()
                if abs(timestep) > 50 * 60 or timestep < 0:
                    timeshifts.append([row[0],timestep])

            if row[0] >= DTS_Begin and row[0] < DTS_End:
                df1.at[index, 'ShiftedTimeFlag'] = True

                df1.iat[index, 0] = row[0] + timedelta(hours=1)

            else:

                df1.at[index, 'ShiftedTimeFlag'] = False
                
            if convert_mm_to_m:
                df1.iloc[index, level_col_index] = df1.iloc[index, level_col_index] / 1000
                
            if index > 8:
                prev_shifted_time_flag = df1.at[index - 1, 'ShiftedTimeFlag']
                current_shifted_time_flag = df1.at[index, 'ShiftedTimeFlag']
                if prev_shifted_time_flag == True and current_shifted_time_flag == False:
                    df1.at[index, 'TimeShiftMarker'] = 'DTS End'
                elif prev_shifted_time_flag == False and current_shifted_time_flag == True:
                    df1.at[index, 'TimeShiftMarker'] = 'DTS Begin'
                else:
                    df1.at[index, 'TimeShiftMarker'] = ''
                    
                

            previousTime = row[0]

        #Remove first hour at duplicate
        df1 = pd.concat([df1.iloc[:8], df1.iloc[8:][~df1.iloc[8:, 0].duplicated(keep='last')]], ignore_index=True)
        
        # # Add empty rows to substitute the missing hour at daylight savings start
        # if first_date_index is not None and first_date_index < 8:
        #     rows_to_add = 8 - first_date_index
        #     empty_rows = pd.DataFrame(np.nan, index=range(rows_to_add), columns=df1.columns)
        #     df1 = pd.concat([df1.iloc[:first_date_index], empty_rows, df1.iloc[first_date_index:]], ignore_index=True)
        
        df1.iloc[3,0] = 'At daylight savings end, the first hour is deleted.' 
        df1.iloc[1,6] = 'CTRL+SHIFT+DOWN'
        df1.iloc[2,6] = 'to jump to'
        df1.iloc[3,6] = 'next time shift'
        
        if convert_mm_to_m:
            df1.iloc[4,0] = '*Source Level read as mm and converted to m. Please verify source.' 
            df1.iloc[1,level_col_index] = '*m'
           
        if len(timeshifts)>0:
            warningString = 'Warning! Timeshift detected in ' + f + '\n'
            for timeshift in timeshifts:
                warningString += str(timeshift[0]) + ", " + str(timeshift[1]) + " minutes\n"
                warningStrings.append(warningString)

# =============================================================================
#             MessageBox = ctypes.windll.user32.MessageBoxA
#             if MessageBox(None, warningString + "\nDo you still wish to perform DTS shift?", 'Info', 4) == 7:
#                 continue
# =============================================================================

        #else:
        df1.to_excel(workingFolder + '\\' + f[:-5] + suffix + '.xlsx',index=False)

warning_df = pd.DataFrame(warningStrings,columns =['Warning'],index=None)
warning_df.to_csv(workingFolder + '\\' + 'Warnings.csv',index=None)



