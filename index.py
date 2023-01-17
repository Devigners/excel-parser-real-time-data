from datetime import datetime
import xlwings as xw
import pandas as pd
import os
import json

def checkCondition(df, i, sum_col, strike_col):
    strikes = []
    
    if(i > 0 and df[sum_col][i-1] > 0):                             # former is positive
        if(df[sum_col][i] < 0):                                     # self is negative
            if(i < df.index.stop and df[sum_col][i+1] < 0):         # next is also negative
                strikes.append(df[strike_col][i])
    elif(i == 0):
        if(df[sum_col][i] < 0):                                     # self is negative
            if(i < df.index.stop and df[sum_col][i+1] < 0):         # next is also negative
                strikes.append(df[strike_col][i])
    
    return strikes

# ## Creating last data file
def start():
    wb = xw.Book('Stocks Spread Calculation.xlsx')
    sheet = wb.sheets['MasterSheet']

    needed_cols = {'Sum Bid 1':'R', 'Strike 1':'S', 'Sum Bid 2':'W', 'Strike 2':'X', 'Sum Bid 3':'AB', 'Strike 3':'AC', 'Sum Bid 4':'AG', 'Strike 4':'AH', }
    col_values = {}

    for col_key in needed_cols.keys():
        col_A = sheet.range(needed_cols[col_key]+'10:'+needed_cols[col_key]+'49').options(ndim=1)
        values = col_A.value
        col_values[col_key] = values
    
    df = pd.DataFrame(col_values)
    
    try:
        with open('output.txt', 'r') as f:
            # Read all the lines of the file
            lines = f.readlines()
            # Get the last line
            last_line = lines[-1]
            print('[STATUS]: Previous record found.')    
            file_strikes = json.loads(last_line)
    except:
        file_strikes = None
        print('[STATUS]: No previous records were found.')    
        
    strikes = {'1':[], '2':[], '3':[], '4':[]}

    for i in df.index:
        for j in range(4):
            col_strikes = checkCondition(df, i, 'Sum Bid '+str(j+1), 'Strike '+str(j+1))
            
            if len(col_strikes) > 0:
                strikes[str(j+1)].append(col_strikes)

    for key in strikes.keys():
        strikes[key] = [int(i[0]) for i in strikes[key]]
        

    if(strikes != file_strikes):
        print('[STATUS]: Change found.')
        with open('output.txt', 'a') as file:
            file.write(json.dumps(strikes))
            file.write('\n')
            print('[STATUS]: Record written to file.')    
    else:
        print('[STATUS]: No change found.')