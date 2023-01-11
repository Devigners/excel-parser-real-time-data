from datetime import datetime
import xlwings as xw
import pandas as pd
import os

def read_excel():
    wb = xw.Book('Stocks Spread Calculation.xlsx')
    sheet = wb.sheets['MasterSheet']

    needed_cols = {'Sum Bid 1':'R', 'Strike 1':'S', 'Sum Bid 2':'W', 'Strike 2':'X', 'Sum Bid 3':'AB', 'Strike 3':'AC', 'Sum Bid 4':'AG', 'Strike 4':'AH', }
    col_values = {}

    for col_key in needed_cols.keys():
        col_A = sheet.range(needed_cols[col_key]+'10:'+needed_cols[col_key]+'49').options(ndim=1)
        values = col_A.value
        col_values[col_key] = values
    
    df = pd.DataFrame(col_values)
    df.to_csv('last_data.csv', index=False)
    if(not os.path.exists('last_data.csv')):
        print('[STATUS]: Saving data.')
    else:
        print('[STATUS]: Resaving data.')
    
    return df

# ## Creating last data file
def start():
    if(not os.path.exists('last_data.csv') or os.path.getsize('last_data.csv') == 0 ):
        print('[STATUS]: No previous records were found.')
        read_excel()
    else:
        print('[STATUS]: Previous record found.')
        
        # ## Creating needed dataframes
        df = pd.read_csv('last_data.csv')
        df_excel = read_excel()


        non_matching_indexes = df_excel[~df_excel[['Sum Bid 1', 'Sum Bid 2', 'Sum Bid 3', 'Sum Bid 4']].apply(tuple,1).isin(df[['Sum Bid 1', 'Sum Bid 2', 'Sum Bid 3', 'Sum Bid 4']].apply(tuple,1))].index
        if(len(non_matching_indexes.to_list()) > 0):
            print('[STATUS]: Changed rows found.')
            export_df = df_excel[~df_excel[['Sum Bid 1', 'Sum Bid 2', 'Sum Bid 3', 'Sum Bid 4']].apply(tuple,1).isin(df[['Sum Bid 1', 'Sum Bid 2', 'Sum Bid 3', 'Sum Bid 4']].apply(tuple,1))]
            export_df = export_df.copy()
            export_df['Time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            if(not os.path.exists('output.txt')):
                with open('output.txt', 'w') as file:
                    file.write('')

            export_df.to_csv('output.txt', sep=',', index=False, header=os.path.getsize('output.txt')==0, mode='a')
            print('[STATUS]: Record added to output.txt')
        else:
            print('[STATUS]: No data has changed.')