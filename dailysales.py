import pandas as pd
import numpy as np
import datetime as dt
import xlsxwriter
import csv
from openpyxl import load_workbook
import os

def create_sales_df(now):
    # Import Sales
    sales_date = now['month'] + '-' + now['day'] + '-' + now['year']
    sales_xl = pd.ExcelFile(sales_date + '.xlsx')
    sales_df = sales_xl.parse('Sheet1')

    # Remove Amounts that equal zero
    sales_df['Amount'] = sales_df['Amount'].fillna(0)
    sales_df = sales_df[sales_df['Amount'].isin([0])==False]
    
    # Import GL Codes
    gl_xl = pd.ExcelFile('gl-codes.xlsx')
    gl_df_base = gl_xl.parse('Base', index_col='(Item) Name')
    gl_df_site = gl_xl.parse('Site', index_col='Site')

    # Match Description to GL Codes (essentially VLOOKUP)
    sales_df['GL Account #'] = sales_df['(Item) Name'].map(gl_df_base['GL Account #'].astype(str))
    sales_df['Site_ID'] = sales_df['Site'].map(gl_df_site['Site_ID'].astype(str))
    
    # Items that aren't GL coded - add to gl-codes
    sales_df_miss = sales_df[sales_df['GL Account #'].isnull()][['(Item) Name', 'GL Account #']]
    sales_df_miss = sales_df_miss.set_index('(Item) Name').fillna('-')
    if sales_df_miss.empty == False:
        print(str(sales_df.ix[sales_df.head(1).index[0]]['End Date']) + ' has missing GL Codes')
        print(sales_df_miss)
        gl_df_base = gl_df_base.append(sales_df_miss).reset_index()
        
        book = load_workbook('gl-codes.xlsx')
        writer = pd.ExcelWriter('gl-codes.xlsx', engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        gl_df_base.to_excel(writer, 'Base', header=True, index=False)
        writer.save()
    
    # Sites that aren't GL coded - add to gl-codes
    sales_df_miss = sales_df[sales_df['Site_ID'].isnull()][['Site', 'Site_ID']]
    sales_df_miss = sales_df_miss.set_index('Site').fillna('-')
    if sales_df_miss.empty == False:
        print(str(sales_df.ix[sales_df.head(1).index[0]]['End Date']) + ' has missing GL Codes')
        gl_df_site = gl_df_site.append(sales_df_miss).reset_index()
        
        book = load_workbook('gl-codes.xlsx')
        writer = pd.ExcelWriter('gl-codes.xlsx', engine='openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        gl_df_base.to_excel(writer, 'Site', header=True, index=False)
        writer.save()


    # Verify Amount column reconciles
    sales_df_rec = round(sales_df['Amount'].sum(),2)
    if sales_df_rec != 0:
        print(str(sales_df.ix[sales_df.head(1).index[0]]['Start Date']) + ' is off by $' + str(sales_df_rec))
    
    # Create GL Upload Columns
    sales_df['RECORD'] = 'GLT'
    sales_df['ACCOUNT'] = sales_df['Site_ID'] + '-' + sales_df['GL Account #'] + '.000'
    sales_df['ACCNTG DATE'] = sales_df['End Date'].dt.strftime('%-m/%-d/%y')
    sales_df['JOURNAL'] = 10
    sales_df['REF 1'] = ''
    sales_df['REF 2'] = ''
    sales_df['DESCRIPTION'] = sales_df['(Item) Name'].str[0:30]
    sales_df['DEBIT'] = np.where(sales_df['Amount'] > 0, sales_df['Amount'],0)
    sales_df['CREDIT'] = np.where(sales_df['Amount'] < 0, sales_df['Amount'] * -1,0)
    sales_df['ACCRUAL OR CASH'] = 1
    
    # Finish Magic
    update_workbook(now, sales_df)

def update_workbook(now, sales_df):
    # Index Site Name
    site_name = sales_df.ix[sales_df.head(1).index[0]]['Site']

    # DS Filename/Path
    ds_filename = 'Daily Sales Entry - ' + site_name + ' - ' + now['month-name'] + ' 20' + now['year'] + '.xlsx'
    ds_filepath = cwd + '/' + ds_filename
    
    # Create DS File if it doesn't exist
    if os.path.exists(ds_filepath) == False:
        workbook = xlsxwriter.Workbook(ds_filename)
        worksheet = workbook.add_worksheet('1')
        workbook.close()
    
    # Copy old Workbook
    book = load_workbook(ds_filename)
    writer = pd.ExcelWriter(ds_filename, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    # Slice GL Upload Format
    upload_format = ['RECORD', 'ACCOUNT', 'ACCNTG DATE', 'JOURNAL', 'REF 1', 'REF 2', 'DESCRIPTION', 'DEBIT', 'CREDIT', 'ACCRUAL OR CASH']
    sales_df = sales_df[upload_format]

    # Save Workbook with new Data
    sales_df.to_excel(writer, now['day'], header=True, index=False)
    writer.save()

# Date Variables
def ds_start():
    # List all spreadsheets that aren't the main DS workbook
    files = []
    files += [each for each in os.listdir(cwd) if each.endswith('.xlsx') and 'Daily Sales' not in each and 'gl-codes' not in each]
    month_names = {'1': 'January', '2': 'February', '3': 'March',
                    '4': 'April', '5': 'May', '6': 'June',
                    '7': 'July', '8': 'August', '9': 'September',
                    '10': 'October', '11': 'November', '12': 'December'}

    dates = []
    for file in files:
        file = file.replace('.xlsx', '').split('-')
        file_date = {'month-name': month_names[str(file[0])],
                'month': file[0],
                'day': file[1],
                'year': file[2]}
        dates.append(file_date)

    for now in dates:
        create_sales_df(now)

# Main
if __name__ == '__main__':
    cwd = os.getcwd()
    ds_start()
