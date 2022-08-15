#!/usr/bin/env python
# coding: utf-8

# # BOM Comparer
# ### Take an old BOM sheet and a new BOM sheet and compares what is new or missing.

# In[1]:


# xlsx comparer
# purpose: find difference between two xlsx files and indicate them in a new file.
# library to use: pandas, xlsxwriter
# Update 2.2.1: 
# + Added exit functionality; kernel will now automatically close upon closing window.
from datetime import datetime
import pandas as pd
import xlsxwriter
import PySimpleGUI as sg
pd.set_option('display.max_columns', 10)

# Helper method to edit cells based on met conditions. 
# Input: series containing the values of old and new 
# Output: string to replace the cell.
def report_diff(x):
    if x[0] == x[1]:
        return x[0]
    elif x.isnull()[0]:
        return 'ADDED ---> %s' % x[1]
    elif x.isnull()[1]:
        return '%s ---> REMOVED' % x[0]
    else:
        return '{} ---> {}'.format(*x)

def file_compare(old_file, new_file, uid):
    NoUIDRaise = False
    # *** Read the files, old vs new ***
    old = pd.read_excel(old_file, sheet_name=None)
    new = pd.read_excel(new_file, sheet_name=None)
    now = str(datetime.now())
    date_now = ('output_' + now.split(' ')[0] + '_' + now.split(' ')[1].split('.')[0].replace(':', '_'))
    finalWriter = pd.ExcelWriter('%s.xlsx' % date_now, 
                                engine='xlsxwriter')
    # *** for loop to iterate through multiple pages in the excel file.
    # Assumes old and new have the same page length and contents for each page ***
    for i in range(len(list(old))):
        old_df = (list(old.values())[i])
        new_df = (list(new.values())[i])
        # Creating a new DataFrame df using
        # 1. merge and the flags required for a comparison table (indicator & how).
        # 2. a lambda function that returns only values NOT present in both old and new (!='both')
        df = (old_df.merge(new_df, indicator=True, how='outer')).loc[lambda v: v['_merge'] != 'both']
        df.reset_index(inplace=True, drop=True)
        df.rename(columns={'_merge': 'version'}, inplace=True)
        df['version'] = df['version'].replace(['left_only'], 'old')
        df['version'] = df['version'].replace(['right_only'], 'new')

        # This code was derived from: https://pbpython.com/excel-diff-pandas-update.html
        # Make new DataFrames based on the categories of 'old' and 'new' from the base DataFrame.
        old_sep = df[(df['version'] == 'old')]
        new_sep = df[(df['version'] == 'new')]

        old_sep.reset_index(inplace=True, drop=True)
        new_sep.reset_index(inplace=True, drop=True)
    
        old_sep = old_sep.drop(['version'], axis=1)
        new_sep = new_sep.drop(['version'], axis=1)
    
        # Set uid as the index for proper indexing of rows. 
        # Any column with static, unique data can work for uid.
        if uid in df:
            NoUIDRaise = True
            old_sep.set_index(old_sep[uid], inplace=True)
            new_sep.set_index(new_sep[uid], inplace=True)

        # Make a new DataFrame by combining old_sep and new_sep. This makes side-by-side tables
        # of the old and new sheets.
        df = pd.concat([old_sep, new_sep], axis='columns', keys=['old', 'new'], join='outer')

        # Swap the levels between the descriptor categories (Manufacturer #, Designator, LibRef, etc. ) 
        # and the version categories (old & new). The result is a df that displays old and new versions of
        # each descriptor in side-by-side cells rather than side-by-side tables.
        df = df.swaplevel(axis='columns')[new_sep.columns[0:]]
        # Makes a df that groups cells by descriptor category(level=0) and column(axis=1), turning each group into a 2-item series 
        # containing the old and new items. 
        # Performs report_diff to these series and applies the result to the cell within df.groupby.
        df_changed = df.groupby(level=0, axis=1).apply(lambda frame: frame.apply(report_diff, axis=1))
        df_changed = df_changed.reset_index(drop=True)
        df = df_changed
    
        # CODE FOR HIGHLIGHTING THE SHEET, gold = modified item, red = removed item, green = new item.
        # TODO: Highlight the sheet in way that does not interfere with text wrap formatting.
        if uid in df:
            df_styled = df.style                        .applymap(lambda x: 'background-color: gold' if '--->' in str(x) else '')                        .apply(lambda x: ['background-color: salmon' if 'REMOVED' in x else '' for x in df[uid]], axis=0)                        .apply(lambda x: ['background-color: mediumseagreen' if 'ADDED' in x else '' for x in df[uid]], axis=0)
        else:
            df_styled = df.style                        .applymap(lambda x: 'background-color: gold' if '--->' in str(x) else '')
    
        #display(df_styled)
        
        # Put 'df_styled' instead of 'df' to get a highlighted sheet.
        df_styled.to_excel(finalWriter, 'Sheet%d' % (i + 1))  # send df to writer
        
    finalWriter.save()
    if NoUIDRaise == False:
        return 'Inputted UID %s was not found in files. Possible faulty output: %s' % (uid, date_now) + '.xlsx' 
    
    return 'Comparison complete. Output file saved as %s' % date_now + '.xlsx'


# In[2]:


### Heavily modified from PySimpleGUI Demo Programs: https://github.com/PySimpleGUI/PySimpleGUI/blob/master/DemoPrograms/Demo_Compare_Files.py ###

sg.theme('Dark Blue 3')

def xlsx_check(string): 
        return list(reversed(string.split(".")))[0]
    
def main():

    form_rows = [[sg.Text('Enter 2 files to compare. \nEnter in the name of a column that identifies unique items (Default: LibRef).')],
                     [sg.Text('File 1', size=(15, 2)),
                        sg.InputText(key='-file1-'), sg.FileBrowse()],
                     [sg.Text('File 2', size=(15, 1)), sg.InputText(key='-file2-'),
                      sg.FileBrowse(target='-file2-')],
                     [sg.Text('UID (Default: LibRef)', size=(15, 2)),
                        sg.InputText(key='-uid-')],
                     [sg.Submit(), sg.Text('', key= '-OUTPUT-')]]
    window = sg.Window('BOM Comparison', form_rows)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        f1, f2, uid = values['-file1-'], values['-file2-'], values['-uid-']
        if uid == '':
            uid = 'LibRef'
            
        if (xlsx_check(f1) != 'xlsx') | (xlsx_check(f2) != 'xlsx'):
            window['-OUTPUT-'].update('Please select 2 xlsx files.')
        else:
            message = file_compare(f1, f2, uid)
            window['-OUTPUT-'].update(message)
            
    window.close()
    exit()
    

if __name__ == '__main__':
    main()


# In[ ]:





# In[ ]:




