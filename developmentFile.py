import os
from openpyxl import load_workbook
import sqlite3

dir = '/Users/anthonyperpetua/Desktop/development/fantasy_hvac/test_lineups'

def loopThroughSpreadsheets(directory):
    '''
    This method loops through all of the files in the provided directory, accessing 
    each xlsx file and printing out the name of the person who submitted it
    '''

    def captureLineup(worksheet):
        '''
        Helper Function
        Loop through a person's lineup spreadsheet and capture their information: name, week, and lineup
        :param: worksheet - worksheet object stems from workbook object using the openpyxl library and an excel file
        :return: data - information user's submit as their weekly lineup - name, week, lineup choices
        '''
        name = worksheet['b2'].value
        week = worksheet['b3'].value

        data = [name, week]

        for row in worksheet['B5:B18']:
            for idx, cell in enumerate(row):
                data.append(cell.value)

        print(f'Week {week} Lineup for: ', name, "\n", data, "\n-----------------")
        return data
    
    def buildEntry(data):
        '''
        Helper Function
        Uses data captured in a spreadsheet to build a SQL entry
        '''

        idx = str(data[0]) + "(" + str(data[1]) + ")"

        statement = '''
        INSERT INTO fantasy_entries VALUES (
        :idx, :name, :week, :mgr, :ca1, :ca2, :prs1, :prs2, :opptotech1, :opptotech2, :flex1, :flex2, :flex3, :flex4, :branch
        )''' 
        context = {'idx':idx, 'name':data[0], 'week':data[1], 'mgr':data[2], 'ca1':data[3], 'ca2':data[4], 'prs1':data[5], 'prs2':data[6], 'opptotech1':data[7], 'opptotech2':data[8], 'flex1':data[9], 'flex2':data[10], 'flex3':data[11], 'flex4':data[12], 'branch':data[13]}


        return [statement, context]

    def submitEntry(conn, sql_statement, context):
        '''
        Helper Function
        Takes a SQL statement (a person's entry) and commits it to the database
        '''
        cur = conn.cursor()
        cur.execute(sql_statement, context)
        conn.commit()
        print("You submitted an entry to the database!")
        conn.close()

    print(f"Looping through excel files in the directory passed in")
    for filename in os.listdir(directory): # for each file in the passed directory
        if filename.endswith('.xlsx'): # when we encounter an excel file...
            wb = load_workbook(f'{os.path.join(directory, filename)}') # load the file as a workbook object
            ws = wb['Lineup_Template'] # access the lineup sheet
            data = captureLineup(ws) # Method that returns data from the spreadsheet, to be enterred into the database
            submitEntry(connectDB(db_location), buildEntry(data)[0], buildEntry(data)[1])
            wb.close()

    print("Program ran successfully!")
        
# loopThroughSpreadsheets(dir)

#cwd = os.getcwd()

db_location = 'fantasy_logDB.sqlite'
def connectDB(database):
    '''
    Connect to a local database
    '''
    try:
        conn = sqlite3.connect(database)
        return conn
    except:
        return print("Failed to connect to database")

def createTable(conn, sql_statement):
    '''
    Create a table for each entry from a spreadsheet
    index: name(week), name, week, lineup[] 
    '''
    try:
        cur = conn.cursor()
        cur.execute(sql_statement)
    except: 
        return print("Failed to execute your SQL statement")

create_table_statement = '''
CREATE TABLE IF NOT EXISTS fantasy_entries (
idx TEXT PRIMARY KEY,
name TEXT,
week TEXT,
mgr TEXT,
ca1 TEXT, ca2 TEXT,
prs1 TEXT, prs2 TEXT,
opptotech1 TEXT, opptotech2 TEXT,
flex1 TEXT, flex2 TEXT, flex3 TEXT, flex4 TEXT,
branch TEXT
)
'''

# createTable(connectDB(db_location), create_table_statement)

loopThroughSpreadsheets(dir)