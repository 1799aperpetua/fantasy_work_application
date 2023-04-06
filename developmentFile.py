import os
from openpyxl import load_workbook
import sqlite3

dir = '/Users/anthonyperpetua/Desktop/development/fantasy_hvac/test_lineups'

def loopThroughSpreadsheets(directory):
    '''
    This method loops through all of the files in the provided directory, accessing 
    each xlsx file and printing out the name of the person who submitted it
    '''
    print(f"Looping through excel files in the directory passed in")
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            #print("Workbook location:", f'{os.path.join(directory, filename)}')
            wb = load_workbook(f'{os.path.join(directory, filename)}')
            ws = wb['Lineup_Template']
            #print("User Submitting Data: ", ws["b2"].value)
            captureLineup(ws)
            wb.close()

    print("Program ran successfully!")
        
def captureLineup(worksheet):
    '''
    Loop through a person's lineup spreadsheet and capture their information:
        name, week, lineup
    '''
    name = worksheet['b2'].value
    week = worksheet['b3'].value

    data = []

    for row in worksheet['B5:B18']:
        for idx, cell in enumerate(row):
            data.append(cell.value)

    print(f'Week {week} Lineup for: ', name, "\n", data, "\n-----------------")

# loopThroughSpreadsheets(dir)

cwd = os.getcwd()

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
    except: return print("Failed to execute your SQL statement")

create_table_statement = '''
CREATE TABLE IF NOT EXISTS fantasy_entries (
idx TEXT PRIMARY KEY,
name TEXT,
week TEXT,
mgr TEXT,
ca1 TEXT,
ca2 TEXT,
prs1 TEXT,
prs2 TEXT,
opptotech1 TEXT,
opptotech2 TEXT,
flex1 TEXT,
flex2 TEXT,
flex3 TEXT,
flex4 TEXT,
branch TEXT
)
'''

createTable(connectDB(db_location), create_table_statement)