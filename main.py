import re
import pandas as pd
import numpy as np
import sqlite3


def read_data(file_name):
    """
    Reads in data from a xlsx file and returns a dataframe dictionary
    """
    xlsx = pd.ExcelFile(file_name)
    return {sheet_name: xlsx.parse(sheet_name) for sheet_name in xlsx.sheet_names}


def save_sqlite(data, db_name):
    """
    Saves data to a sqlite database
    """
    conn = sqlite3.connect(db_name)

    for sheet_name, df in data.items():
        df.to_sql(sheet_name, conn, if_exists='replace', index=False)
    conn.close()

def getDateRange(table, start, end):
    """
    Returns a specific date range with total min max columns from a table
    """
    conn = sqlite3.connect('data.db')
    df = pd.read_sql_query("SELECT SUM(Total) FROM '" + table + "' WHERE Date > '"+start+"' AND Date < '"+ end + "'" , conn)
    minDate = pd.read_sql_query("SELECT Date, MIN(Total) FROM '" + table + "' WHERE Date > '"+start+"' AND Date < '"+ end + "'" , conn)
    maxDate = pd.read_sql_query("SELECT Date, MAX(Total) FROM '" + table + "' WHERE Date > '"+start+"' AND Date < '"+ end + "'" , conn)
    conn.close()
    return {'summary': df.values.tolist()[0][0], 'min': minDate.values.tolist()[0], 'max': maxDate.values.tolist()[0]}

def main():
    """
    Main function
    """
    data = read_data('TestCaseData.xlsx')
    save_sqlite(data, 'data.db')
    startDate = ''
    endDate = ''
    while True:
        startDate = input("Enter start date: YYYY-MM-DD: ")
        if re.match(r'^\d{4}-\d{2}-\d{2}$', startDate):
            break
        else:
            print("Invalid date format")
    while True:
        endDate = input("Enter start date: YYYY-MM-DD: ")
        if re.match(r'^\d{4}-\d{2}-\d{2}$', endDate):
            break
        else:
            print("Invalid date format")
            
    result = getDateRange(list(data.keys())[0], startDate, endDate)
    print("Date Range: " + startDate + " to " + endDate)
    print("Total: " + str(result['summary']))
    print("Min: " + str(result['min'][0])[:10]+": " +str(result['min'][1]))
    print("Max: " + str(result['max'][0])[:10]+": " +str(result['max'][1]))

if __name__ == '__main__':
    main()
