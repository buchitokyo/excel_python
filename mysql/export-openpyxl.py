#!/usr/bin/env python
# coding: utf-8

# In[2]:


import openpyxl
from openpyxl.styles.fonts import Font
from openpyxl.styles.borders import Border, Side
import mysql.connector as mysql
import warnings
import time

FILE: str = "GLOBIS学び放題フレッシャーズ集計エクセル.xlsx"

def init():
    cnx = None

    cnx = mysql.connect (
        host = "db",
        port = "3306" ,
        user = "root",
        password = "root",
        database = "globis-dev",
    )

    return cnx

def export(table_name):
    start = time.perf_counter()
    print('Start: ' + str(start))
    warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

    wb = openpyxl.load_workbook(FILE)
    
    # switch sheet
    # ws = wb.worksheets[2]
    wb, ws = switch(wb, 2)

    # fetch
    header, rows = fetch_data(table_name)
    
    # format
    formatData(ws, rows)

    wb.save(filename = FILE)
    wb.close()

    end = time.perf_counter()
    print('End: ' + str(start))
    print(end - start)

def switch(wb, num):
    ws = wb.worksheets[num]
    return wb, ws

def fetch_data(table_name):
    # cnx = None
    try:
        # The connect() constructor creates a connection to the MySQL server and returns a MySQLConnection object.
        cnx = init()

        if cnx.is_connected:
            print('DB Connected.')
        cursor = cnx.cursor(dictionary=True)
        cursor.execute('select * from ' + table_name)

        print('Data Fetching...')
        header = [row[0] for row in cursor.description]
        rows: List[Dict[str, Any]] = cursor.fetchall()

        # Closing connection
        cnx.close()
    except Exception as e:
        print("Error Occurred: {e}")
        print("Failed to connect with DB.")
    finally:
        if cnx is not None and cnx.is_connected():
            cnx.close()
            
    return header, rows
            
def formatData(ws, rows):
    row_index = 1
    row_index += 1
    for row in rows:
        column_index = 1

        # dict
        for column in row.values():
            # print(column)
            cell = ws.cell(row = row_index, column = column_index)
            cell.value = column

            column_index += 1
        row_index += 1

    print(str(row_index) + ' rows written successfully to ' + FILE)
    return

# Tables to be exported
export('maintenances')
# export('TABLE_2')


# In[ ]:




