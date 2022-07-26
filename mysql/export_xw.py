#!/usr/bin/env python
# coding: utf-8

# In[1]:


import xlsxwriter
import mysql.connector


def fetch_table_data(table_name):
    # The connect() constructor creates a connection to the MySQL server and returns a MySQLConnection object.
    cnx = mysql.connector.connect(
        host='db',
        database='globis-dev',
        user='root',
        password='root'
    )

    cursor = cnx.cursor()
    cursor.execute('select * from ' + table_name)

    header = [row[0] for row in cursor.description]

    rows = cursor.fetchall()

    # Closing connection
    cnx.close()

    return header, rows


def export(table_name):
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('グロービス学び放題フレッシャーズ集計エクセル.xlsx')
    worksheet = workbook.add_worksheet('R_ ①コース単位')
    print(worksheet)
    # worksheet = workbook.add_worksheet('MENU')

    # Create style for cells
    header, rows = fetch_table_data(table_name)
    
    body_cell_format = workbook.add_format({'border': False})

    row_index = 0
    column_index = 0

    row_index += 1
    for row in rows:
        column_index = 0
        for column in row:
            worksheet.write(row_index, column_index, column, body_cell_format) # , body_cell_format
            column_index += 1
        row_index += 1

    print(str(row_index) + ' rows written successfully to ' + workbook.filename)

    # Closing workbook
    workbook.close()


# Tables to be exported
export('maintenances')
# export('TABLE_2')

