#!/usr/bin/env python
# coding: utf-8

# In[1]:


import sys
sys.dont_write_bytecode = True
import openpyxl
from openpyxl.styles.fonts import Font
from openpyxl.styles.borders import Border, Side
import warnings
import time
from mysql_connector import MysqlConnector


class Excel:

    FILE: str = ""

    def __init__(self):
        warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

    def export(self, table_name):
        start = time.perf_counter()
        print('Start: ' + str(start))
        wb = openpyxl.load_workbook(self.FILE)

        # switch sheet
        wb, ws = Excel.switch(self, wb, 2)

        print("Open file.")

        # fetch
        sql = 'select * from ' + table_name
        connect = MysqlConnector()
        header, rows = connect.fetch(sql)
        # print(rows)

        # format
        Excel.formatData(self, ws, rows)

        wb.save(filename = self.FILE)
        wb.close()

        end = time.perf_counter()
        print('End: ' + str(start))
        print(end - start)

    def switch(self, wb, num):
        ws = wb.worksheets[num]
        return wb, ws

    def formatData(self, ws, rows):
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

        print(str(row_index) + ' rows written successfully to ' + self.FILE)
        return

