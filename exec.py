#!/usr/bin/env python
# coding: utf-8

# In[1]:


import sys
sys.dont_write_bytecode = True
sys.path.append("/work/src/mysql")
sys.path.append("/work/src/excel")

from mysql_connector import MysqlConnector
from export import Excel

# connect = MysqlConnector()
# header, rows = connect.fetch('select * from admin_roles')
# print(rows)
spreadsheet = Excel()
spreadsheet.export('maintenances')


# In[ ]:




