#!/usr/bin/env python
# coding: utf-8

# In[4]:


import sys
sys.dont_write_bytecode = True
sys.path.append("/work/mysql")
sys.path.append("/work/excel")
from mysql_connector import MysqlConnector
from export import Excel

spreadsheet = Excel()
spreadsheet.export('maintenances')


# In[ ]:




