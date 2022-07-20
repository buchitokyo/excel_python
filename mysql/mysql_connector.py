#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import sys
sys.dont_write_bytecode = True
import mysql.connector as mydb
from importlib import import_module, reload


class MysqlConnector:
    def __init__(self, host='db', port='3306', user='root', password='root', database='globis-dev'):
        self.cnx = None
        try:
            self.cnx = mydb.connect(
                host=host,
                port=port,
                user=user,
                password=password,
                database=database
            )

            self.cnx.ping(reconnect=True)
            self.cnx.autocommit = False

            self.cursor = self.cnx.cursor(dictionary=True)
            if (self.cnx.is_connected()):
                print('Connected with DB.')

        except (mydb.errors.ProgrammingError) as e:
            print(e)
            print("Failed to connect with DB.")
            sys.exit(1)


    def fetch(self, sql):
        try:
            self.cursor.execute(sql)
            print('Data fetching...')

            header = [row[0] for row in self.cursor.description]
            rows: List[Dict[str, Any]] = self.cursor.fetchall()

            MysqlConnector._close_connection(self)

            return header, rows
        except mydb.Error as e:
            print('Data fetching error')
            print(e)
            sys.exit(1)


    def _close_connection(self) -> None:
        self.cursor.close
        self.cnx.close
        self.cursor = None
        self.cnx = None
        
        
    if __name__ == "__main__":
        print("mysql_connector")

