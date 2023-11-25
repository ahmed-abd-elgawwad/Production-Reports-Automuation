import pandas as pd
import pyodbc
from process_file import OneFileFetch



class DBInsert:

    def __init__(self):

        #initialize the connection into the database ( int this case access database )
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=(the file path of your local database.accdb);"
        )
        cnxn = pyodbc.connect(conn_str, autocommit=True)
        self.crsr = cnxn.cursor()

    def insert_production_data(self, prod_data):
        try:
            # insert the rows from the dataframe into a table named "produciton_table"
            self.crsr.executemany(
                "INSERT INTO production_table ( columns ) VALUES (?,?,?,?,?)",
                prod_data.itertuples(index=False)
            )
            print("Produciton Data inserted into Access database successfully.")

        except pyodbc.IntegrityError:
            print("The Produciton data for this Report is already inserted.")

   
