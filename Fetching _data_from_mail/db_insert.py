import pandas as pd
import pyodbc
from process_file import OneFileFetch



class DBInsert:

    def __init__(self):

        #initialize the connection into the database
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=D:/NPC DB/ProductionDB.accdb;"
        )
        cnxn = pyodbc.connect(conn_str, autocommit=True)
        self.crsr = cnxn.cursor()

    def insert_production_data(self, prod_data):
        try:
            # insert the rows from the dataframe into a table named "people"
            self.crsr.executemany(
                "INSERT INTO ProdDaily ( [Date], [CompletionID], [Pump_Type] , [InjLineChock] , [GrossProduction] , [WC] , [OilProd] , [CumOilYesterday] , [CumOilToday] , [InjectionRate] , [WHIP] , [FlowLineP] , [OilAPI] ,[PFS] , [GasProd] , [GOR] , [AvgDesalterTemp] , [RunningHours] , [Status] , [Remarks] ) VALUES ( ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? , ? )",
                prod_data.itertuples(index=False)
            )
            print("Produciton Data inserted into Access database successfully.")

        except pyodbc.IntegrityError:
            print("The  Produciton data for this Report is already inserted.")

    def insert_test_data(self, test_data):
        try:
            # insert the rows from the dataframe into a table named "people"
            self.crsr.executemany(
                "INSERT INTO TestFetched ( [Date], [CompletionID], [GrossReturn] , [OilProd] , [WC] , [Gross-WC]  ) VALUES ( ? , ? , ? , ? , ? , ? )",
                test_data.itertuples(index=False)
            )
            print("Test Data inserted into Access database successfully.")

        except pyodbc.IntegrityError:
            print("This Test data is already inserted.")

    def insert_tank_data(self , tank_data):
        try:
            # insert the rows from the dataframe into a table named "people"
            self.crsr.executemany(
                "INSERT INTO Tanks ( [Date], [Tank Name],[Closed Stock], [Opening Stock] , [Shipping] , [Total Drain Water] , [Total Oil Received From MOPU] ) VALUES ( ? , ? , ? , ? , ? , ?,? )",
                tank_data.itertuples(index=False)
            )
            print("Tanks data inserted succesfully.")

        except pyodbc.IntegrityError:
            print("tanks data aleary exist")

        except Exception as e :
            print(str(e))
        
    
    def insert_hps_data(self, hps_data):
        pass
    
    def insert_shipping_data(self, shipping_data ):
        pass

if __name__ == "__main__":

    file_path = r"C:\Users\asayed\Desktop\Email_Fetching\PZN_Daily_Production__Report 03-10-2023.xlsx".replace("\\","/")
    fetcher = OneFileFetch(file_path)
    prod_data,test_data,hps_data,tank_data,shipping_data = fetcher.fetch_all_data()
    print(prod_data)

    db_inserter = DBInsert()
    db_inserter.insert_production_data(prod_data)