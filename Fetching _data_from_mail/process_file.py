import pandas as pd
import openpyxl
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore")


# ------------------------------------ Fetching Functions -------------------------------------------------

class OneFileFetch:

    def __init__(self, file_path):

        self.file_path = file_path
        # reading the file
        wb =  openpyxl.load_workbook(self.file_path, read_only=True,data_only=True) 
        self.ws = wb['The Sheet Name']

    def fetch_production_data(self ):

        start = " the start cell of the data table"
        end = " the End cell of the data table"
        
        # convert the cells value into a dataframe
        production_data = self.ws[f"A{start}" : f"P{end}" ]
        rows_prod = [ [ value.value for value   in row ] for row in production_data ]
        d_prod = pd.DataFrame(rows_prod[1:] )

        # apply your processing on the data
        #---------------------
        
        return d_prod


    
