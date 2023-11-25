from Today_file import DownloadFile
from process_file import OneFileFetch
from db_insert import DBInsert
from save_one_file import FileTansfer
import tempfile
'------------------------------ inputs ---------------------------------'
report_file_distination_folder = "D:/Daily-Production-Report"
share_distination = "S:/SAZ/Operations & Drilling/Ahmed Elsayed/Production Data/Daily Produciton Reports"

# create a temporary directory
with tempfile.TemporaryDirectory() as tmpdir:
    # First: download Today's file from outlook 
    downloader = DownloadFile()
    file_name = downloader.save_file(location=tmpdir)
    print(file_name)
    # Last move the file into it's location in the share
    transfer = FileTansfer(source_folder=tmpdir , destination_base_folder= share_distination)
    transfer.transfer()

    # Second: process the file downloaded and get the data 
    fetcher = OneFileFetch(transfer.file_path )
    prod_data,test_data,hps_data,tank_data,shipping_data = fetcher.fetch_all_data()

    print(prod_data,test_data,tank_data)

    # # Third : insert the data into the database
    db_inserter = DBInsert()
    db_inserter.insert_production_data(prod_data)
    db_inserter.insert_test_data(test_data)
    db_inserter.insert_tank_data(tank_data)






