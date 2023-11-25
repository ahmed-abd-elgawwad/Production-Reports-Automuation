import os
import shutil
from datetime import datetime
import re

class FileTansfer:

    def __init__(self, source_folder , destination_base_folder):

        self.source_folder = source_folder
        self.destination_base_folder = destination_base_folder

        # file path
        self.file_path = ""

    def transfer(self):

        filename = os.listdir(self.source_folder)[0]

        # check it is a daily produciton report
        if filename.startswith("The of your daily report ( any part of it )"):
            # Regular expression pattern to extract date from the filename
            date_pattern = r"(\d{2}-\d{2}-\d{4})"
            match = re.search(date_pattern, filename)
            if match:
                date_string = match.group(1)
            file_date = datetime.strptime(date_string, '%d-%m-%Y').date()

            # Create destination folders based on the file's date
            year_folder = os.path.join(self.destination_base_folder, str(file_date.year))
            month_folder = os.path.join(year_folder, f"{file_date.month:02d}-{file_date.strftime('%B')}")
            day_folder = os.path.join(month_folder, file_date.strftime('%B.%d.%Y'))

            # Create destination folders if they don't exist
            os.makedirs(year_folder, exist_ok=True)
            os.makedirs(month_folder, exist_ok=True)
            os.makedirs(day_folder, exist_ok=True)

            # Move the file to the appropriate folder
            source_file_path = os.path.join(self.source_folder, filename)
            destination_file_path = os.path.join(day_folder, filename)
            shutil.move(source_file_path, destination_file_path)

            self.file_path = destination_file_path

            # show the results
            print(f" daily report of date [{ str(file_date) }] is moved successfully.")
    
