from pathlib import Path
import win32com.client 
import datetime
import os

class DownloadFile:

    def __init__(self):
        # connecting to outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # get the inbox folder
        inbox  =  outlook.GetDefaultFolder(6)

        # get all the messages
        self.messages = inbox.Items

        # Sort the items by received time in descending order
        self.messages.Sort("[ReceivedTime]", True)

        # initiate the file
        self.file_path = None

    def save_file(self, location):
            
        # initialize the counter and the breaking criteria
        counter = 0
        file_found = False

        for message in self.messages:
                
            if file_found:
                break

            # get mail info 
            subject = message.Subject
            body = message.body
            attachments =  message.Attachments

            # check if it is a daily report email
            if "Daily Production Report" in subject:

                for attachment in attachments:

                    # check to save only the excel file not any other attachments
                    if attachment.FileName.endswith(".xlsx"):

                        # get the mail sent date
                        sent_date = (message.SentOn).date()
                        todays_date = datetime.date.today()

                        if sent_date == datetime.date(2023,11,20):
                            file_found = True
                            print(f"Found Today's File , Date sent {sent_date}")
                        
                            # check if it is todays report
                            self.file_path = attachment.FileName
                            attachment.SaveAsFile(os.path.join(location,attachment.FileName ) )
                            counter +=1 
                            print(f"Emails Searched = { counter } emails")
                            break
                        else:
                            print("No Report found for this date")

        return self.file_path

                    
if __name__ == "__main__":
    location = r"C:\Users\asayed\Desktop\Email_Fetching".replace("\\","/")
    downloader = DownloadFile()
    downloader.save_file(location=location)