

import pickle  # OBJECT SERIALIZATION
import os  # ACCESS FOLDERS
import glob  # ACCESS FOLDERS
import gspread  # GSHEETS API
from googleapiclient.discovery import build  # GOOGLE AUTHORIZATION
from oauth2client.service_account import ServiceAccountCredentials  # GOOGLE AUTHORIZATION


'''
yt_item() - YouTube Item Object
Holds informations about a video
'''


class yt_item():
    def __init__(self):
        self.video_id = ''
        self.views = ''
        self.watch_time = ''
        self.subs = ''
        self.impress = ''
        self.video_name = ''


'''
Google Sheets API
'''
class Goole_Sheets_API(object):

    def __init__(self, sheet_id, sheet_name):

        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly',
                       'https://www.googleapis.com/auth/spreadsheets',
                       'https://www.googleapis.com/auth/drive',
                       'https://www.googleapis.com/auth/drive.readonly',
                       'https://www.googleapis.com/auth/drive.file']

        # SET UP PATH TO THE AUTH KEY!!!
        self.CLIENT_SECRET_FILE = os.environ['GSHEETS_AUTH_KEY']
        self.items_list = []
        self.sh = None
        self.obj = None
        self.sheet_name = sheet_name
        self.gsheetId = sheet_id

    def get_authenticated_services(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.CLIENT_SECRET_FILE,
            scopes=self.SCOPES)
        service = gspread.authorize(credentials)
        self.sh = service.open_by_key(self.gsheetId)
        print('THE SERVICE OBTAINED')

    def get_items(self):
        os.chdir("data")
        self.get_pkl()

    def get_pkl(self):
        cwd = os.getcwd()
        for file in glob.glob(f"*.pkl"):
            # print(file)
            with open(f"{cwd}/{file}", 'rb') as token:
                item = yt_item()
                item = pickle.load(token)
                self.items_list.append(item)
                print(f"Vide Name: {item.video_name} Views: {item.views}")

    def populate_sheet(self):
        self.get_authenticated_services()
        self.get_items()
        counter = 1
        worksheet1 = self.sh.worksheet(self.sheet_name)

        header = ["Video Name", "Video Id", "Watch Time",
                  "Subscribers", "Impressions", "Views"]
        index = 1
        worksheet1.insert_row(header, index)

        for item in self.items_list:

            #print(f"Views: \t\t\t\t\t\t\t{item.views}")
            worksheet1.update(
                f'A{counter+1}', [[str(item.video_name)]])
            worksheet1.update(
                f'B{counter+1}', [[str(item.video_id)]])
            worksheet1.update(
                f'C{counter+1}', [[str(item.watch_time)]])
            worksheet1.update(
                f'D{counter+1}', [[str(item.subs)]])
            worksheet1.update(
                f'E{counter+1}', [[str(item.impress)]])
            worksheet1.update(
                f'F{counter+1}', [[str(item.views)]])
            print(
                f'ITEM {item.video_name} added to {self.sheet_name} successfully!')
            counter += 1


def main():

    sheet_id = '1LvLuJl_HyD4AiuuKF5-cLYh0d7yUE2T6UBRxgkPtj58'  # CHANGE TO YOUR GSHEET ID
    sheet_name = 'Sheet1'
    obj = Goole_Sheets_API(sheet_id, sheet_name)
    obj.populate_sheet()


if __name__ == '__main__':
    main()
