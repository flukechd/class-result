import pandas as pd
import requests
import pickle
import os
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
def Create_Service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    print(SCOPES)

    cred = None

    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'
    # print(pickle_file)

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None

ClIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'Drive'
API_VERSION='v3'
filename="123456"
sheetname="room2"
SCOPES = ['https://www.googleapis.com/auth/drive']

service = Create_Service(ClIENT_SECRET_FILE,API_NAME,API_VERSION,SCOPES)

folder_id = '1v5Dtp6m6OV1EgXB_jokyd9c0Wnj-Jsw4'
query = f"parents='{folder_id}'"

response = service.files().list(q=query).execute()
files = response.get('files')
nextPageToken = response.get('nextPageToken')

pd.set_option('display.max_columns', 100)
pd.set_option('display.max_rows', 500)
pd.set_option('display.min_rows', 500)
pd.set_option('display.max_colwidth', 150)
pd.set_option('display.width', 200)
pd.set_option('expand_frame_repr', True)
df = pd.DataFrame(files)
print("https://drive.google.com/open?id="+df.id)
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
scope = ["https://spreadsheets.google.com/feeds"
,'https://www.googleapis.com/auth/spreadsheets'
,"https://www.googleapis.com/auth/drive.file"
,"https://www.googleapis.com/auth/drive"]
cerds = ServiceAccountCredentials.from_json_keyfile_name("cerds.json", scope)
client = gspread.authorize(cerds)
checkworksheets =0
sheet = client.open(filename)
for i in range(len(sheet.worksheets())):
    if sheet.worksheets()[i].title==sheetname:
        checkworksheets =1
if checkworksheets==0:
    worksheet = sheet.add_worksheet(title=sheetname, rows="100", cols="20")
sheet = client.open(filename).worksheet(sheetname) # เป็นการเปิดไปยังหน้าชีตนั้นๆ
sheet.format('A1:E1', {'textFormat': {'bold': True}})
sheet.update_cell(1,1,"GoogleDrive file URL")
sheet.update_cell(1,2,"FileName")
sheet.update_cell(1,3,"Time")
sheet.update_cell(1,4,"Frame")
cell=sheet.cell(1,1).value
cell1=sheet.cell(2,1).value
pprint(cell)
for i in range(0,len(df.id)):
    sheet.update_cell(i+2,1,"https://drive.google.com/open?id="+df.id[i])
cell=sheet.cell(1,1).value
for i in range(0,len(df.name)):
    sheet.update_cell(i+2,2,df.name[i])
cell1=sheet.cell(2,1).value
pprint(cell)

gid = "edit#gid=" + str(sheet.id)
baseURL = "https://docs.google.com/spreadsheets/d/1ghQDOccY6b6F0HN_prVU0mMSe_SBCJYcnOPhs794yfw/"
fullURL = baseURL + gid
print(fullURL)



