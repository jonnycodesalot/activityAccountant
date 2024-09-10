import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from apiclient import http
import logging
import io

CREDSFILE = os.environ['GOOGLE_APPLICATION_CREDENTIALS']


# To list folders
def downloadDirectory(service, filid, des):
    
  #     var q = "mimeType = 'application/vnd.google-apps.folder' and '"+folderId+"' in parents";
  # var children = Drive.Files.list({q:q});
    q = "'"+filid+"' in parents"
    results = service.files().list(
        pageSize=1000, q=q,
        fields="nextPageToken, files(id, name, mimeType)",
              includeItemsFromAllDrives=True,
              supportsAllDrives=True).execute()
    logging.debug(results)
    folder = results.get('files', [])
    logging.debug(folder)
    for item in folder:
        if str(item['mimeType']) == str('application/vnd.google-apps.folder'):
            if not os.path.isdir(des+"/"+item['name']):
                os.mkdir(path=des+"/"+item['name'])
            print(item['name'])
            downloadDirectory(service, item['id'], des+"/"+item['name'])  # LOOP un-till the files are found
        elif (item['mimeType'] == str('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')):
            downloadSpreadsheet(service, item['id'], item['name'], des)
            print(item['name'])
        else:
           print(f"Skipping download of non-spreadsheet file {item['name']}")
    return folder


# To Download Files
def downloadSpreadsheet(service, dowid, name,dfilespath):
    request = service.files().get_media(fileId=dowid)
    fh = io.BytesIO()
    downloader = http.MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    with io.open(dfilespath + "/" + name, 'wb') as f:
        fh.seek(0)
        f.write(fh.read())

# def uploadSpreadsheet(service,)


def main():
  """Shows basic usage of the Drive v3 API.
  Prints the names and ids of the first 10 files the user has access to.
  """
  print(os.environ['GOOGLE_APPLICATION_CREDENTIALS'])


#   creds = None
#   # The file token.json stores the user's access and refresh tokens, and is
#   # created automatically when the authorization flow completes for the first
#   # time.
#   if os.path.exists("token.json"):
#     creds = Credentials.from_authorized_user_file("token.json", SCOPES)
#   # If there are no (valid) credentials available, let the user log in.
#   if not creds or not creds.valid:
#     if creds and creds.expired and creds.refresh_token:
#       creds.refresh(Request())
#     else:
#       flow = InstalledAppFlow.from_service_account_file(
#           "/tmp/activityAccountant/creds/credentials.json", SCOPES
#       )
#       creds = flow.run_local_server(port=0)
#     # Save the credentials for the next run
#     with open("token.json", "w") as token:
#       token.write(creds.to_json())

  try:
    service = build("drive", "v3", credentials=service_account.Credentials.from_service_account_file(CREDSFILE))

    
    downloadDirectory(service,getFolderIdByName(service,"ActivityAccounting"),'/tmp')
  except HttpError as error:
    # TODO(developer) - Handle errors from drive API.
    print(f"An error occurred: {error}")

def getFolderIdByName(service, name):
       # Call the Drive v3 API
    folderList = (
        service.files()
        .list(pageSize=10,
              q="mimeType = 'application/vnd.google-apps.folder' and name = '" + name + "'", 
              fields="nextPageToken, files(id, name)",
              includeItemsFromAllDrives=True,
              supportsAllDrives=True)
        .execute()
    )
    items = folderList.get("files", [])
    if not items:
      print("Shared folder '" + name + "' not found.")
      return
    if (items.__len__() != 1):
      print("There must be exactly one '" + name + "' folders shared with this user. Check sharing and try again.")
      return
    return items[0]['id']
   

if __name__ == "__main__":
  main()