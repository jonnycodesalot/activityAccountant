import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
from apiclient import http
import logging
import io

CREDSFILE = os.environ["GOOGLE_APPLICATION_CREDENTIALS"]


def createService():
    return build(
        "drive",
        "v3",
        credentials=service_account.Credentials.from_service_account_file(CREDSFILE),
    )


def getChildId(service, parentId, childName):
    q = "'" + parentId + "' in parents"
    results = (
        service.files()
        .list(
            pageSize=1000,
            q=q,
            fields="nextPageToken, files(id, name, mimeType)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    folder = results.get("files", [])
    for item in folder:
        if item["name"] == childName:
            return item["id"]
    return None


def updateSpreadsheet(service, fileId, localPath, remoteName=None):
    if not remoteName:
        remoteName = os.path.split(localPath)[-1]
    file_metadata = {
        "name": remoteName,
        "fileId": [fileId],
    }
    abspath = os.path.abspath(localPath)
    media = MediaFileUpload(
        abspath,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    service.files().update(
        fileId=fileId,
        body=file_metadata,
        media_body=media,
        supportsAllDrives=True,
    ).execute()


# To list folders
def downloadExcelDirectory(service, fileId, des, ignoreNames=[], recursive=True):

    #     var q = "mimeType = 'application/vnd.google-apps.folder' and '"+folderId+"' in parents";
    # var children = Drive.Files.list({q:q});
    q = "'" + fileId + "' in parents"
    results = (
        service.files()
        .list(
            pageSize=1000,
            q=q,
            fields="nextPageToken, files(id, name, mimeType)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    logging.debug(results)
    folder = results.get("files", [])
    logging.debug(folder)
    if not os.path.isdir(des):
        os.makedirs(des, exist_ok=True)
    for item in folder:
        if item["name"] not in ignoreNames:
            if str(item["mimeType"]) == str("application/vnd.google-apps.folder"):
                if recursive:
                    if not os.path.isdir(des + "/" + item["name"]):
                        os.makedirs(des + "/" + item["name"], exist_ok=True)
                    print(item["name"])
                    downloadExcelDirectory(
                        service,
                        item["id"],
                        des + "/" + item["name"],
                        recursive=recursive,
                    )  # LOOP un-till the files are found
            elif item["mimeType"] == str(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ):
                downloadExcel(service, item["id"], item["name"], des)
                print(item["name"])
            else:
                print(f"Skipping download of non-spreadsheet file {item['name']}")
        # else:
        #     print(f"Ignoring item {item["name"]} as requested.")
    return folder


# To Download Files
def downloadExcel(service, fileId, fileName, destDir):
    request = service.files().get_media(fileId=fileId)
    fh = io.BytesIO()
    downloader = http.MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    with io.open(destDir + "/" + fileName, "wb") as f:
        fh.seek(0)
        f.write(fh.read())


def main():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    print(os.environ["GOOGLE_APPLICATION_CREDENTIALS"])

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
        service = build(
            "drive",
            "v3",
            credentials=service_account.Credentials.from_service_account_file(
                CREDSFILE
            ),
        )

        downloadExcelDirectory(
            service, getFolderIdByName(service, "ActivityAccounting"), "/tmp"
        )
        uploadSpreadsheet(
            service,
            getFolderIdByName(service, "scoring"),
            "./test/eventExports/eventsList.xlsx",
        )
    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f"An error occurred: {error}")


def getFolderIdByName(service, name):
    # Call the Drive v3 API
    folderList = (
        service.files()
        .list(
            pageSize=10,
            q="mimeType = 'application/vnd.google-apps.folder' and name = '"
            + name
            + "'",
            fields="nextPageToken, files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        )
        .execute()
    )
    items = folderList.get("files", [])
    if not items:
        print("Shared folder '" + name + "' not found.")
        return
    if items.__len__() != 1:
        print(
            "There must be exactly one '"
            + name
            + "' folders shared with this user. Check sharing and try again."
        )
        return
    return items[0]["id"]


def uploadSpreadsheet(service, parentFolderId, localPath):
    try:
        file_metadata = {
            "name": os.path.split(localPath)[-1],
            "parents": [parentFolderId],
        }
        abspath = os.path.abspath(localPath)
        media = MediaFileUpload(
            abspath,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )
        # pylint: disable=maybe-no-member
        file = (
            service.files()
            .create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )
        print(f'File ID: "{file.get("id")}".')
        return file.get("id")

    except HttpError as error:
        print(f"An error occurred: {error}")
        return None


# if __name__ == "__main__":
#     main()
