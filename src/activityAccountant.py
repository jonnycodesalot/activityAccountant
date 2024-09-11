import pandas as pd
import os
import math
import googleDriveClient as gd
import datetime as dt

REGISTRANT_SUBDIR = "registrantExports/"
EVENT_SUBDIR = "eventExports/"
SCORE_FILE_SUFFIX = "scoring.xlsx"

# We don't count events whose dates are older than a certain amount
MAXIMUM_EVENT_AGE = pd.DateOffset(years=3)


class Event:
    def __init__(self, id, name, date, activityPoints):
        self.name = name
        self.date = date
        self.activityPoints = activityPoints
        self.id = int(id)

    def __str__(self):
        return f"{self.name} {self.date} - {self.activityPoints} points\n"


class Attendee:
    def __init__(self, firstName, lastName, email):
        self.firstName = firstName
        self.lastName = lastName
        self.email = email
        self.points = 0
        self.attended = list()

    def __str__(self):
        str = (
            f"{self.firstName} {self.lastName} - {self.points} points - {self.email}\n"
        )
        for eventId in self.attended:
            str += f"\tEvent {eventId}"
        return str

    def addEvent(self, eventId):
        if int(eventId) not in self.attended:
            self.attended.append(int(eventId))


class Accountant:
    def __init__(self, inputDir, outputDir):
        self.userMap = dict()
        self.eventMap = dict()
        self.inputBaseDir = inputDir
        self.outputBaseDir = outputDir

    def combineDuplicateUsers(self):
        # Placeholder. We're keying users on email address, but if we
        # have duplicates to deal with, we can combine records in the userMap
        # here before we assign points from the events list.
        pass

    def getUser(self, firstName, lastName, email):
        email = email.strip()
        if not self.userMap.__contains__(email):
            self.userMap[email] = Attendee(firstName.strip(), lastName.strip(), email)
        return self.userMap[email]

    def addUniqueEvent(self, id, name, date, pointCount):
        name = name.strip()
        if not self.eventMap.__contains__(id):
            self.eventMap[id] = Event(id, name, date.strip(), int(pointCount))
        return self.eventMap[id]

    def printAttendees(self):
        for key in self.userMap:
            print(self.userMap[key])

    def printEvents(self):
        for key in self.eventMap:
            print(self.eventMap[key].name)

    def openAndValidateSheet(self, directory, file):
        if (not file.endswith(".xlsx")) or file.startswith("~"):
            return None
        spreadsheet = pd.ExcelFile(directory + file)
        sheetCount = spreadsheet.sheet_names.__len__()
        if sheetCount != 1:
            raise Exception(
                f"File {file} has {sheetCount} sheets, but we only support single-sheet XLSX files.\n"
            )
        sheet = pd.read_excel(spreadsheet, spreadsheet.sheet_names[0])
        return sheet

    def buildEventList(self):
        eventDir = os.path.join(self.inputBaseDir, EVENT_SUBDIR)
        currTime = pd.to_datetime("now")
        for file in os.listdir(eventDir):
            sheet = self.openAndValidateSheet(eventDir, file)
            if sheet is None:
                continue
            for ndx in range(0, sheet.__len__()):
                activityPoints = sheet["activity_points"].iloc[ndx]
                if (activityPoints == 0) or math.isnan(activityPoints):
                    # No point looking at events with no point count
                    continue
                eventName = sheet["title"].iloc[ndx]
                eventEndDate = sheet["event_end_date"].iloc[ndx]
                if pd.to_datetime(eventEndDate) < (currTime - MAXIMUM_EVENT_AGE):
                    print(
                        f"Event {eventName}'s end date is older than the maximum event age. It will not be counted."
                    )
                    continue
                if pd.to_datetime(eventEndDate) > (currTime):
                    print(
                        f"Event {eventName} has not yet ended. It will not be counted."
                    )
                    continue
                eventBeginDate = sheet["event_date"].iloc[ndx]
                activityPoints = activityPoints
                self.addUniqueEvent(
                    sheet["id"].iloc[ndx], eventName, eventBeginDate, activityPoints
                )

    def buildAttendeeList(self):
        registrantDir = os.path.join(self.inputBaseDir, REGISTRANT_SUBDIR)
        for file in os.listdir(registrantDir):
            sheet = self.openAndValidateSheet(registrantDir, file)
            if sheet is None:
                continue
            print(f"Processing Registrant Export {file}...")
            for ndx in range(0, sheet.__len__()):
                if sheet["Payment Status"].iloc[ndx] != "Paid":
                    # Skip records that are cancelled or pending
                    continue
                attendee = self.getUser(
                    firstName=sheet["First Name"].iloc[ndx],
                    lastName=sheet["Last Name"].iloc[ndx],
                    email=sheet["Email"].iloc[ndx],
                )
                attendee.addEvent(sheet["Event ID"].iloc[ndx])
                # diagnostic print
                # print(f"\t Attended by {attendee.firstName} {attendee.lastName} - {attendee.email}: total is {attendee.points} points")

    def assignPoints(self):
        # iterate over events, and assign points to every user that has that event
        for eventId, event in self.eventMap.items():
            eventId = int(eventId)
            for attendeeEmail, attendee in self.userMap.items():
                if eventId in attendee.attended:
                    attendee.points += event.activityPoints

    def exportResults(self):
        firstNames = list()
        lastNames = list()
        emails = list()
        points = list()
        pd
        for attendeeEmail, attendee in self.userMap.items():
            firstNames.append(attendee.firstName)
            lastNames.append(attendee.lastName)
            emails.append(attendeeEmail)
            points.append(attendee.points)
        dataFrame = pd.DataFrame(
            {
                "First Name": firstNames,
                "Last Name": lastNames,
                "Email": emails,
                "ActivityPoints": points,
            }
        )
        os.makedirs(self.outputBaseDir, exist_ok=True)
        resultFilePath = os.path.join(
            self.outputBaseDir,
            dt.datetime.now().strftime("%Y-%m-%d-%H:%M:%S_") + SCORE_FILE_SUFFIX,
        )
        dataFrame.to_excel(resultFilePath)
        return resultFilePath


if __name__ == "__main__":
    gdService = gd.createService()
    localInputDir = "/tmp/activityAccountant/input"
    gd.downloadDirectory(
        gdService,
        gd.getFolderIdByName(gdService, "ActivityAccounting"),
        localInputDir,
    )
    localOutputDir = "/tmp/activityAccountant/results"
    accountant = Accountant(localInputDir, localOutputDir)
    accountant.buildEventList()
    accountant.buildAttendeeList()
    accountant.printEvents()
    accountant.assignPoints()
    # accountant.printAttendees()
    resultFilePath = accountant.exportResults()
    gd.uploadSpreadsheet(
        gdService,
        gd.getFolderIdByName(gdService, "scoring"),
        resultFilePath,
    )
