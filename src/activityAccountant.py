import pandas as pd
import os
import math

REGISTRANT_SUBDIR = "registrantExports/"
EVENT_SUBDIR = "eventExports/"
OUTPUT_SUBDIR = "scoring/"
SCORE_FILE = "scoring.xlsx"

# We don't count events whose dates are older than a certain amount
MAXIMUM_EVENT_AGE = pd.DateOffset(years=3)


class Event:
    def __init__(self, name, date, activityPoints):
        self.name = name
        self.date = date
        self.activityPoints = activityPoints

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
        for event in self.attended:
            str += f"\t{event}\n"
        return str


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

    def addUniqueEvent(self, name, date, pointCount):
        name = name.strip()
        if not self.eventMap.__contains__(name):
            self.eventMap[name] = Event(name, date.strip(), int(pointCount))
        return self.eventMap[name]

    def printAttendees(self):
        for key in self.userMap:
            print(self.userMap[key])

    def printEvents(self):
        for key in self.eventMap:
            print(self.eventMap[key])

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
        eventDir = self.inputBaseDir + EVENT_SUBDIR
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
                self.addUniqueEvent(eventName, eventBeginDate, activityPoints)

    def buildAttendeeList(self):
        registrantDir = self.inputBaseDir + REGISTRANT_SUBDIR
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
                attendee.attended.append(str(sheet["Event"].iloc[0]))
                # diagnostic print
                # print(f"\t Attended by {attendee.firstName} {attendee.lastName} - {attendee.email}: total is {attendee.points} points")

    def assignPoints(self):
        # iterate over events, and assign points to every user that has that event
        for eventName, event in self.eventMap.items():
            for attendeeEmail, attendee in self.userMap.items():
                if eventName in attendee.attended:
                    attendee.points += event.activityPoints

    def exportResults(self):
        scoringDir = self.outputBaseDir + OUTPUT_SUBDIR
        firstNames = list()
        lastNames = list()
        emails = list()
        points = list()
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
        os.makedirs(os.path.dirname(scoringDir), exist_ok=True)
        dataFrame.to_excel(scoringDir + SCORE_FILE)


if __name__ == "__main__":
    accountant = Accountant("./test/", "/tmp/")
    accountant.buildEventList()
    accountant.buildAttendeeList()
    # accountant.printEvents()
    accountant.assignPoints()
    # accountant.printAttendees()
    accountant.exportResults()
