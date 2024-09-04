import pandas as pd
import os
import math

#SOURCE_ROOT = 
ATTENDEE_EXPORTS = "./test/registrantExports/"
EVENT_EXPORTS = "./test/eventExports/"
OUTPUT_DIR = "/tmp/activityAccountant/"
SCORE_FILE = "scoring.xlsx"
EVENT_FILE = "eventsList.xlsx"

class Event:
    def __init__(self,name,date,activityPoints):
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
        str = f"{self.firstName} {self.lastName} - {self.points} points - {self.email}\n"
        for event in self.attended:
            str += f"\t{event}\n"
        return str

class Accountant:
    def __init__(self):
        self.userMap = dict()
        self.eventMap = dict()

    def combineDuplicateUsers(self):
        # Placeholder. We're keying users on email address, but if we
        # have duplicates to deal with, we can combine records in the userMap
        # here before we assign points from the events list.
        pass

    def getUser(self, firstName, lastName, email):
        email = email.strip()
        if (not self.userMap.__contains__(email)):
            self.userMap[email] = Attendee(firstName.strip(),lastName.strip(),email)
        return self.userMap[email]
    
    def addUniqueEvent(self, name, date, pointCount):
        name = name.strip()
        if (not self.eventMap.__contains__(name)):
            self.eventMap[name] = Event(name,date.strip(),int(pointCount))
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
            raise Exception(f"File {file} has {sheetCount} sheets, but we only support single-sheet XLSX files.\n")
        sheet = pd.read_excel(spreadsheet,spreadsheet.sheet_names[0])
        return sheet

    def buildEventList(self):
        for file in os.listdir(EVENT_EXPORTS):
            sheet = self.openAndValidateSheet(EVENT_EXPORTS, file)
            if (sheet is None):
                continue
            for ndx in range(0,sheet.__len__()):
                activityPoints = sheet['activity_points'].iloc[ndx]
                if (activityPoints == 0) or math.isnan(activityPoints):
                    # No point looking at events with no point count
                    continue
                activityPoints = activityPoints
                eventName = sheet['title'].iloc[ndx]
                eventDate = sheet['event_date'].iloc[ndx]
                self.addUniqueEvent(eventName,eventDate,activityPoints)
                
    def buildAttendeeList(self):
        for file in os.listdir(ATTENDEE_EXPORTS):
            sheet = self.openAndValidateSheet(ATTENDEE_EXPORTS, file)
            if (sheet is None):
                continue
            print(f"Processing Registrant Export {file}...")
            for ndx in range(0,sheet.__len__()):
                if (sheet['Payment Status'].iloc[ndx] != 'Paid'):
                    # Skip records that are cancelled or pending
                    continue
                attendee = self.getUser(
                    firstName=sheet['First Name'].iloc[ndx],
                    lastName=sheet['Last Name'].iloc[ndx],
                    email=sheet['Email'].iloc[ndx])
                attendee.attended.append(str(sheet['Event'].iloc[0]))
                # diagnostic print
                # print(f"\t Attended by {attendee.firstName} {attendee.lastName} - {attendee.email}: total is {attendee.points} points")

    def assignPoints(self):
        # iterate over events, and assign points to every user that has that event
        for eventName, event in self.eventMap.items():
            for attendeeEmail, attendee in self.userMap.items():
                if eventName in attendee.attended:
                    attendee.points += event.activityPoints


if __name__ == "__main__":
    accountant = Accountant()
    accountant.buildEventList()
    accountant.buildAttendeeList()
    accountant.printEvents()
    accountant.assignPoints()
    accountant.printAttendees()
