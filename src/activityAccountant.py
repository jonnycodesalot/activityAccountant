import pandas as pd
import os

#SOURCE_ROOT = 
ATTENDEE_EXPORT = "./test/events/"
OUTPUT_DIRECTORY = "./test/result/"
SCORE_FILE = "scoring.xlsx"
EVENT_FILE = "eventsList.xlsx"

class Event:
    def __init__(self,name,pointCount):
        self.name = name
        self.pointCount = pointCount


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

    # def populateEventList(self):
    #     if not os.path.exists(".xlsx"):
    #         # Skip it
    #         continue


    def getUser(self, firstName, lastName, email):
        if (not self.userMap.__contains__(email)):
            self.userMap[email] = Attendee(firstName,lastName,email)
        return self.userMap[email]
    
    def printAttendees(self):
        for key in self.userMap:
            print(self.userMap[key])


    def processEventFiles(self):
        for file in os.listdir(ATTENDEE_EXPORT):
            if not file.endswith(".xlsx"):
                # Skip it
                continue;
            print(f"Processing {file}...")
            spreadsheet = pd.ExcelFile(ATTENDEE_EXPORT + file)
            sheetCount = spreadsheet.sheet_names.__len__()
            if sheetCount != 1:
                raise Exception(f"File {file} has {sheetCount} sheets, but we only support single-sheet XLSX files.\n")
            sheet = pd.read_excel(spreadsheet,spreadsheet.sheet_names[0])
            eventName = sheet['Event'].iloc[0]
            eventDate = sheet['Event Date'].iloc[0]
            print(f"\tEvent Name: {eventName}")
            print(f"\tEvent Date: {eventDate}")
            for ndx in range(0,sheet.__len__()):
                if (sheet['Payment Status'].iloc[ndx] != 'Paid'):
                    print("skipping non-paid record...")
                    continue
                attendee = self.getUser(
                    firstName=sheet['First Name'].iloc[ndx],
                    lastName=sheet['Last Name'].iloc[ndx],
                    email=sheet['Email'].iloc[ndx])
                attendee.attended.append(eventName)
                print(f"\t Attended by {attendee.firstName} {attendee.lastName} - {attendee.email}: total is {attendee.points} points")

if __name__ == "__main__":
    accountant = Accountant()
    accountant.processEventFiles()
    accountant.printAttendees()
