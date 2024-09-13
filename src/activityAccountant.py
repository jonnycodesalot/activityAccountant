import pandas as pd
import os
import math
import googleDriveClient as gd
import datetime as dt
import xlsxwriter
import shutil

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
    def __init__(self, firstName, lastName, email, memberId):
        self.firstName = firstName
        self.lastName = lastName
        self.email = email
        self.points = 0
        self.id = memberId
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

    def getUserFromName(self, firstName, lastName):
        for email, user in self.userMap.items():
            if (
                user.firstName.strip().lower() == firstName.strip().lower()
                and user.lastName.strip().lower() == lastName.strip().lower()
            ):
                return email
        return None

    def getUser(self, firstName, lastName, email, memberId):
        email = email.strip().lower()
        if not self.userMap.__contains__(email):
            # Try searching by name (may have changed email)
            emailInList = self.getUserFromName(firstName, lastName)
            if emailInList is not None:
                email = emailInList
            else:
                self.userMap[email] = Attendee(
                    firstName.strip(), lastName.strip(), email, memberId
                )
        toReturn = self.userMap[email]
        if toReturn.id == 0:
            toReturn.id = memberId
        return toReturn

    def addUniqueEvent(self, id, name, date, pointCount):
        name = name.strip()
        if not self.eventMap.__contains__(id):
            self.eventMap[id] = Event(int(id), str(name), str(date), int(pointCount))
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
                if int(sheet["Event ID"].iloc[ndx]) not in self.eventMap:
                    # if the event for this registrant record isn't in our
                    # list, ignore it.
                    continue
                attendee = self.getUser(
                    firstName=str(sheet["First Name"].iloc[ndx]),
                    lastName=str(sheet["Last Name"].iloc[ndx]),
                    email=str(sheet["Email"].iloc[ndx]),
                    memberId=int(sheet["User ID"].iloc[ndx]),
                )
                # Note that we don't filter out what users to include based
                # on any event information here. If we've ever processed a
                # record for them, we want to make sure we continue having
                # a record in the output, even if it's always 0 points.
                # Otherwise, we risk leaving old scores around for people
                # who haven't earned in a very long time.
                attendee.addEvent(int(sheet["Event ID"].iloc[ndx]))

    def assignPoints(self):
        # iterate over events, and assign points to every user that has that event
        for eventId, event in self.eventMap.items():
            eventId = int(eventId)
            for attendeeEmail, attendee in self.userMap.items():
                if eventId in attendee.attended:
                    attendee.points += event.activityPoints

    def exportResults(self):
        userIds = list()
        firstNames = list()
        lastNames = list()
        emails = list()
        points = list()
        ranks = list()
        sameRankCount = list()
        inputCols = {
            "User ID": userIds,
            "First Name": firstNames,
            "Last Name": lastNames,
            "Email": emails,
            "ActivityPoints": points,
            "ActivityRank": ranks,
            "SameRankCount": sameRankCount,
        }
        sortedEvents = sorted(
            self.eventMap.items(), key=lambda event: event[1].date, reverse=True
        )
        for event in sortedEvents:
            if event[1].name in inputCols:
                raise Exception(
                    f"There appears to multiple events with the title {event.name}. This is not supportd."
                )
            inputCols[event[1].name] = list()
        sortedUsers = sorted(
            self.userMap.items(), key=lambda attendee: attendee[1].points, reverse=True
        )
        rank = 1
        numberWithSameRank = 0
        lastScoreExamined = None
        for attendee in sortedUsers:
            if lastScoreExamined is None:
                lastScoreExamined = attendee[1].points
            userIds.append(attendee[1].id)
            firstNames.append(attendee[1].firstName)
            lastNames.append(attendee[1].lastName)
            emails.append(attendee[1].email)
            points.append(attendee[1].points)
            for eventId, event in self.eventMap.items():
                if int(eventId) in attendee[1].attended:
                    inputCols[event.name].append(event.activityPoints)
                else:
                    inputCols[event.name].append(" ")
            if attendee[1].points == lastScoreExamined:
                numberWithSameRank += 1
            else:
                # record how many people have the previous rank, for every row
                for it in range(0, numberWithSameRank):
                    sameRankCount.append(numberWithSameRank)
                # Update to the new rank
                rank += numberWithSameRank
                # Reset the number of people with the same rank
                numberWithSameRank = 1
            lastScoreExamined = attendee[1].points
            ranks.append(rank)
        for it in range(sameRankCount.__len__(), firstNames.__len__()):
            sameRankCount.append(numberWithSameRank)
        dataFrame = pd.DataFrame(inputCols)
        os.makedirs(self.outputBaseDir, exist_ok=True)
        resultFilePath = os.path.join(
            self.outputBaseDir,
            dt.datetime.now().strftime("%Y-%m-%d-%H:%M:%S_") + SCORE_FILE_SUFFIX,
        )
        # export to excel, freezing the top row
        writer = pd.ExcelWriter(resultFilePath)
        dataFrame.to_excel(
            writer, sheet_name="scores", index=False, freeze_panes=(1, 0), na_rep="NaN"
        )
        # Set column widths to bring us joy
        for column in dataFrame:
            column_length = max(
                dataFrame[column].astype(str).map(len).max(), len(column)
            )
            col_idx = dataFrame.columns.get_loc(column)
            writer.sheets["scores"].set_column(col_idx, col_idx, column_length)
        writer.close()
        # Return the path to the file
        return resultFilePath


if __name__ == "__main__":
    gdService = gd.createService()
    localInputDir = "/tmp/activityAccountant/input"
    # if os.path.isdir(localInputDir):
    #     shutil.rmtree(localInputDir)
    # gd.downloadDirectory(
    #     gdService,
    #     gd.getFolderIdByName(gdService, "ActivityAccounting"),
    #     localInputDir,
    # )
    localOutputDir = "/tmp/activityAccountant/results"
    if os.path.isdir(localOutputDir):
        shutil.rmtree(localOutputDir)
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
