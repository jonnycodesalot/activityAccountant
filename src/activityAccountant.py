import pandas as pd
import os
import math
import datetime as dt
import xlsxwriter
import shutil
import pytz

REGISTRANT_SUBDIR = "registrantExports"
EVENT_SUBDIR = "eventExports"
EMAIL_ALIAS_FILE = "emailAliases.xlsx"

# We don't count events whose dates are older than a
# certain amount (based on the end date)
MAXIMUM_EVENT_AGE = pd.DateOffset(years=3)
OLDEST_REGISTRANT_ALLOWED = dt.datetime(
    year=2023, month=9, day=1, hour=9, minute=0, second=0
)
# Overwritten by "now" in code if earlier than now.
LATEST_EVENT_END_DATE = pd.to_datetime("2025/03/17")


class Event:
    def __init__(self, id, name, date, endDate, activityPoints):
        self.name = name
        self.date = date
        self.endDate = endDate
        self.activityPoints = activityPoints
        self.id = int(id)

    def __str__(self):
        return f"{self.name} {self.date} - {self.activityPoints} points\n"


class Registrant:
    def __init__(self, firstName, lastName, email, memberId, sourceEventDate):
        self.firstName = firstName
        self.lastName = lastName
        self.email = email
        self.points = 0
        self.id = memberId
        self.eventMultipliers = dict()
        self.sourceEventDate = sourceEventDate

    def __str__(self):
        str = (
            f"{self.firstName} {self.lastName} - {self.points} points - {self.email}\n"
        )
        for eventId in self.eventMultipliers:
            str += f"\tEvent {eventId}"
        return str

    def addEvent(self, eventId, multiplier=1):
        # Never allow the same event to be recorded twice for a registrant;
        # They may have the same event registered in multiple rows in our
        # input, due to clerical error. Let's be smart enough to ignore it
        if int(eventId) not in self.eventMultipliers:
            self.eventMultipliers[int(eventId)] = multiplier


class Accountant:
    def __init__(self, inputDir, outputDir):
        self.userMap = dict()
        self.eventMap = dict()
        self.inputBaseDir = inputDir
        self.outputBaseDir = outputDir
        self.aliases = dict()
        self.loadEmailAliases()
        self.buildEventList()
        self.buildAttendeeList()
        self.eliminateOutdatedRegistrants()
        self.assignPoints()

    def eliminateOutdatedRegistrants(self):
        # if your latest registration is older than OLDEST_REGISTRANT_ALLOWED,
        # then your record is thrown out.
        toDelete = list()
        for email, member in self.userMap.items():
            if pd.to_datetime(member.sourceEventDate) < pd.to_datetime(
                OLDEST_REGISTRANT_ALLOWED
            ):
                toDelete.append(email)
        for email in toDelete:
            self.userMap.__delitem__(email)

    def loadEmailAliases(self):
        # We support the use of an aliases file to deal with the fact that many
        # users may change emails from event to event, throwing off the keying
        # mechanism. The aliases file is an excel spreadsheet, with one column
        # added for every user that has multiple emails. The title row is ignored
        # (put the person's name, for documentation), but every cell in that colum
        # is taken to be a different email for the same person.
        #
        # This function reads the file and builds a dictionary of list to look up
        # aliases as it builds the user list.
        aliasData = self.openAndValidateSheet(self.inputBaseDir, EMAIL_ALIAS_FILE)
        for column in aliasData.columns:
            aliasList = list()
            [
                aliasList.append(str(alias).strip().lower())
                for alias in aliasData[column]
            ]
            for alias in aliasList:
                reducedList = aliasList.copy()
                reducedList.remove(alias)
                self.aliases[alias] = reducedList

    def getUserFromEmailOrAlias(self, email):
        # Given a user's email, will return the email that is a key into that
        # user's record (which may be the same email, or an alias). If no such
        # user has been recorded, returns None.
        if self.userMap.__contains__(email):
            return email
        elif email in self.aliases:
            aliasList = self.aliases[email]
            for alias in aliasList:
                if alias in self.userMap:
                    return alias
        return None

    def getUserFromName(self, firstName, lastName):
        # Matches record by name, using a basically exact compare. Unlikely
        # to be useful very often.
        for email, user in self.userMap.items():
            if (
                user.firstName.strip().lower() == firstName.strip().lower()
                and user.lastName.strip().lower() == lastName.strip().lower()
            ):
                return email
        return None

    def getUserFromId(self, memberId):
        # Gets the user based on their member ID. This is the ideal case,
        # but we don't use it exclusively because new members will generally
        # not have an ID associated with their first registration.
        if memberId == 0:
            # Can't map based on zero. Return nothing.
            return None
        for email, user in self.userMap.items():
            if user.id == memberId:
                return email
        return None

    def getCreateOrUpdateUser(
        self, firstName, lastName, email, memberId, eventRecordDate
    ):
        email = email.strip().lower()
        existingEmail = self.getUserFromEmailOrAlias(email)
        if existingEmail is None:
            # try searching by ID
            existingEmail = self.getUserFromId(memberId)
            if existingEmail is None:
                # Try searching by name (may have changed email)
                existingEmail = self.getUserFromName(firstName, lastName)
        if existingEmail is not None:
            # Check which to keep
            existing = self.userMap[existingEmail]
            # Check if there is a duplicate email (to the extent we can), and if so, fail out.
            if existing.id != 0 and memberId != 0 and memberId != existing.id:
                raise Exception(
                    f"It appears that two distinct members, {firstName} {lastName} "
                    + f"(ID {memberId}), and {existing.firstName} {existing.lastName} "
                    + f"(ID {existing.id}), are using the same email address. This is not allowed."
                )
            if existing.sourceEventDate < eventRecordDate:
                # This entry is newer. Leave the old ID - we'll check on that
                # below to make sure we eliminate the 0 record
                existing.firstName = firstName
                existing.lastName = lastName
                existing.email = email
                existing.sourceEventDate = eventRecordDate
                self.userMap.__delitem__(existingEmail)
                self.userMap[email] = existing
            # If the old one didn't have an ID, overwrite it with the new one
            if existing.id == 0:
                existing.id = memberId
            return existing
        else:
            # Make a new record
            newRecord = Registrant(
                firstName.strip(),
                lastName.strip(),
                email,
                memberId,
                eventRecordDate,
            )
            self.userMap[email] = newRecord
            return newRecord

    def addUniqueEvent(self, id, name, date, endDate, pointCount):
        name = name.strip()
        if not self.eventMap.__contains__(id):
            self.eventMap[id] = Event(
                int(id), str(name), str(date), str(endDate), int(pointCount)
            )
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
        spreadsheet = pd.ExcelFile(os.path.join(directory, file))
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
        # LATEST_EVENT_END_DATE can force us to let
        # more events in, but it can't force us to leave out events that
        #  have finished
        latestAllowedEventEndDate = LATEST_EVENT_END_DATE
        if latestAllowedEventEndDate is None or latestAllowedEventEndDate < currTime:
            latestAllowedEventEndDate = currTime
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
                if str(eventEndDate).startswith("0000"):
                    eventEndDate = sheet["event_date"].iloc[ndx]
                eventEndDate = pd.to_datetime(eventEndDate)
                if eventEndDate < (currTime - MAXIMUM_EVENT_AGE):
                    print(
                        f"***Event {eventName}'s end date is older than the maximum event age. It will not be counted."
                    )
                    continue
                if eventEndDate > (latestAllowedEventEndDate):
                    print(
                        f"***Event {eventName} ends after the latest allowed event end date. It will not be counted."
                    )
                    continue
                eventBeginDate = sheet["event_date"].iloc[ndx]
                activityPoints = activityPoints
                self.addUniqueEvent(
                    sheet["id"].iloc[ndx],
                    eventName,
                    eventBeginDate,
                    eventEndDate,
                    activityPoints,
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
                eventId = int(sheet["Event ID"].iloc[ndx])
                if eventId not in self.eventMap:
                    # if the event for this registrant record isn't in our
                    # list, ignore it.
                    continue
                memberId = sheet["User ID"].iloc[ndx]
                group = None
                if "Group: " in sheet:
                    group = sheet["Group: "].iloc[ndx]
                if group and str(group) != "" and str(group) != "nan":
                    # If they're part of a group signup, the memberId
                    # will be the person who signed everyone up. Omit
                    # the ID and match on the other attributes
                    memberId = 0
                else:
                    memberId = int(memberId)
                attendee = self.getCreateOrUpdateUser(
                    firstName=str(sheet["First Name"].iloc[ndx]),
                    lastName=str(sheet["Last Name"].iloc[ndx]),
                    email=str(sheet["Email"].iloc[ndx]),
                    memberId=memberId,
                    eventRecordDate=pd.Timestamp(self.eventMap[eventId].date),
                )
                # Cull any registrations that were actually NoShowed.
                if "Attendance" in sheet:
                    val = sheet["Attendance"].iloc[ndx]
                    if val and str(val) != "" and str(val) != "nan":
                        tempstr = str(val)
                        if str(val) == "NoShow":
                            # Skip this registration...they didn't show up, so they get no points.
                            continue
                        else:
                            raise Exception(
                                f"The 'Attendance' field has an "
                                + f"unexpected value in row {ndx}. It must be 'NoShow' or empty"
                            )
                # Apply a special multiplier if one exists. This is used for custom events,
                # to give a user variable number of points for a single-point event
                # (say, construction work)
                multiplier = 1
                if "multiplier" in sheet:
                    val = sheet["multiplier"].iloc[ndx]
                    if val and str(val) != "" and str(val) != "nan":
                        multiplier = int(val)
                # Mark the attendee for this event, with the specified multiplier
                attendee.addEvent(int(sheet["Event ID"].iloc[ndx]), multiplier)

    def assignPoints(self):
        # iterate over events, and assign points to every user that has that event
        for eventId, event in self.eventMap.items():
            eventId = int(eventId)
            for attendeeEmail, attendee in self.userMap.items():
                if eventId in attendee.eventMultipliers:
                    attendee.points += (
                        event.activityPoints * attendee.eventMultipliers[eventId]
                    )

    def exportResults(self, afterTimestampName, includeEmails=False):
        userIds = list()
        firstNames = list()
        lastNames = list()
        emails = list()
        points = list()
        ranks = list()
        sameRankCount = list()
        inputCols = {
            "User ID": userIds,
        }
        if includeEmails:
            inputCols["Email"] = emails
        inputCols["First Name"] = firstNames
        inputCols["Last Name"] = lastNames
        inputCols["ActivityPoints"] = points
        inputCols["ActivityRank"] = ranks
        inputCols["SameRankCount"] = sameRankCount
        sortedEvents = sorted(
            self.eventMap.items(), key=lambda event: event[1].endDate, reverse=True
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
            if includeEmails:
                emails.append(attendee[1].email)
            points.append(attendee[1].points)
            for eventId, event in self.eventMap.items():
                if int(eventId) in attendee[1].eventMultipliers:
                    inputCols[event.name].append(
                        event.activityPoints
                        * attendee[1].eventMultipliers[int(eventId)]
                    )
                else:
                    inputCols[event.name].append("")
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
            dt.datetime.now()
            .astimezone(pytz.timezone("America/New_York"))
            .strftime("%Y-%m-%d-%H:%M:%S_")
            + afterTimestampName
            + ".xlsx",
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
