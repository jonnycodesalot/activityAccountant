# activityAccountant

ActivityAccountant tracks activity points earned by members of a club or volunteer organization. Scripts are written to ingest spreadsheets exported by Joomla or another CCM, and create an output spreadsheet that totals the activity points earned by registrant based on what events they attended.

# Directory Structure

The script assumes the following directory structure.

* `<root input dir>`/
  * `eventExports/`
    * `eventExport1.xlsx`
    * `eventExport2.xlsx`
    * `...`
  * `registrantExports/`
    * `registrantExport1.xlsx`
    * `registrantExport2.xlsx`
    * `...`
  * `emailAliases.xlsx`
* `<root output dir>`
  * `scores/`

## Input Files

All spreadsheets must be *Excel* spreadsheets with an `.xlsx` extension. This is because `xlsx` is the format exported by the Joomla Events plugin, which is the target use case for this script. 

When downloading, the script will ignore all content that isn't an excel spreadsheet or a folder, so you can put readmes, instructions, etc. in there no problem.

### eventExports/

This directory contains files describing the events to be tabulated. The file names are ignored. The following columns are required:

* `id` - the numeric ID of the event. This must be unique across all events.
* `title` - the name of the event
* `event_date` - the event start date
* `event_end_date` - When the event ends. Events older than `MAXIMUM_EVENT_AGE`(currently 3 years) when the script runs will be ignored. 
* `activity_points` - the number of points to be awarded to a registrant for the event. If this field is empty or 0, the event will be ignored.

If an event occurs more than once (using the id as key) across all contents of eventExports, then the details of the last read one will prevail. The code isn't currently written to control the order in which files are read though, so it's best to avoid duplicates.

### registrantExports/

This directory contains files describing the registrants who attended events. The file names are ignored. The following columns are required:

* `User ID` - A number representing the registrant's user ID. This would ideally be our key, but if you have events with non-members (who have no id), this field is 0, and that doesn't work. The script will coalesce any records with the same User ID to represent the same registrant.
* `First Name` - Obvious.
* `Last Name` - Obvious. First and Last name *can* be used to coalesce records, but since a simple compare between two copies of the same name taken at different times is often fraught (Jon vs. Jonathan, OConnor vs. O'Connor), this is unlikely to be very helpful.
* `Email` - The registrant's email. This is used as the primary key, and multiple records with the same email will be coalesced into one registrant. However, some people register with different emails at different times; this problem can be solved with the `emailAliases.xlsx` file.
* `Event ID` - The event that this record registers the registrant for
* `Payment Status` - All records with a value other than `Paid` are ignored. This field in Joomla is used to track registration status like cancellations, or registrant records whose payment processing was never finished. So this is important to filter out irrelevant records.

There is also an optional field:
* `multiplier` - If this exists for a sheet, and is filled in for a particular record with a numeric value, that value will be multiplied by the `activity_points` column in the corresponding event when computing the total activity points for the registrant. This allows you to easily allocate different numbers of points for different registrants at the same event.

Note that where records are coalesced (due to email, ID, or name), the fields from the latest available record, as judged by the corresponding event's date, will be kept. This allows you to ensure that the output describes the registrant by their most recent registration.

### emailAliases.xlsx

This works around the problem that some registrations for the same registrant may be different email addresses, different exact spellings of their names, and may lack the User ID. Each column is for a single registrant, though the top row is ignored (put the registrant's name for documentation purposes). Every other cell in the column is taken to be an email address for that registrant. When enumerating registrations, the script will coalesce all the registrations for the emails in a column to belong to the same registrant.

## Output Files

There is only one output file, scoring.xlsx, meant to help a registrant for a new event understand where they stand based on past activity. This includes the following columns:
* `User ID` - The ID of the registrant. This may be zero if no registrant records had it recorded for them.
* `First Name`
* `Last Name`
* `Email` (omitted by default) - The latest email from all the coalesced records for the registrant
* `ActivityPoints` - The total points allocated to the registrants, from all registrations taken as input, including any multipliers
* `ActivityRank` - The rank of the registrant, according to how many points they have. When registering for an event with limited capacity, this is actually more useful to the registrant than the point count; it shows how many people are ahead of you in line. If multiple registrants have the same number of points, they will all share the rank of the first person with that many points.
* `SameRankCount` - How many users have the same number of points as a given user. So if your ActivityRank + SameRankCount is 10, then you are at worst 10th in line for any registration.
* `<Event Names>` - For every event recorded, a column is added, and each row will be marked to show how many points, if any, a given registrant earned in that column. This makes it easy for a registrant to audit the output and convince themselves the calculations are correct, or diagnose problems.

# GitHub Actions and Google Drive

There are some GitHub workflows defined to allow execution of the tracker without setting up a local environment.

**Note** that hwen using Google Drive, the parent folder of all directories is taken to be `ActivityAccountant`, and this search is performed by name among all folders shared with the relevant user. So if you have multiple folders by this name available, things will bbreak.

## GitHub Authentication

Set up a secret in your GitHub repo/fork called `GOOGLE_APPLICATION_CREDENTIALS`, which contains the Google Drive credentials from a generated credentials.json file. See the web for instructions on how to obtain this.

It's recommended to use a service account rather than an individual credentials, for security reasons. In either case, you can share the root folder with the relevant account the way you would share any other Google Drive document.

## Workflow - Update Activity Scores

Downloads all the files from Drive, computes the new scores, and uploads the results