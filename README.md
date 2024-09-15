# activityAccountant

ActivityAccountant tracks activity points earned by members of a club or volunteer organization. Scripts are written to ingest spreadsheets exported by Joomla or another CCM, and create an output spreadsheet that totals the activity points earned by user based on what events they attended.

## Directory Structure

The script assumes the following directory structure:

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

Note that the names of files in the `registrantExports` and `eventExports` directory do not matter - any number of spreadsheets will be accepted. 

All spreadsheets must be *Excel* spreadsheets with an `.xlsx` extension. This is because `xlsx` is the format exported by the Joomla Events plugin, which is the target use case for this script.

### eventExports/

This directory contains files describing the events to be tabulated. The following columns are required:

* `id` - the numeric ID of the event. This must be unique across all events.
* `title` - the name of the event
* `event_date` - the event start date
* `event_end_date` - When the event ends. Events older than `MAXIMUM_EVENT_AGE`(currently 3 years) when the script runs will be ignored. 
* `activity_points` - the number of points to be awarded to a registrant for the event. If this field is empty or 0, the event will be ignored.

If an event occurs more than once (using the id as key) across all contents of eventExports, then the details of the last read one will prevail. The code isn't currently written to control the order in which files are read though, so it's best to avoid duplicates.

### registrantExports/


