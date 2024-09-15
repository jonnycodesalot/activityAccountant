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

Note that the names of files in the `registrantExports` and `eventExports` directory do not matter - any number of spreadsheets will be accepted. However, they must be *Excel* spreadsheets with an `.xlsx` extension. This is because this is the format exported by the Joomla Events plugin, which is the target use case.

