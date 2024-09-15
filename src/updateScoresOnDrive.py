import googleDriveClient as gd
import activityAccountant as aa
import os
import shutil


def downloadInputFiles(gdService, rootDirId, localInputDir):
    # Download the registrant subdir
    gd.downloadExcelDirectory(
        gdService,
        gd.getChildId(
            gdService,
            rootDirId,
            aa.REGISTRANT_SUBDIR,
        ),
        os.path.join(localInputDir, aa.REGISTRANT_SUBDIR),
    )
    # Download the event subdir
    gd.downloadExcelDirectory(
        gdService,
        gd.getChildId(
            gdService,
            rootDirId,
            aa.EVENT_SUBDIR,
        ),
        os.path.join(localInputDir, aa.EVENT_SUBDIR),
    )
    # Download the aliases file
    gd.downloadExcel(
        gdService,
        fileId=gd.getChildId(gdService, rootDirId, aa.EMAIL_ALIAS_FILE),
        fileName=aa.EMAIL_ALIAS_FILE,
        destDir=localInputDir,
    )


if __name__ == "__main__":
    # Create the google service and download all hte files we need
    gdService = gd.createService()
    localInputDir = "/tmp/activityAccountant/input"
    if os.path.isdir(localInputDir):
        shutil.rmtree(localInputDir)
    rootDirId = gd.getFolderIdByName(gdService, "ActivityAccounting")
    # Download input files
    downloadInputFiles(gdService, rootDirId, localInputDir)
    # Delete the old output dir
    localOutputDir = "/tmp/activityAccountant/results"
    if os.path.isdir(localOutputDir):
        shutil.rmtree(localOutputDir)
    # Create the accountant - this is where the magic happens
    accountant = aa.Accountant(localInputDir, localOutputDir)
    # Export the results
    resultFilePath = accountant.exportResults()
    # Upload the results to drive
    gd.uploadSpreadsheet(
        gdService,
        gd.getFolderIdByName(gdService, "scoring"),
        resultFilePath,
    )
